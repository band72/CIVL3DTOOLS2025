using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Exception = System.Exception;

namespace RCS.C3D2025.Tools
{
    public class LocalizeCommand
    {
        [CommandMethod("RCS_LOCALIZE")]
        public void LocalizeCogoPoints()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                PromptSelectionOptions pso = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect COGO points to transform (Press Enter to auto-transform Record Points 500-599): "
                };

                TypedValue[] tvs = { new TypedValue((int)DxfCode.Start, "AECC_COGO_POINT") };
                SelectionFilter filter = new SelectionFilter(tvs);

                PromptSelectionResult psr = ed.GetSelection(pso, filter);
                SelectionSet ss = null;

                if (psr.Status == PromptStatus.OK)
                    ss = psr.Value;
                else if (psr.Status == PromptStatus.Error)
                {
                    PromptSelectionResult psrAll = ed.SelectAll(filter);
                    if (psrAll.Status == PromptStatus.OK)
                        ss = psrAll.Value;
                }

                if (ss == null || ss.Count == 0)
                {
                    ed.WriteMessage("\nNo COGO points selected or found.");
                    return;
                }

                using (DocumentLock docLock = doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // ── 1. Discover all valid matched pairs (501→1, 502→2, …) ──────
                    var pairs = new List<(uint recordNum, uint fieldNum, Point3d recordLoc, Point3d fieldLoc)>();

                    for (uint rp = 501; rp <= 599; rp++)
                    {
                        uint fp = rp - 500;
                        CogoPoint recPt  = FindCogoPointByNumber(db, tr, rp);
                        CogoPoint fldPt  = FindCogoPointByNumber(db, tr, fp);
                        if (recPt != null && fldPt != null)
                            pairs.Add((rp, fp, recPt.Location, fldPt.Location));
                    }

                    if (pairs.Count < 2)
                    {
                        ed.WriteMessage("\nNot enough matching Record/Field point pairs found (e.g. 501→1, 502→2). Aborting.");
                        tr.Abort();
                        return;
                    }

                    ed.WriteMessage($"\n[RCS_LOCALIZE] Found {pairs.Count} matched pairs. Searching for best anchor/orient combo...");

                    // ── 2. Try every unique (anchor, orient) combination ───────────
                    int    bestAnchorIdx  = 0, bestOrientIdx = 1;
                    double bestMeanResidual = double.MaxValue;

                    // Store every combo result for full CSV report
                    var allComboResults = new List<(int ai, int oi, double meanRes, List<string> rows)>();

                    for (int ai = 0; ai < pairs.Count; ai++)
                    {
                        for (int oi = 0; oi < pairs.Count; oi++)
                        {
                            if (oi == ai) continue;

                            Matrix3d mat = BuildTransform(pairs[ai].recordLoc, pairs[oi].recordLoc,
                                                          pairs[ai].fieldLoc,  pairs[oi].fieldLoc);
                            if (mat == Matrix3d.Identity) continue;   // degenerate baseline

                            double sumRes = 0;
                            int    count  = 0;
                            var comboRows = new List<string>();

                            foreach (var p in pairs)
                            {
                                Point3d transformed = p.recordLoc.TransformBy(mat);
                                double dN = p.fieldLoc.Y - transformed.Y;
                                double dE = p.fieldLoc.X - transformed.X;
                                double res = Math.Sqrt(dN * dN + dE * dE);
                                sumRes += res;
                                count++;
                                comboRows.Add($"{p.recordNum},{p.fieldNum},{dN:F4},{dE:F4},{res:F4}");
                            }

                            double mean = count > 0 ? sumRes / count : double.MaxValue;
                            allComboResults.Add((ai, oi, mean, comboRows));

                            if (mean < bestMeanResidual)
                            {
                                bestMeanResidual = mean;
                                bestAnchorIdx    = ai;
                                bestOrientIdx    = oi;
                            }
                        }
                    }

                    // Sort all combos best → worst
                    allComboResults.Sort((a, b) => a.meanRes.CompareTo(b.meanRes));

                    var anchorPair = pairs[bestAnchorIdx];
                    var orientPair = pairs[bestOrientIdx];

                    ed.WriteMessage($"\n[RCS_LOCALIZE] Best Anchor : Record {anchorPair.recordNum} → Field {anchorPair.fieldNum}");
                    ed.WriteMessage($"\n[RCS_LOCALIZE] Best Orient : Record {orientPair.recordNum} → Field {orientPair.fieldNum}");
                    ed.WriteMessage($"\n[RCS_LOCALIZE] Expected Mean Residual: {bestMeanResidual:F4} ft");

                    // ── 3. Build the final transform ──────────────────────────────
                    Matrix3d finalMat = BuildTransform(anchorPair.recordLoc, orientPair.recordLoc,
                                                       anchorPair.fieldLoc,  orientPair.fieldLoc);

                    Vector3d translation = new Vector3d(
                        anchorPair.fieldLoc.X - anchorPair.recordLoc.X,
                        anchorPair.fieldLoc.Y - anchorPair.recordLoc.Y, 0);

                    Vector3d vecRecord = new Vector3d(orientPair.recordLoc.X - anchorPair.recordLoc.X,
                                                      orientPair.recordLoc.Y - anchorPair.recordLoc.Y, 0);
                    Vector3d vecField  = new Vector3d(orientPair.fieldLoc.X  - anchorPair.fieldLoc.X,
                                                      orientPair.fieldLoc.Y  - anchorPair.fieldLoc.Y,  0);
                    double angleRad    = Math.Atan2(vecField.Y, vecField.X) - Math.Atan2(vecRecord.Y, vecRecord.X);

                    // ── 4. Apply transform to matching record points ──────────────
                    int transformedCount = 0, failures = 0;
                    var residualLines = new List<string>
                    {
                        "RecordPoint,FieldPoint,Delta_Northing,Delta_Easting,Horizontal_Residual"
                    };

                    foreach (SelectedObject selObj in ss)
                    {
                        try
                        {
                            CogoPoint pt = tr.GetObject(selObj.ObjectId, OpenMode.ForWrite) as CogoPoint;
                            if (pt == null) continue;

                            // Only transform record points 500-599
                            if (pt.PointNumber < 500 || pt.PointNumber > 599) continue;

                            Point3d newLoc = pt.Location.TransformBy(finalMat);
                            pt.Easting  = newLoc.X;
                            pt.Northing = newLoc.Y;
                            transformedCount++;

                            CogoPoint fieldMatch = FindCogoPointByNumber(db, tr, pt.PointNumber - 500);
                            if (fieldMatch != null)
                            {
                                double dN  = fieldMatch.Northing - pt.Northing;
                                double dE  = fieldMatch.Easting  - pt.Easting;
                                double res = Math.Sqrt(dN * dN + dE * dE);
                                residualLines.Add($"{pt.PointNumber},{fieldMatch.PointNumber},{dN:F4},{dE:F4},{res:F4}");
                            }
                        }
                        catch (Exception innerEx)
                        {
                            failures++;
                            ed.WriteMessage($"\nWarning: Could not process point (Handle: {selObj.ObjectId.Handle}). {innerEx.Message}");
                        }
                    }

                    tr.Commit();

                    ed.WriteMessage($"\nLocalization complete. Transformed {transformedCount} points.");
                    if (failures > 0) ed.WriteMessage($"\nSkipped/failed: {failures} points.");
                    ed.WriteMessage($"\nTranslation: dE={translation.X:F4}, dN={translation.Y:F4}");
                    ed.WriteMessage($"\nRotation: {angleRad * 180.0 / Math.PI:F6} degrees");

                    // ── 5. Write full all-combos residual CSV ────────────────────
                    try
                    {
                        string dwgPath  = doc.Name;
                        string fallback = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        string outDir   = (!string.IsNullOrEmpty(dwgPath) && Path.IsPathRooted(dwgPath))
                                           ? (Path.GetDirectoryName(dwgPath) ?? fallback) : fallback;
                        string dwgName  = (!string.IsNullOrEmpty(dwgPath)) ? Path.GetFileNameWithoutExtension(dwgPath) : "Drawing";
                        string outFile  = Path.Combine(outDir, $"{dwgName}_Localization_Residuals_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

                        var allLines = new List<string>();

                        // ── Best-result applied section (top) ──
                        allLines.Add("=== APPLIED TRANSFORMATION (BEST FIT) ===");
                        allLines.Add($"Anchor: Record {anchorPair.recordNum} -> Field {anchorPair.fieldNum},Orient: Record {orientPair.recordNum} -> Field {orientPair.fieldNum},Mean Residual: {bestMeanResidual:F4}");
                        allLines.Add("RecordPoint,FieldPoint,Delta_Northing,Delta_Easting,Horizontal_Residual");
                        allLines.AddRange(residualLines.GetRange(1, residualLines.Count - 1));
                        allLines.Add("");

                        // ── All combos section ──
                        allLines.Add("=== ALL ANCHOR/ORIENT COMBINATIONS (sorted best to worst) ===");
                        allLines.Add("");

                        int rank = 1;
                        foreach (var combo in allComboResults)
                        {
                            var ap = pairs[combo.ai];
                            var op = pairs[combo.oi];
                            bool isBest = combo.ai == bestAnchorIdx && combo.oi == bestOrientIdx;
                            allLines.Add($"Rank {rank},{(isBest ? "*** BEST FIT ***" : "")}");
                            allLines.Add($"Anchor: Record {ap.recordNum} -> Field {ap.fieldNum},Orient: Record {op.recordNum} -> Field {op.fieldNum},Mean Residual: {combo.meanRes:F4}");
                            allLines.Add("RecordPoint,FieldPoint,Delta_Northing,Delta_Easting,Horizontal_Residual");
                            allLines.AddRange(combo.rows);
                            allLines.Add("");
                            rank++;
                        }

                        File.WriteAllLines(outFile, allLines);
                        ed.WriteMessage($"\nFull Combo Report ({allComboResults.Count} combinations): {outFile}");
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nError in RCS_LOCALIZE: {ex.Message}");
            }
        }

        /// <summary>
        /// Builds a Translate-then-Rotate matrix from two record/field point pairs.
        /// Returns Matrix3d.Identity if either baseline is degenerate (< 1mm).
        /// </summary>
        private static Matrix3d BuildTransform(
            Point3d recordAnchor, Point3d recordOrient,
            Point3d fieldAnchor,  Point3d fieldOrient)
        {
            Vector3d vecRec = new Vector3d(recordOrient.X - recordAnchor.X, recordOrient.Y - recordAnchor.Y, 0);
            Vector3d vecFld = new Vector3d(fieldOrient.X  - fieldAnchor.X,  fieldOrient.Y  - fieldAnchor.Y,  0);

            if (vecRec.Length < 0.001 || vecFld.Length < 0.001)
                return Matrix3d.Identity;

            double angle = Math.Atan2(vecFld.Y, vecFld.X) - Math.Atan2(vecRec.Y, vecRec.X);

            Vector3d translation = new Vector3d(
                fieldAnchor.X - recordAnchor.X,
                fieldAnchor.Y - recordAnchor.Y, 0);

            Matrix3d transMat = Matrix3d.Displacement(translation);
            Matrix3d rotMat   = Matrix3d.Rotation(angle, Vector3d.ZAxis,
                                    new Point3d(fieldAnchor.X, fieldAnchor.Y, 0));
            return rotMat * transMat;
        }

        private CogoPoint FindCogoPointByNumber(Database db, Transaction tr, uint pointNum)
        {
            CogoPointCollection cogoPoints = CogoPointCollection.GetCogoPoints(db);
            try
            {
                ObjectId ptId = cogoPoints.GetPointByPointNumber(pointNum);
                if (ptId != ObjectId.Null)
                    return tr.GetObject(ptId, OpenMode.ForRead) as CogoPoint;
            }
            catch
            {
                foreach (ObjectId id in cogoPoints)
                {
                    CogoPoint pt = tr.GetObject(id, OpenMode.ForRead) as CogoPoint;
                    if (pt != null && pt.PointNumber == pointNum) return pt;
                }
            }
            return null;
        }
    }
}
