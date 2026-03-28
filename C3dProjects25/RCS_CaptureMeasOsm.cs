using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025.Tools
{
    /// <summary>
    /// RCS_CAPTURE_MEAS_OSM
    /// Captures Original Survey Measurements (OSM) from selected COGO points and exports
    /// a tab-delimited field capture report: Point#, RawDesc, Northing, Easting, Elevation.
    /// Output is written to the same directory as the current drawing.
    /// </summary>
    public class CaptureMeasOsmCommand
    {
        [CommandMethod("RCS_CAPTURE_MEAS_OSM")]
        public void CaptureMeasOsm()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // 1. Select COGO points (or Enter for ALL)
                PromptSelectionOptions pso = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect COGO points to capture (Enter = ALL): ",
                    AllowDuplicates  = false
                };

                TypedValue[]    filterValues = { new TypedValue((int)DxfCode.Start, "AECC_COGO_POINT") };
                SelectionFilter filter       = new SelectionFilter(filterValues);

                PromptSelectionResult psr = ed.GetSelection(pso, filter);

                SelectionSet ss;
                if (psr.Status == PromptStatus.OK)
                {
                    ss = psr.Value;
                }
                else if (psr.Status == PromptStatus.None)
                {
                    // Enter pressed with no selection → capture ALL
                    PromptSelectionResult allPsr = ed.SelectAll(filter);
                    if (allPsr.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nNo COGO points found in drawing.");
                        return;
                    }
                    ss = allPsr.Value;
                }
                else
                {
                    return; // Cancelled
                }

                if (ss == null || ss.Count == 0)
                {
                    ed.WriteMessage("\nNo COGO points selected.");
                    return;
                }

                // 2. Resolve output path — same folder as the DWG, or Documents as fallback
                string dwgPath  = doc.Name;
                string outDir   = (!string.IsNullOrEmpty(dwgPath) && Path.IsPathRooted(dwgPath))
                                    ? Path.GetDirectoryName(dwgPath)
                                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                string dwgName  = (!string.IsNullOrEmpty(dwgPath))
                                    ? Path.GetFileNameWithoutExtension(dwgPath)
                                    : "Drawing";

                string outFile  = Path.Combine(outDir, $"{dwgName}_OSM_Capture_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

                // 3. Read points inside a transaction
                var rows = new List<string>
                {
                    "PointNo\tRawDescription\tNorthing\tEasting\tElevation"
                };

                int captured = 0;
                int skipped  = 0;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in ss)
                    {
                        if (selObj == null) continue;

                        CogoPoint pt = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as CogoPoint;
                        if (pt == null) { skipped++; continue; }

                        string rawDesc = pt.RawDescription ?? string.Empty;

                        rows.Add(string.Format("{0}\t{1}\t{2:F4}\t{3:F4}\t{4:F4}",
                            pt.PointNumber,
                            rawDesc.Replace("\t", " "),
                            pt.Northing,
                            pt.Easting,
                            pt.Elevation));

                        captured++;
                    }

                    tr.Commit();
                }

                // 4. Write report
                File.WriteAllLines(outFile, rows, Encoding.UTF8);

                ed.WriteMessage($"\nOSM Capture complete.");
                ed.WriteMessage($"\n  Points captured : {captured}");
                if (skipped > 0)
                    ed.WriteMessage($"\n  Skipped (non-COGO): {skipped}");
                ed.WriteMessage($"\n  Output file     : {outFile}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in RCS_CAPTURE_MEAS_OSM: {ex.Message}");
            }
        }
    }
}
