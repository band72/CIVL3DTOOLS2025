using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using System;
using System.IO;
using System.Text;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025.Tools
{
    /// <summary>
    /// RCS_EXPORT_COGO_POINTS
    /// Exports all (or selected) Civil 3D COGO points to a CSV file.
    /// Format: PointNo,Northing,Easting,Elevation,RawDescription,FullDescription
    /// </summary>
    public class ExportCogoPointsCommand
    {
        [CommandMethod("RCS_EXPORT_COGO_POINTS")]
        public void RunExportCogoPoints()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed  = doc.Editor;
            var db  = doc.Database;

            try
            {
                // ── 1. Selection ────────────────────────────────────────────────
                TypedValue[]    filterValues = { new TypedValue((int)DxfCode.Start, "AECC_COGO_POINT") };
                SelectionFilter filter       = new SelectionFilter(filterValues);

                var pso = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect COGO points to export (Enter = ALL): ",
                    AllowDuplicates  = false
                };

                PromptSelectionResult psr = ed.GetSelection(pso, filter);

                SelectionSet ss;
                if (psr.Status == PromptStatus.OK)
                {
                    ss = psr.Value;
                }
                else if (psr.Status == PromptStatus.Error)
                {
                    // User pressed Enter without making a selection → export ALL
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
                    // Cancelled
                    return;
                }

                if (ss == null || ss.Count == 0)
                {
                    ed.WriteMessage("\nNo COGO points selected.");
                    return;
                }

                // ── 2. Output file path ─────────────────────────────────────────
                string defaultPath = GetDefaultExportPath(db);

                // Prompt user for save path (using AutoCAD string prompt — no WinForms dependency)
                var pathOpt = new PromptStringOptions($"\nOutput CSV path <{defaultPath}>: ")
                {
                    AllowSpaces     = true,
                    DefaultValue    = defaultPath,
                    UseDefaultValue = true
                };

                PromptResult pathResult = ed.GetString(pathOpt);
                if (pathResult.Status != PromptStatus.OK) return;

                string outPath = pathResult.StringResult?.Trim();
                if (string.IsNullOrWhiteSpace(outPath)) outPath = defaultPath;

                // Make sure the parent directory exists
                string dir = Path.GetDirectoryName(outPath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                // ── 3. Export ───────────────────────────────────────────────────
                int exported = 0, skipped = 0;

                var sb = new StringBuilder();
                sb.AppendLine("PointNo,Northing,Easting,Elevation,RawDescription,FullDescription");

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in ss)
                    {
                        if (selObj == null) continue;

                        try
                        {
                            var pt = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as CogoPoint;
                            if (pt == null) { skipped++; continue; }

                            // Point number, coordinates
                            uint   ptNo   = pt.PointNumber;
                            double north  = pt.Northing;
                            double east   = pt.Easting;
                            double elev   = pt.Elevation;
                            string rawDesc  = EscapeCsv(pt.RawDescription    ?? "");
                            string fullDesc = EscapeCsv(pt.DescriptionFormat  ?? "");

                            sb.AppendLine(
                                $"{ptNo}," +
                                $"{north:F4}," +
                                $"{east:F4},"  +
                                $"{elev:F4},"  +
                                $"{rawDesc},"  +
                                $"{fullDesc}");

                            exported++;
                        }
                        catch (System.Exception ex)
                        {
                            skipped++;
                            ed.WriteMessage($"\n  [WARN] Could not read point: {ex.Message}");
                        }
                    }

                    tr.Commit();
                }

                File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8);

                ed.WriteMessage(
                    $"\nExport complete. " +
                    $"Exported={exported}  Skipped={skipped}" +
                    $"\nFile: {outPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nRCS_EXPORT_COGO_POINTS error: {ex.Message}");
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Determines a sensible default export path based on the current drawing.
        /// Falls back to C:\temp\ if the drawing has not been saved.
        /// </summary>
        private static string GetDefaultExportPath(Database db)
        {
            try
            {
                string dwgPath = db.Filename;
                if (!string.IsNullOrWhiteSpace(dwgPath) && File.Exists(dwgPath))
                {
                    string folder = Path.GetDirectoryName(dwgPath);
                    string dwgName = Path.GetFileNameWithoutExtension(dwgPath);
                    return Path.Combine(folder, $"{dwgName}_CogoPoints.csv");
                }
            }
            catch { }

            return @"C:\temp\rcs_cogo_points_export.csv";
        }

        /// <summary>
        /// Wraps a value in quotes and escapes internal quotes for RFC-4180 CSV.
        /// </summary>
        private static string EscapeCsv(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            // If the value contains commas, quotes, or newlines — wrap in quotes
            if (value.IndexOf(',')  >= 0 ||
                value.IndexOf('"')  >= 0 ||
                value.IndexOf('\n') >= 0 ||
                value.IndexOf('\r') >= 0)
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }
    }
}
