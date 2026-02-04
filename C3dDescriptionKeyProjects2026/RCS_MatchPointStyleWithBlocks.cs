using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace RCS.C3D2025
{
    public class RCS_PointStyle_BlockMarkerMatch
    {
        private const string CsvApplyPath = @"C:\temp\PointStyle_BlockMarker_Match.csv";
        private const string CsvAuditPath = @"C:\temp\PointStyle_BlockMarker_Audit.csv";
        private const string LogPath = @"C:\temp\c3d_pointstyle_errors.log";

        // ✅ APPLY COMMAND
        [CommandMethod("RCS_MATCH_POINTSTYLE_BLOCK_MARKERS_NET")]
        public void Cmd_ApplyMatchPointStylesToBlocks()
        {
            RunApply(matchOnly: false);
        }

        // ✅ AUDIT / DRY RUN COMMAND (no changes)
        [CommandMethod("RCS_LIST_POINTSTYLES_AND_BLOCKS")]
        public void Cmd_AuditPointStylesVsBlocks()
        {
            RunApply(matchOnly: true);
        }

        // ---------------- Core Runner ----------------

        private void RunApply(bool matchOnly)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            Directory.CreateDirectory(@"C:\temp");
            Log($"=== START {(matchOnly ? "AUDIT" : "APPLY")} ===");

            if (doc == null)
            {
                Log("ERROR: ActiveDocument is null.");
                return;
            }

            var civDoc = CivilApplication.ActiveDocument;
            if (civDoc == null)
            {
                Log("ERROR: CivilApplication.ActiveDocument is null.");
                return;
            }

            Dictionary<string, string> blockMap;
            try
            {
                blockMap = BuildBlockNameMap(doc.Database);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Log("FATAL: Failed to build block map.", ex);
                return;
            }

            string csvPath = matchOnly ? CsvAuditPath : CsvApplyPath;

            int scanned = 0, hasMatch = 0, updated = 0, skipped = 0, failed = 0;

            using (var sw = new StreamWriter(csvPath, false))
            {
                sw.WriteLine("PointStyleName,MatchedBlockName,WouldUpdateOrUpdated,Status,Message");

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (ObjectId psId in civDoc.Styles.PointStyles)
                    {
                        scanned++;

                        try
                        {
                            var ps = tr.GetObject(psId, matchOnly ? OpenMode.ForRead : OpenMode.ForWrite) as PointStyle;
                            if (ps == null)
                            {
                                failed++;
                                sw.WriteLine($",,,FAIL,{Csv("Null PointStyle object")}");
                                continue;
                            }

                            string psName = ps.Name ?? "";
                            string key = psName.ToUpperInvariant();

                            if (!blockMap.TryGetValue(key, out string blockName))
                            {
                                skipped++;
                                sw.WriteLine($"{Csv(psName)},,No,SKIP,{Csv("No matching block in this DWG")}");
                                continue;
                            }

                            hasMatch++;

                            if (matchOnly)
                            {
                                sw.WriteLine($"{Csv(psName)},{Csv(blockName)},Yes,OK,{Csv("Audit only (no changes)")}");
                                continue;
                            }

                            bool markerTypeOk = SetMarkerTypeUseSymbol(ps);
                            ps.MarkerSymbolName = blockName;

                            bool symbolOk = string.Equals(ps.MarkerSymbolName ?? "", blockName, StringComparison.OrdinalIgnoreCase);

                            if (markerTypeOk && symbolOk)
                            {
                                updated++;
                                sw.WriteLine($"{Csv(psName)},{Csv(blockName)},Yes,OK,");
                            }
                            else
                            {
                                sw.WriteLine($"{Csv(psName)},{Csv(blockName)},Yes,VERIFY,{Csv($"MarkerTypeOk={markerTypeOk}, SymbolOk={symbolOk}")}");
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception exStyle)
                        {
                            failed++;
                            Log($"ERROR processing style index {scanned}", exStyle);
                            sw.WriteLine($",,,FAIL,{Csv(exStyle.Message)}");
                        }
                    }

                    tr.Commit();
                }
            }

            ed?.WriteMessage(
                $"\n{(matchOnly ? "AUDIT" : "APPLY")} complete." +
                $"\nScanned: {scanned} | Matches: {hasMatch} | Updated: {updated} | Skipped: {skipped} | Failed: {failed}" +
                $"\nCSV: {csvPath}" +
                $"\nLog: {LogPath}\n"
            );

            Log($"SUMMARY: scanned={scanned}, matches={hasMatch}, updated={updated}, skipped={skipped}, failed={failed}");
            Log($"=== END {(matchOnly ? "AUDIT" : "APPLY")} ===");
        }

        // ---------------- Helpers ----------------

        private static Dictionary<string, string> BuildBlockNameMap(Database db)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId id in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    string name = btr.Name ?? "";

                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (name.StartsWith("*", StringComparison.Ordinal)) continue;
                    if (name.Equals(BlockTableRecord.ModelSpace, StringComparison.OrdinalIgnoreCase)) continue;
                    if (name.Equals(BlockTableRecord.PaperSpace, StringComparison.OrdinalIgnoreCase)) continue;

                    map[name.ToUpperInvariant()] = name;
                }

                tr.Commit();
            }

            return map;
        }

        private static bool SetMarkerTypeUseSymbol(PointStyle ps)
        {
            try
            {
                var prop = ps.GetType().GetProperty("MarkerType", BindingFlags.Public | BindingFlags.Instance);
                if (prop == null || !prop.CanWrite) return false;

                var t = prop.PropertyType;

                if (t.IsEnum)
                {
                    var names = Enum.GetNames(t);

                    var exact = names.FirstOrDefault(n => n.Equals("UseSymbolForMarker", StringComparison.OrdinalIgnoreCase));
                    if (exact != null)
                    {
                        prop.SetValue(ps, Enum.Parse(t, exact, true));
                        return true;
                    }

                    var heuristic = names.FirstOrDefault(n =>
                        n.IndexOf("Symbol", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        n.IndexOf("Marker", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (heuristic != null)
                    {
                        prop.SetValue(ps, Enum.Parse(t, heuristic, true));
                        return true;
                    }

                    if (Enum.IsDefined(t, 1))
                    {
                        prop.SetValue(ps, Enum.ToObject(t, 1));
                        return true;
                    }

                    return false;
                }

                if (t == typeof(int))
                {
                    prop.SetValue(ps, 1);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Log("ERROR setting MarkerType", ex);
                return false;
            }
        }

        private static void Log(string msg, Exception ex = null)
        {
            try
            {
                using (var sw = new StreamWriter(LogPath, true))
                {
                    sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {msg}");
                    if (ex != null)
                        sw.WriteLine($"  {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                }
            }
            catch { }
        }

        private static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}
