#nullable enable
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace C3dProjects25
{
    public class RcsImportPointStylesNet8
    {
        private const string LogPath = @"C:\temp\c3doutput.txt";
        private const string DefaultCsvPath = @"C:\temp\rcs_pointstyles_export.csv";

        public enum RcsViewType { Plan, Model, Profile, Section }

        [CommandMethod("RCS_IMPORT_POINTSTYLES_V4")]
        public static void RCS_IMPORT_POINTSTYLES_V4()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            RunCommandSafe("RCS_IMPORT_POINTSTYLES_V4", () =>
            {
                var pso = new PromptStringOptions($"\nCSV path to IMPORT PointStyles from <{DefaultCsvPath}>: ")
                {
                    AllowSpaces = true,
                    DefaultValue = DefaultCsvPath,
                    UseDefaultValue = true
                };
                var prPath = ed.GetString(pso);
                if (prPath.Status != PromptStatus.OK) return;
                string csvPath = prPath.StringResult;

                if (!File.Exists(csvPath))
                {
                    ed.WriteMessage($"\nError: File not found: {csvPath}");
                    return;
                }

                bool overwrite = PromptYesNo(ed, "If a PointStyle exists, delete & recreate it (overwrite)?", false);
                bool autoCreateLayers = PromptYesNo(ed, "Auto-create missing layers referenced by the CSV?", true);

                var rows = LoadCsv(csvPath);
                Info(ed, $"Loaded {rows.Count} rows from CSV.");

                ProcessImport(doc, rows, overwrite, autoCreateLayers);
            });
        }

        private static void ProcessImport(Document doc, List<CsvStyleRow> rows, bool overwrite, bool autoCreateLayers)
        {
            var db = doc.Database;
            var ed = doc.Editor;
            var civDoc = CivilApplication.ActiveDocument;

            using (doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var pointStyles = civDoc.Styles.PointStyles;
                int created = 0, unlayered = 0, skipped = 0, failed = 0;

                foreach (var row in rows)
                {
                    try
                    {
                        string safeName = SanitizeName(row.StyleName);
                        bool exists = pointStyles.Contains(safeName);

                        if (exists)
                        {
                            if (overwrite)
                            {
                                var oldId = pointStyles[safeName];
                                var oldStyle = tr.GetObject(oldId, OpenMode.ForWrite) as PointStyle;
                                try { oldStyle?.Erase(); }
                                catch { Warn(ed, $"Could not erase '{safeName}' (In Use). Updating properties instead."); }
                            }
                            else
                            {
                                skipped++;
                                continue;
                            }
                        }

                        ObjectId styleId = pointStyles.Contains(safeName)
                            ? pointStyles[safeName]
                            : pointStyles.Add(safeName);

                        if (autoCreateLayers)
                        {
                            var layersNeeded = new HashSet<string> {
                                row.PlanMarkerLayer, row.PlanLabelLayer,
                                row.ModelMarkerLayer, row.ModelLabelLayer,
                                row.ProfileMarkerLayer, row.ProfileLabelLayer,
                                row.SectionMarkerLayer, row.SectionLabelLayer
                            };
                            EnsureLayersExist(db, tr, layersNeeded);
                        }

                        var ps = tr.GetObject(styleId, OpenMode.ForWrite) as PointStyle;
                        if (ps == null) { failed++; continue; }

                        bool layerSuccess = true;
                        try
                        {
                            SetComponentLayer(ps, RcsViewType.Plan, PointDisplayStyleType.Marker, row.PlanMarkerLayer);
                            SetComponentLayer(ps, RcsViewType.Plan, PointDisplayStyleType.Label, row.PlanLabelLayer);
                            SetComponentLayer(ps, RcsViewType.Model, PointDisplayStyleType.Marker, row.ModelMarkerLayer);
                            SetComponentLayer(ps, RcsViewType.Model, PointDisplayStyleType.Label, row.ModelLabelLayer);
                            SetComponentLayer(ps, RcsViewType.Profile, PointDisplayStyleType.Marker, row.ProfileMarkerLayer);
                            SetComponentLayer(ps, RcsViewType.Section, PointDisplayStyleType.Marker, row.SectionMarkerLayer);

                            // --- FIXED BLOCK ASSIGNMENT ---
                            if (!string.IsNullOrWhiteSpace(row.MarkerSymbolName))
                            {
                                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                                if (bt.Has(row.MarkerSymbolName))
                                {
                                    // 1. Set MarkerType to "Use Block Symbol"
                                    //    NOTE: API Enum member is 'UseBlockSymbol', NOT 'UseAutoCADBlockSymbol'
                                    ps.MarkerType = PointMarkerDisplayType.UseSymbolForMarker;

                                    // 2. Assign the Block Name string
                                    //    Property is 'MarkerSymbolName', NOT 'MarkerBlockName'
                                    ps.MarkerSymbolName = row.MarkerSymbolName;
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            Warn(ed, $"Layer/Block assignment failed for '{safeName}': {ex.Message}");
                            layerSuccess = false;
                        }

                        if (layerSuccess) created++; else unlayered++;
                    }
                    catch (System.Exception ex)
                    {
                        failed++;
                        Warn(ed, $"Failed to process '{row.StyleName}': {ex.Message}");
                    }
                }
                tr.Commit();
                string msg = $"Import complete. Created={created} Unlayered={unlayered} Skipped={skipped} Failed={failed}";
                Info(ed, msg);
                ed.WriteMessage($"\nRCS_IMPORT_POINTSTYLES: {msg}\n");
            }
        }

        private static void SetComponentLayer(PointStyle ps, RcsViewType viewType, PointDisplayStyleType component, string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName) || layerName == "0") return;

            DisplayStyle? ds = null;
            try
            {
                switch (viewType)
                {
                    case RcsViewType.Plan:
                        ds = ps.GetDisplayStylePlan(component);
                        break;
                    case RcsViewType.Model:
                        ds = ps.GetDisplayStyleModel(component);
                        break;
                    case RcsViewType.Profile:
                        if (component == PointDisplayStyleType.Marker) ds = ps.GetDisplayStyleProfile();
                        break;
                    case RcsViewType.Section:
                        if (component == PointDisplayStyleType.Marker) ds = ps.GetDisplayStyleSection();
                        break;
                }

                if (ds != null)
                {
                    ds.Layer = layerName;
                    ds.Visible = true;
                }
            }
            catch { }
        }

        private static void EnsureLayersExist(Database db, Transaction tr, HashSet<string> layers)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            foreach (var layName in layers)
            {
                if (string.IsNullOrWhiteSpace(layName) || layName == "0") continue;
                if (!lt.Has(layName))
                {
                    try
                    {
                        lt.UpgradeOpen();
                        var ltr = new LayerTableRecord { Name = layName };
                        lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }
                    catch { }
                }
            }
        }

        private static string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "InvalidName";
            string s = name.Replace('+', '_').Replace(':', '_').Replace('*', '_').Replace('?', '_');
            char[] invalid = Path.GetInvalidFileNameChars();
            return new string(s.Where(c => !invalid.Contains(c)).ToArray());
        }

        private static List<CsvStyleRow> LoadCsv(string path)
        {
            var list = new List<CsvStyleRow>();
            try
            {
                var lines = File.ReadAllLines(path);
                int start = 1;
                for (int i = start; i < lines.Length; i++)
                {
                    var parts = lines[i].Split(',');
                    if (parts.Length < 10) continue;

                    var row = new CsvStyleRow
                    {
                        StyleName = parts[0].Trim(),
                        PlanMarkerLayer = parts[1].Trim(),
                        PlanLabelLayer = parts[2].Trim(),
                        ModelMarkerLayer = parts[3].Trim(),
                        ModelLabelLayer = parts[4].Trim(),
                        ProfileMarkerLayer = parts[5].Trim(),
                        ProfileLabelLayer = parts[6].Trim(),
                        SectionMarkerLayer = parts[7].Trim(),
                        SectionLabelLayer = parts[8].Trim(),
                        MarkerType = parts[9].Trim(),
                        MarkerSymbolName = parts[10].Trim()
                    };
                    list.Add(row);
                }
            }
            catch { }
            return list;
        }

        private class CsvStyleRow
        {
            public string StyleName { get; set; } = "";
            public string PlanMarkerLayer { get; set; } = "";
            public string PlanLabelLayer { get; set; } = "";
            public string ModelMarkerLayer { get; set; } = "";
            public string ModelLabelLayer { get; set; } = "";
            public string ProfileMarkerLayer { get; set; } = "";
            public string ProfileLabelLayer { get; set; } = "";
            public string SectionMarkerLayer { get; set; } = "";
            public string SectionLabelLayer { get; set; } = "";
            public string MarkerType { get; set; } = "";
            public string MarkerSymbolName { get; set; } = "";
        }

        private static void RunCommandSafe(string name, Action act)
        {
            try { Info(null, $"{name} START"); act(); }
            catch (System.Exception ex) { Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\nCMD Error: {ex.Message}"); }
        }

        private static bool PromptYesNo(Editor ed, string msg, bool defaultYes)
        {
            var pko = new PromptKeywordOptions(msg) { AppendKeywordsToMessage = true };
            pko.Keywords.Add("No");
            pko.Keywords.Add("Yes");
            pko.Keywords.Default = defaultYes ? "Yes" : "No";
            try { return ed.GetKeywords(pko).StringResult == "Yes"; } catch { return defaultYes; }
        }

        private static void Info(Editor? ed, string msg) => Write(ed, "INFO", msg);
        private static void Warn(Editor? ed, string msg) => Write(ed, "WARN", msg);
        private static void Write(Editor? ed, string lvl, string msg)
        {
            try { ed?.WriteMessage($"\n[{lvl}] {msg}"); } catch { }
            try { File.AppendAllText(LogPath, $"[{DateTime.Now}] [{lvl}] {msg}\n"); } catch { }
        }
    }
}