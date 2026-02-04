
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace RCS.C3D2025
{
    /// <summary>
    /// RCS PointStyle Export/Delete/Import (CSV) — MERGED v1
    /// .NET Framework 4.8-safe
    ///
    /// Export current PointStyles (layers for Plan/Model/Profile/Section Marker + Label where applicable) to CSV,
    /// then optionally delete and re-create from that CSV.
    ///
    /// Export CSV: C:\temp\rcs_pointstyles_export.csv
    /// Log:        C:\temp\c3doutput.txt
    ///
    /// Commands:
    ///  - RCS_EXPORT_POINTSTYLES_CSV
    ///  - RCS_DELETE_POINTSTYLES_FROM_CSV
    ///  - RCS_IMPORT_POINTSTYLES_FROM_CSV
    ///  - RCS_EXPORT_DELETE_IMPORT_POINTSTYLES (guided)
    /// </summary>
    public class RcsPointStyleExportDeleteImport
    {
        private const string ExportCsvPath = @"C:\temp\rcs_pointstyles_export.csv";

        // -------------------------
        // Commands
        // -------------------------

        [CommandMethod("RCS_EXPORT_POINTSTYLES_CSV")]
        public static void RCS_EXPORT_POINTSTYLES_CSV()
        {
            RcsLog.RunCommandSafe("RCS_EXPORT_POINTSTYLES_CSV", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active drawing.");
                var ed = doc.Editor;
                var db = doc.Database;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var civDoc = CivilApplication.ActiveDocument ?? throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    PreflightCanonicalLayers(db, tr, ed);

                    var rows = ExportAllPointStylesToRows(db, tr, civDoc, ed);
                    WriteCsv(rows, ExportCsvPath);

                    tr.Commit();
                    ed.WriteMessage($"\nExported {rows.Count} point styles to: {ExportCsvPath}\nLog: {RcsLog.LogPath}\n");
                }
            });
        }

        [CommandMethod("RCS_DELETE_POINTSTYLES_FROM_CSV")]
        public static void RCS_DELETE_POINTSTYLES_FROM_CSV()
        {
            RcsLog.RunCommandSafe("RCS_DELETE_POINTSTYLES_FROM_CSV", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active drawing.");
                var ed = doc.Editor;
                var db = doc.Database;

                string path = PromptForPathOrDefault(ed, "\nCSV path to DELETE PointStyles from", ExportCsvPath);
                if (!File.Exists(path)) throw new FileNotFoundException("CSV not found: " + path);

                if (!PromptYesNo(ed, "\nThis will ERASE point styles listed in the CSV. Continue?", false))
                {
                    ed.WriteMessage("\nCanceled.\n");
                    return;
                }

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var civDoc = CivilApplication.ActiveDocument ?? throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    var rows = ReadCsv(path);
                    int deleted = 0, skipped = 0, missing = 0, failed = 0;

                    var map = BuildPointStyleNameMap(tr, civDoc, ed);

                    foreach (var r in rows)
                    {
                        if (string.IsNullOrWhiteSpace(r.StyleName)) continue;

                        if (!map.TryGetValue(r.StyleName, out ObjectId id))
                        {
                            missing++;
                            continue;
                        }

                        try
                        {
                            var psObj = tr.GetObject(id, OpenMode.ForWrite, false);
                            if (psObj == null) { missing++; continue; }

                            if (string.Equals(SafeName(psObj), "Standard", StringComparison.OrdinalIgnoreCase))
                            {
                                skipped++;
                                continue;
                            }

                            psObj.Erase(true);
                            deleted++;
                        }
                        catch (System.Exception ex)
                        {
                            failed++;
                            RcsLog.Warn(ed, $"Delete failed for '{r.StyleName}': {ex.Message}");
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nDelete complete. Deleted={deleted} Skipped={skipped} Missing={missing} Failed={failed}\nLog: {RcsLog.LogPath}\n");
                }
            });
        }

        [CommandMethod("RCS_IMPORT_POINTSTYLES_FROM_CSV")]
        public static void RCS_IMPORT_POINTSTYLES_FROM_CSV()
        {
            RcsLog.RunCommandSafe("RCS_IMPORT_POINTSTYLES_FROM_CSV", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active drawing.");
                var ed = doc.Editor;
                var db = doc.Database;

                string path = PromptForPathOrDefault(ed, "\nCSV path to IMPORT PointStyles from", ExportCsvPath);
                if (!File.Exists(path)) throw new FileNotFoundException("CSV not found: " + path);

                bool overwrite   = PromptYesNo(ed, "\nIf a PointStyle exists, delete & recreate it (overwrite)?", false);
                bool setAllViews = PromptYesNo(ed, "\nSet Model/Profile/Section layers too? (Plan always set)", true);
                bool ensureLayers = PromptYesNo(ed, "\nAuto-create missing layers referenced by the CSV?", true);

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var civDoc = CivilApplication.ActiveDocument ?? throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    PreflightCanonicalLayers(db, tr, ed);

                    var rows = ReadCsv(path);
                    if (rows.Count == 0) throw new InvalidOperationException("CSV is empty or invalid: " + path);

                    int created = 0, createdButUnlayered = 0, skipped = 0, failed = 0;

                    var map = BuildPointStyleNameMap(tr, civDoc, ed);

                    foreach (var r in rows)
                    {
                        if (string.IsNullOrWhiteSpace(r.StyleName)) { skipped++; continue; }

                        try
                        {
                            // normalize
                            r.PlanMarkerLayer    = NormalizeLayerName(r.PlanMarkerLayer);
                            r.PlanLabelLayer     = NormalizeLayerName(r.PlanLabelLayer);
                            r.ModelMarkerLayer   = NormalizeLayerName(r.ModelMarkerLayer);
                            r.ModelLabelLayer    = NormalizeLayerName(r.ModelLabelLayer);
                            r.ProfileMarkerLayer = NormalizeLayerName(r.ProfileMarkerLayer);
                            r.SectionMarkerLayer = NormalizeLayerName(r.SectionMarkerLayer);

                            // optional layer creation
                            ObjectId planMarkerId, planLabelId, modelMarkerId, modelLabelId, profileMarkerId, sectionMarkerId;

                            if (ensureLayers)
                            {
                                planMarkerId = EnsureLayer(db, tr, r.PlanMarkerLayer);
                                planLabelId  = EnsureLayer(db, tr, r.PlanLabelLayer);

                                modelMarkerId   = EnsureLayer(db, tr, r.ModelMarkerLayer);
                                modelLabelId    = EnsureLayer(db, tr, r.ModelLabelLayer);
                                profileMarkerId = EnsureLayer(db, tr, r.ProfileMarkerLayer);
                                sectionMarkerId = EnsureLayer(db, tr, r.SectionMarkerLayer);
                            }
                            else
                            {
                                planMarkerId = GetLayerIdIfExists(db, tr, r.PlanMarkerLayer);
                                planLabelId  = GetLayerIdIfExists(db, tr, r.PlanLabelLayer);

                                modelMarkerId   = GetLayerIdIfExists(db, tr, r.ModelMarkerLayer);
                                modelLabelId    = GetLayerIdIfExists(db, tr, r.ModelLabelLayer);
                                profileMarkerId = GetLayerIdIfExists(db, tr, r.ProfileMarkerLayer);
                                sectionMarkerId = GetLayerIdIfExists(db, tr, r.SectionMarkerLayer);
                            }

                            // overwrite delete
                            if (map.TryGetValue(r.StyleName, out ObjectId existingId))
                            {
                                if (!overwrite) { skipped++; continue; }

                                try
                                {
                                    var existing = tr.GetObject(existingId, OpenMode.ForWrite, false);
                                    if (existing != null && !existing.IsErased) existing.Erase(true);
                                    map.Remove(r.StyleName);
                                }
                                catch (System.Exception exDel)
                                {
                                    failed++;
                                    RcsLog.Warn(ed, $"Overwrite delete failed for '{r.StyleName}': {exDel.Message}");
                                    continue;
                                }
                            }

                            // create
                            ObjectId newId = civDoc.Styles.PointStyles.Add(r.StyleName);
                            var ps = tr.GetObject(newId, OpenMode.ForWrite, false) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (ps == null) { failed++; RcsLog.Warn(ed, $"Create returned null PointStyle for '{r.StyleName}'."); continue; }

                            bool okAny = false;

                            okAny |= SetPointStyleComponentLayer(ps, "Plan", "Marker", r.PlanMarkerLayer, planMarkerId, ed);
                            okAny |= SetPointStyleComponentLayer(ps, "Plan", "Label",  r.PlanLabelLayer,  planLabelId,  ed);

                            if (setAllViews)
                            {
                                okAny |= SetPointStyleComponentLayer(ps, "Model",   "Marker", r.ModelMarkerLayer,   modelMarkerId,   ed);
                                okAny |= SetPointStyleComponentLayer(ps, "Model",   "Label",  r.ModelLabelLayer,    modelLabelId,    ed);

                                okAny |= SetPointStyleComponentLayer(ps, "Profile", "Marker", r.ProfileMarkerLayer, profileMarkerId, ed);
                                okAny |= SetPointStyleComponentLayer(ps, "Section", "Marker", r.SectionMarkerLayer, sectionMarkerId, ed);
                            }

                            if (okAny) created++;
                            else createdButUnlayered++;

                            map[r.StyleName] = newId;
                        }
                        catch (System.Exception ex)
                        {
                            failed++;
                            RcsLog.Warn(ed, $"Import failed for '{r.StyleName}': {ex.Message}");
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nImport complete. Created={created} CreatedButUnlayered={createdButUnlayered} Skipped={skipped} Failed={failed}\nLog: {RcsLog.LogPath}\n");
                }
            });
        }

        [CommandMethod("RCS_EXPORT_DELETE_IMPORT_POINTSTYLES")]
        public static void RCS_EXPORT_DELETE_IMPORT_POINTSTYLES()
        {
            RcsLog.RunCommandSafe("RCS_EXPORT_DELETE_IMPORT_POINTSTYLES", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active drawing.");
                var ed = doc.Editor;

                bool doExport = PromptYesNo(ed, "\nPhase 1: Export current PointStyles to CSV?", true);
                bool doDelete = PromptYesNo(ed, "\nPhase 2: Delete PointStyles listed in the CSV?", false);
                bool doImport = PromptYesNo(ed, "\nPhase 3: Import (re-create) PointStyles from the CSV?", true);

                if (doExport) RCS_EXPORT_POINTSTYLES_CSV();
                if (doDelete) RCS_DELETE_POINTSTYLES_FROM_CSV();
                if (doImport) RCS_IMPORT_POINTSTYLES_FROM_CSV();
            });
        }

        // -------------------------
        // Row model
        // -------------------------
        private class Row
        {
            public string StyleName;

            public string PlanMarkerLayer;
            public string PlanLabelLayer;

            public string ModelMarkerLayer;
            public string ModelLabelLayer;

            public string ProfileMarkerLayer;   // marker only
            public string SectionMarkerLayer;   // marker only

            public string Notes;
        }

        // -------------------------
        // Export
        // -------------------------
        private static List<Row> ExportAllPointStylesToRows(Database db, Transaction tr, CivilDocument civDoc, Editor ed)
        {
            var rows = new List<Row>();
            int scanned = 0, exported = 0, failed = 0;

            foreach (ObjectId id in civDoc.Styles.PointStyles)
            {
                scanned++;
                try
                {
                    var ps = tr.GetObject(id, OpenMode.ForRead, false) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                    if (ps == null) continue;

                    string name = GetStyleName(ps);
                    if (string.IsNullOrWhiteSpace(name)) name = SafeName(ps);
                    if (string.IsNullOrWhiteSpace(name)) name = "(unnamed)";

                    var r = new Row();
                    r.StyleName = name;

                    r.PlanMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Plan", "Marker"));
                    r.PlanLabelLayer  = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Plan", "Label"));

                    r.ModelMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Model", "Marker"));
                    r.ModelLabelLayer  = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Model", "Label"));

                    r.ProfileMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Profile", "Marker"));
                    r.SectionMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Section", "Marker"));

                    // stabilize blanks
                    if (string.IsNullOrWhiteSpace(r.PlanMarkerLayer)) r.PlanMarkerLayer = "0";
                    if (string.IsNullOrWhiteSpace(r.PlanLabelLayer))  r.PlanLabelLayer  = "0";

                    if (string.IsNullOrWhiteSpace(r.ModelMarkerLayer)) r.ModelMarkerLayer = r.PlanMarkerLayer;
                    if (string.IsNullOrWhiteSpace(r.ModelLabelLayer))  r.ModelLabelLayer  = r.PlanLabelLayer;

                    if (string.IsNullOrWhiteSpace(r.ProfileMarkerLayer)) r.ProfileMarkerLayer = r.PlanMarkerLayer;
                    if (string.IsNullOrWhiteSpace(r.SectionMarkerLayer)) r.SectionMarkerLayer = r.PlanMarkerLayer;

                    rows.Add(r);
                    exported++;
                }
                catch (System.Exception ex)
                {
                    failed++;
                    RcsLog.Warn(ed, $"Export failed for style id={id.Handle}: {ex.Message}");
                }
            }

            RcsLog.Info(ed, $"EXPORT_POINTSTYLES: scanned={scanned} exported={exported} failed={failed}");
            return rows;
        }

        private static string ReadComponentLayer(Database db, Transaction tr, Autodesk.Civil.DatabaseServices.Styles.PointStyle ps, string viewName, string componentHint)
        {
            try
            {
                object col = GetDisplayCollection(ps, viewName);
                if (col == null && string.Equals(viewName, "Model", StringComparison.OrdinalIgnoreCase))
                    col = GetDisplayCollection(ps, "Plan");
                if (col == null) return "0";

                object comp = ResolveComponentFromDisplayCollection(col, componentHint);
                if (comp == null) return "0";

                // string
                string ln = ReadLayerNameProperty(comp);
                if (!string.IsNullOrWhiteSpace(ln)) return ln;

                // ObjectId
                ObjectId lid = ReadLayerIdProperty(comp);
                if (!lid.IsNull && !lid.IsErased)
                {
                    var ltr = tr.GetObject(lid, OpenMode.ForRead, false) as LayerTableRecord;
                    if (ltr != null && !string.IsNullOrWhiteSpace(ltr.Name)) return ltr.Name;
                }
            }
            catch { }
            return "0";
        }

        private static string ReadLayerNameProperty(object displayStyle)
        {
            if (displayStyle == null) return string.Empty;
            Type t = displayStyle.GetType();

            foreach (string pn in new[] { "LayerName", "Layer" })
            {
                try
                {
                    var pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null && pi.CanRead && pi.PropertyType == typeof(string))
                        return (string)(pi.GetValue(displayStyle, null) ?? string.Empty);
                }
                catch { }
            }
            return string.Empty;
        }

        private static ObjectId ReadLayerIdProperty(object displayStyle)
        {
            if (displayStyle == null) return ObjectId.Null;
            Type t = displayStyle.GetType();

            foreach (string pn in new[] { "LayerId", "LayerObjectId" })
            {
                try
                {
                    var pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null && pi.CanRead && pi.PropertyType == typeof(ObjectId))
                    {
                        object v = pi.GetValue(displayStyle, null);
                        if (v is ObjectId) return (ObjectId)v;
                    }
                }
                catch { }
            }
            return ObjectId.Null;
        }

        // -------------------------
        // Import setters
        // -------------------------

        private static bool SetPointStyleComponentLayer(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            try
            {
                if (string.Equals(componentHint, "Label", StringComparison.OrdinalIgnoreCase) &&
                    (string.Equals(viewName, "Profile", StringComparison.OrdinalIgnoreCase) || string.Equals(viewName, "Section", StringComparison.OrdinalIgnoreCase)))
                    return false;

                object col = GetDisplayCollection(ps, viewName);
                if (col == null && string.Equals(viewName, "Model", StringComparison.OrdinalIgnoreCase))
                    col = GetDisplayCollection(ps, "Plan");

                if (col == null) return false;

                object comp = ResolveComponentFromDisplayCollection(col, componentHint);
                if (comp != null)
                    return TrySetLayerOnDisplayStyle(comp, layerName, layerId);

                return TrySetLayerOnAllDisplayStyles(col, layerName, layerId);
            }
            catch (System.Exception ex)
            {
                RcsLog.Warn(ed, $"Set layer failed: view='{viewName}' comp='{componentHint}' style='{GetStyleName(ps)}' :: {ex.Message}");
                return false;
            }
        }

        private static object GetDisplayCollection(object pointStyle, string viewName)
        {
            if (pointStyle == null) return null;
            string v = (viewName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(v)) return null;

            try
            {
                Type t = pointStyle.GetType();
                PropertyInfo pi = t.GetProperty("DisplayStyle" + v, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (pi != null && pi.CanRead)
                    return pi.GetValue(pointStyle, null);
            }
            catch { }
            return null;
        }

        private static object ResolveComponentFromDisplayCollection(object displayCollection, string componentHint)
        {
            if (displayCollection == null) return null;
            string c = (componentHint ?? "").Trim();
            if (string.IsNullOrWhiteSpace(c)) return null;

            Type ct = displayCollection.GetType();

            // direct property
            try
            {
                PropertyInfo piComp = ct.GetProperty(c, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (piComp != null && piComp.CanRead)
                {
                    object comp = piComp.GetValue(displayCollection, null);
                    if (comp != null) return comp;
                }
            }
            catch { }

            // aliases
            try
            {
                string[] aliases;
                if (string.Equals(c, "Marker", StringComparison.OrdinalIgnoreCase))
                    aliases = new[] { "Marker", "PointMarker", "Symbol", "Point", "PointSymbol", "Tick", "Ticks" };
                else if (string.Equals(c, "Label", StringComparison.OrdinalIgnoreCase))
                    aliases = new[] { "Label", "Text", "Annotation", "PointLabel" };
                else
                    aliases = new[] { c };

                foreach (var a in aliases)
                {
                    PropertyInfo piA = ct.GetProperty(a, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (piA != null && piA.CanRead)
                    {
                        object compA = piA.GetValue(displayCollection, null);
                        if (compA != null) return compA;
                    }
                }
            }
            catch { }

            // indexer/default member
            try
            {
                MemberInfo[] defaults = ct.GetDefaultMembers();
                foreach (MemberInfo dm in defaults)
                {
                    PropertyInfo idx = dm as PropertyInfo;
                    if (idx == null) continue;

                    ParameterInfo[] pars = idx.GetIndexParameters();
                    if (pars == null || pars.Length != 1) continue;

                    Type indexType = pars[0].ParameterType;
                    object key = null;

                    if (indexType == typeof(string))
                        key = c;
                    else if (indexType != null && indexType.IsEnum)
                    {
                        try { key = Enum.Parse(indexType, c, true); } catch { key = null; }
                    }

                    if (key == null) continue;

                    try
                    {
                        object comp = idx.GetValue(displayCollection, new object[] { key });
                        if (comp != null) return comp;
                    }
                    catch { }
                }
            }
            catch { }

            return null;
        }

        private static bool TrySetLayerOnAllDisplayStyles(object displayCollection, string layerName, ObjectId layerId)
        {
            if (displayCollection == null) return false;

            bool any = false;
            Type ct = displayCollection.GetType();

            try
            {
                foreach (var pi in ct.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    if (!pi.CanRead) continue;
                    if (pi.GetIndexParameters().Length != 0) continue;

                    object val = null;
                    try { val = pi.GetValue(displayCollection, null); } catch { val = null; }
                    if (val == null) continue;

                    if (TrySetLayerOnDisplayStyle(val, layerName, layerId))
                        any = true;
                }
            }
            catch { }

            return any;
        }

        private static bool TrySetLayerOnDisplayStyle(object displayStyle, string layerName, ObjectId layerId)
        {
            if (displayStyle == null) return false;
            Type t = displayStyle.GetType();

            foreach (string pn in new[] { "LayerId", "LayerObjectId", "Layer" })
            {
                try
                {
                    PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null && pi.CanWrite && pi.PropertyType == typeof(ObjectId))
                    {
                        pi.SetValue(displayStyle, layerId, null);
                        return true;
                    }
                }
                catch { }
            }

            foreach (string pn in new[] { "LayerName", "Layer" })
            {
                try
                {
                    PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null && pi.CanWrite && pi.PropertyType == typeof(string))
                    {
                        pi.SetValue(displayStyle, layerName, null);
                        return true;
                    }
                }
                catch { }
            }

            foreach (string mn in new[] { "SetLayerId", "SetLayerName", "SetLayer" })
            {
                try
                {
                    MethodInfo mi = t.GetMethod(mn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (mi == null) continue;

                    var ps = mi.GetParameters();
                    if (ps == null || ps.Length != 1) continue;

                    if (ps[0].ParameterType == typeof(ObjectId))
                    {
                        mi.Invoke(displayStyle, new object[] { layerId });
                        return true;
                    }
                    if (ps[0].ParameterType == typeof(string))
                    {
                        mi.Invoke(displayStyle, new object[] { layerName });
                        return true;
                    }
                }
                catch { }
            }

            return false;
        }

        // -------------------------
        // Layer policy helpers
        // -------------------------

        private static string NormalizeLayerName(string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return "0";
            string s = layerName.Trim();

            if (string.Equals(s, "NONE", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "NULL", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "(NULL)", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "0", StringComparison.OrdinalIgnoreCase)) return "0";

            if (s.StartsWith("txt-", StringComparison.OrdinalIgnoreCase))
                s = "TXT-" + s.Substring(4);
            if (s.StartsWith("pnt-", StringComparison.OrdinalIgnoreCase))
                s = "PNT-" + s.Substring(4);

            return s;
        }

        private static void PreflightCanonicalLayers(Database db, Transaction tr, Editor ed)
        {
            int renamed = AutoRenameLowercasePrefixedLayers(db, tr, ed);

            var remaining = FindLowercasePrefixedLayers(db, tr);
            if (remaining.Count > 0)
            {
                string msg = "FAIL: Found non-canonical lowercase layer prefixes (must be TXT- / PNT-). " +
                             "Remaining=" + remaining.Count + " Example(s): " + string.Join(", ", remaining.Take(10));
                throw new InvalidOperationException(msg);
            }

            if (renamed > 0) RcsLog.Info(ed, "Preflight: renamed lowercase TXT-/PNT- layers to canonical uppercase. Renamed=" + renamed);
            else RcsLog.Info(ed, "Preflight: no lowercase TXT-/PNT- layers found.");
        }

        private static int AutoRenameLowercasePrefixedLayers(Database db, Transaction tr, Editor ed)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            var renames = new List<Tuple<string, string>>();

            foreach (ObjectId lid in lt)
            {
                LayerTableRecord ltr = tr.GetObject(lid, OpenMode.ForRead) as LayerTableRecord;
                if (ltr == null) continue;

                string name = ltr.Name ?? "";
                if (name.StartsWith("txt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("TXT-", StringComparison.Ordinal))
                    renames.Add(Tuple.Create(name, "TXT-" + name.Substring(4)));
                else if (name.StartsWith("pnt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("PNT-", StringComparison.Ordinal))
                    renames.Add(Tuple.Create(name, "PNT-" + name.Substring(4)));
            }

            int renamed = 0;
            foreach (var pair in renames)
            {
                try { RenameLayerWithMerge(db, tr, ed, pair.Item1, pair.Item2); renamed++; }
                catch (System.Exception ex) { RcsLog.Warn(ed, $"Preflight rename failed '{pair.Item1}' -> '{pair.Item2}': {ex.Message}"); }
            }

            return renamed;
        }

        private static List<string> FindLowercasePrefixedLayers(Database db, Transaction tr)
        {
            var bad = new List<string>();
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            foreach (ObjectId lid in lt)
            {
                LayerTableRecord ltr = tr.GetObject(lid, OpenMode.ForRead) as LayerTableRecord;
                if (ltr == null) continue;

                string name = ltr.Name ?? "";
                if (name.StartsWith("txt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("TXT-", StringComparison.Ordinal))
                    bad.Add(name);
                else if (name.StartsWith("pnt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("PNT-", StringComparison.Ordinal))
                    bad.Add(name);
            }
            return bad;
        }

        private static void RenameLayerWithMerge(Database db, Transaction tr, Editor ed, string oldName, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return;
            if (!LayerExists(db, tr, oldName)) return;

            if (LayerExists(db, tr, newName))
            {
                ReassignAllEntitiesLayer(db, tr, oldName, newName);
                var oldLtr = GetLayer(db, tr, oldName, OpenMode.ForWrite);
                if (oldLtr != null && !oldLtr.IsErased)
                {
                    try { oldLtr.Erase(true); } catch { }
                }
                return;
            }

            var ltrWrite = GetLayer(db, tr, oldName, OpenMode.ForWrite);
            if (ltrWrite == null) return;

            try
            {
                var cur = tr.GetObject(db.Clayer, OpenMode.ForRead) as LayerTableRecord;
                if (cur != null && string.Equals(cur.Name, oldName, StringComparison.Ordinal))
                {
                    var zero = GetLayer(db, tr, "0", OpenMode.ForRead);
                    if (zero != null) db.Clayer = zero.ObjectId;
                }
            }
            catch { }

            ltrWrite.Name = newName;
        }

        private static void ReassignAllEntitiesLayer(Database db, Transaction tr, string fromLayer, string toLayer)
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId btrId in bt)
            {
                BlockTableRecord btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null) continue;

                foreach (ObjectId entId in btr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Autodesk.AutoCAD.DatabaseServices.Entity;
                    if (ent == null) continue;
                    if (string.Equals(ent.Layer, fromLayer, StringComparison.Ordinal))
                        ent.Layer = toLayer;
                }
            }
        }

        // -------------------------
        // Layer table helpers
        // -------------------------

        private static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            try
            {
                string ln = NormalizeLayerName(layerName);
                if (!IsValidLayerName(ln)) ln = "0";

                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt.Has(ln)) return lt[ln];

                if (string.Equals(ln, "0", StringComparison.OrdinalIgnoreCase))
                    return lt["0"];

                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord { Name = ln };
                ObjectId id = lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
                return id;
            }
            catch { return ObjectId.Null; }
        }

        private static ObjectId GetLayerIdIfExists(Database db, Transaction tr, string layerName)
        {
            try
            {
                string ln = NormalizeLayerName(layerName);
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt.Has(ln)) return lt[ln];
            }
            catch { }
            return ObjectId.Null;
        }

        private static bool LayerExists(Database db, Transaction tr, string layerName)
        {
            try
            {
                string ln = NormalizeLayerName(layerName);
                if (string.Equals(ln, "0", StringComparison.OrdinalIgnoreCase)) return true;
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                return lt != null && lt.Has(ln);
            }
            catch { return false; }
        }

        private static LayerTableRecord GetLayer(Database db, Transaction tr, string layerName, OpenMode mode)
        {
            try
            {
                string ln = NormalizeLayerName(layerName);
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt == null || !lt.Has(ln)) return null;
                ObjectId id = lt[ln];
                if (id.IsNull || id.IsErased) return null;
                return (LayerTableRecord)tr.GetObject(id, mode);
            }
            catch { return null; }
        }

        private static bool IsValidLayerName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (name.Length > 255) return false;
            char[] bad = { '<', '>', '/', '\\', ':', ';', '"', '?', '*', '|', ',', '=', '`' };
            return name.IndexOfAny(bad) < 0;
        }

        // -------------------------
        // Name map + name helpers
        // -------------------------
        private static Dictionary<string, ObjectId> BuildPointStyleNameMap(Transaction tr, CivilDocument civDoc, Editor ed)
        {
            var map = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
            int scanned = 0, added = 0;

            foreach (ObjectId id in civDoc.Styles.PointStyles)
            {
                scanned++;
                if (id.IsNull || id.IsErased) continue;

                Autodesk.Civil.DatabaseServices.Styles.PointStyle ps = null;
                try { ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle; }
                catch { continue; }
                if (ps == null) continue;

                string name = GetStyleName(ps);
                if (string.IsNullOrWhiteSpace(name)) name = SafeName(ps);
                if (string.IsNullOrWhiteSpace(name)) continue;

                if (!map.ContainsKey(name))
                {
                    map[name] = id;
                    added++;
                }
            }

            RcsLog.Info(ed, $"PointStyles mapped by name: {added} (scanned={scanned})");
            return map;
        }

        private static string GetStyleName(object o)
        {
            if (o == null) return string.Empty;

            try
            {
                var sb = o as Autodesk.Civil.DatabaseServices.Styles.StyleBase;
                if (sb != null && !string.IsNullOrWhiteSpace(sb.Name))
                    return sb.Name;
            }
            catch { }

            try
            {
                Type t = o.GetType();
                foreach (string pn in new[] { "Name", "StyleName", "DisplayName" })
                {
                    PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null && pi.PropertyType == typeof(string))
                    {
                        string v = null;
                        try { v = pi.GetValue(o, null) as string; } catch { v = null; }
                        if (!string.IsNullOrWhiteSpace(v)) return v;
                    }
                }
            }
            catch { }

            return string.Empty;
        }

        private static string SafeName(object o)
        {
            if (o == null) return "(null)";
            try
            {
                PropertyInfo pi = o.GetType().GetProperty("Name", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi != null && pi.PropertyType == typeof(string))
                    return (string)(pi.GetValue(o, null) ?? "(unnamed)");
            }
            catch { }
            return o.GetType().Name;
        }

        // -------------------------
        // CSV IO
        // -------------------------
        private static void WriteCsv(List<Row> rows, string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var sw = new StreamWriter(path, false))
            {
                sw.WriteLine("StyleName,PlanMarkerLayer,PlanLabelLayer,ModelMarkerLayer,ModelLabelLayer,ProfileMarkerLayer,SectionMarkerLayer,Notes");
                foreach (var r in rows)
                {
                    sw.WriteLine(
                        Csv(r.StyleName) + "," +
                        Csv(r.PlanMarkerLayer) + "," +
                        Csv(r.PlanLabelLayer) + "," +
                        Csv(r.ModelMarkerLayer) + "," +
                        Csv(r.ModelLabelLayer) + "," +
                        Csv(r.ProfileMarkerLayer) + "," +
                        Csv(r.SectionMarkerLayer) + "," +
                        Csv(r.Notes)
                    );
                }
            }
        }

        private static List<Row> ReadCsv(string path)
        {
            var rows = new List<Row>();
            using (var sr = new StreamReader(path))
            {
                string header = sr.ReadLine();
                if (header == null) return rows;

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var cols = ParseCsvLine(line);
                    if (cols.Count < 7) continue;

                    rows.Add(new Row
                    {
                        StyleName = Col(cols, 0),
                        PlanMarkerLayer = Col(cols, 1),
                        PlanLabelLayer = Col(cols, 2),
                        ModelMarkerLayer = Col(cols, 3),
                        ModelLabelLayer = Col(cols, 4),
                        ProfileMarkerLayer = Col(cols, 5),
                        SectionMarkerLayer = Col(cols, 6),
                        Notes = Col(cols, 7)
                    });
                }
            }
            return rows;
        }

        private static string Col(List<string> cols, int idx) => (cols != null && idx >= 0 && idx < cols.Count) ? (cols[idx] ?? "") : "";

        private static string Csv(string s)
        {
            if (s == null) return "\"\"";
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        // Basic CSV parser for quoted fields
        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            if (line == null) return result;

            bool inQuotes = false;
            var cur = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQuotes)
                {
                    if (ch == '"')
                    {
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            cur.Append('"');
                            i++;
                        }
                        else inQuotes = false;
                    }
                    else cur.Append(ch);
                }
                else
                {
                    if (ch == '"') inQuotes = true;
                    else if (ch == ',')
                    {
                        result.Add(cur.ToString());
                        cur.Clear();
                    }
                    else cur.Append(ch);
                }
            }
            result.Add(cur.ToString());
            return result;
        }

        // -------------------------
        // Prompts
        // -------------------------
        private static bool PromptYesNo(Editor ed, string message, bool defaultYes)
        {
            try
            {
                var pko = new PromptKeywordOptions(message);
                pko.AllowNone = true;
                pko.Keywords.Add("No");
                pko.Keywords.Add("Yes");
                pko.Keywords.Default = defaultYes ? "Yes" : "No";
                var pr = ed.GetKeywords(pko);
                if (pr.Status == PromptStatus.OK)
                    return string.Equals(pr.StringResult, "Yes", StringComparison.OrdinalIgnoreCase);
                return defaultYes;
            }
            catch { return defaultYes; }
        }

        private static string PromptForPathOrDefault(Editor ed, string message, string defaultPath)
        {
            try
            {
                var pso = new PromptStringOptions(message + $" <{defaultPath}>: ");
                pso.AllowSpaces = true;
                var pr = ed.GetString(pso);
                if (pr.Status != PromptStatus.OK) return defaultPath;
                string s = (pr.StringResult ?? "").Trim();
                return string.IsNullOrWhiteSpace(s) ? defaultPath : s;
            }
            catch { return defaultPath; }
        }

        // -------------------------
        // Minimal logger (self-contained)
        // -------------------------
        private static class RcsLog
        {
            public static readonly string LogPath = @"C:\temp\c3doutput.txt";

            public static void RunCommandSafe(string name, Action act)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                var ed = doc?.Editor;
                try
                {
                    Info(ed, $"--- {name} START --- Drawing={(doc?.Name ?? "(null)")}");
                    act();
                    Info(ed, $"--- {name} END OK ---");
                }
                catch (System.Exception ex)
                {
                    Error(ed, $"--- {name} FAILED --- {ex.Message}\n{ex}");
                    ed?.WriteMessage($"\nERROR: {ex.Message}\n");
                }
            }

            public static void Info(Editor ed, string msg)  => Write(ed, "INFO", msg);
            public static void Warn(Editor ed, string msg)  => Write(ed, "WARN", msg);
            public static void Error(Editor ed, string msg) => Write(ed, "ERROR", msg);

            private static void Write(Editor ed, string lvl, string msg)
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                    File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{lvl}] [DWG={(Application.DocumentManager.MdiActiveDocument?.Name ?? "(null)")}] {msg}\n");
                }
                catch { }
                try { ed?.WriteMessage($"\n[{lvl}] {msg}"); } catch { }
            }
        }
    }
}
