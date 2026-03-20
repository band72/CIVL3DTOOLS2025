using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Entity = Autodesk.AutoCAD.DatabaseServices.Entity;
using Exception = System.Exception;

namespace RCS.C3D2025
{
    /// <summary>
    /// RCS PointStyle Export/Delete/Import (CSV) — MERGED v6.1 (Marker Info Fix)
    /// .NET Framework 4.8-safe
    /// 
    /// Fixes:
    /// - Corrected API call in GetMarkerInfo (GetMarkerStyle instead of GetMarkerDisplayStylePlan).
    /// - MarkerType and MarkerSymbolName now populate correctly.
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
                    var civDoc = GetCivilDoc(db) ?? throw new InvalidOperationException("CivilDocument could not be resolved (ActiveDocument/GetCivilDocument returned null).");

                    var rows = ExportAllPointStylesToRows(db, tr, civDoc, ed);
                    WriteCsv(rows, ExportCsvPath);

                    tr.Commit();
                    ed.WriteMessage($"\nExported {rows.Count} point styles to: {ExportCsvPath}\nLog: {RcsLog.LogPath}\n");
                }
            });
        }

        // Add this helper method inside the RcsPointStyleExportDeleteImport class (private section)
        private static CivilDocument GetCivilDoc(Database db)
        {
            // CivilDocument.GetCivilDocument is the standard way to get the CivilDocument for a Database
            return CivilApplication.ActiveDocument ?? CivilDocument.GetCivilDocument(db);
        }

        [CommandMethod("RCS_POINTSTYLES_JSON_TO_CSV")]
        public static void RCS_POINTSTYLES_JSON_TO_CSV()
        {
            RcsLog.RunCommandSafe("RCS_POINTSTYLES_JSON_TO_CSV", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active drawing.");
                var ed = doc.Editor;

                string jsonPath = PromptForPathOrDefault(ed, "\nJSON path to convert to CSV", @"C:\temp\rcs_pointstyles_export.json");
                if (!File.Exists(jsonPath)) throw new FileNotFoundException("JSON not found: " + jsonPath);

                string csvPath = PromptForPathOrDefault(ed, "\nOutput CSV path", ExportCsvPath);

                int wrote = ConvertJsonExportToCsv(jsonPath, csvPath, ed);
                ed.WriteMessage($"\nConverted JSON -> CSV. Rows={wrote}\nCSV: {csvPath}\nLog: {RcsLog.LogPath}\n");
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
                    var civDoc = GetCivilDoc(db) ?? throw new InvalidOperationException("CivilDocument could not be resolved (ActiveDocument/GetCivilDocument returned null).");

                    PreflightCanonicalLayers(db, tr, ed);

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

                            string psName = "(unknown)";
                            if (psObj is PointStyle ps) psName = ps.Name;

                            if (string.Equals(psName, "Standard", StringComparison.OrdinalIgnoreCase))
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

                bool overwrite = PromptYesNo(ed, "\nIf a PointStyle exists, delete & recreate it (overwrite)?", false);
                bool setAllViews = PromptYesNo(ed, "\nSet Model/Profile/Section layers too? (Plan always set)", true);
                bool ensureLayers = PromptYesNo(ed, "\nAuto-create missing layers referenced by the CSV?", true);

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var civDoc = GetCivilDoc(db) ?? throw new InvalidOperationException("CivilDocument could not be resolved (ActiveDocument/GetCivilDocument returned null).");

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
                            // normalize + enforce label layer rule (TXT-*)
                            r.PlanMarkerLayer = NormalizeLayerName(r.PlanMarkerLayer);
                            r.PlanLabelLayer = NormalizeLabelLayerName(r.PlanLabelLayer);

                            r.ModelMarkerLayer = NormalizeLayerName(r.ModelMarkerLayer);
                            r.ModelLabelLayer = NormalizeLabelLayerName(r.ModelLabelLayer);

                            r.ProfileMarkerLayer = NormalizeLayerName(r.ProfileMarkerLayer);
                            r.ProfileLabelLayer = NormalizeLabelLayerName(r.ProfileLabelLayer);

                            r.SectionMarkerLayer = NormalizeLayerName(r.SectionMarkerLayer);
                            r.SectionLabelLayer = NormalizeLabelLayerName(r.SectionLabelLayer);

                            // optional layer creation / lookup with fallback to Layer 0
                            ObjectId planMarkerId = SafeLayerId(db, tr, r.PlanMarkerLayer, ensureLayers, ed, "PlanMarker");
                            ObjectId planLabelId  = SafeLayerId(db, tr, r.PlanLabelLayer,  ensureLayers, ed, "PlanLabel");

                            ObjectId modelMarkerId = SafeLayerId(db, tr, r.ModelMarkerLayer, ensureLayers, ed, "ModelMarker");
                            ObjectId modelLabelId  = SafeLayerId(db, tr, r.ModelLabelLayer,  ensureLayers, ed, "ModelLabel");

                            ObjectId profileMarkerId = SafeLayerId(db, tr, r.ProfileMarkerLayer, ensureLayers, ed, "ProfileMarker");
                            ObjectId profileLabelId  = SafeLayerId(db, tr, r.ProfileLabelLayer,  ensureLayers, ed, "ProfileLabel");

                            ObjectId sectionMarkerId = SafeLayerId(db, tr, r.SectionMarkerLayer, ensureLayers, ed, "SectionMarker");
                            ObjectId sectionLabelId  = SafeLayerId(db, tr, r.SectionLabelLayer,  ensureLayers, ed, "SectionLabel");


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
                            ObjectId newId = AddPointStyleSafe(civDoc, tr, ed, r.StyleName);
                            if (newId.IsNull) { failed++; RcsLog.Warn(ed, $"Create failed for '{r.StyleName}': Add returned null (likely invalid name or styles container missing)."); continue; }
                            var ps = tr.GetObject(newId, OpenMode.ForWrite, false) as PointStyle;
                            if (ps == null) { failed++; RcsLog.Warn(ed, $"Create returned null PointStyle for '{r.StyleName}'."); continue; }

                            bool okAny = false;

                                                        okAny |= SafeSetComponentLayer(ps, "Plan", "Marker", r.PlanMarkerLayer, planMarkerId, db, tr, ed);
                                                        okAny |= SafeSetComponentLayer(ps, "Plan", "Label", r.PlanLabelLayer, planLabelId, db, tr, ed);

                            if (setAllViews)
                            {
                                                            okAny |= SafeSetComponentLayer(ps, "Model", "Marker", r.ModelMarkerLayer, modelMarkerId, db, tr, ed);
                                                            okAny |= SafeSetComponentLayer(ps, "Model", "Label", r.ModelLabelLayer, modelLabelId, db, tr, ed);

                                                            okAny |= SafeSetComponentLayer(ps, "Profile", "Marker", r.ProfileMarkerLayer, profileMarkerId, db, tr, ed);
                                                            okAny |= SafeSetComponentLayer(ps, "Profile", "Label", r.ProfileLabelLayer, profileLabelId, db, tr, ed);

                                                            okAny |= SafeSetComponentLayer(ps, "Section", "Marker", r.SectionMarkerLayer, sectionMarkerId, db, tr, ed);
                                                            okAny |= SafeSetComponentLayer(ps, "Section", "Label", r.SectionLabelLayer, sectionLabelId, db, tr, ed);
                            }

                            if (okAny) created++;
                            else createdButUnlayered++;

                            map[r.StyleName] = newId;
                        }
                        catch (System.Exception ex)
                        {
                            failed++;
                            RcsLog.Warn(ed, $"Import failed for '{r.StyleName}': {ExDetail(ex)}");
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nImport complete. Created={created} CreatedButUnlayered={createdButUnlayered} Skipped={skipped} Failed={failed}\nLog: {RcsLog.LogPath}\n");
                }
            });

}

        // Add this helper method inside the RcsPointStyleExportDeleteImport class (private section)
        private static string ExDetail(Exception ex)
        {
            if (ex == null) return "";
            return ex.ToString();
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

            public string ProfileMarkerLayer;
            public string ProfileLabelLayer;

            public string SectionMarkerLayer;
            public string SectionLabelLayer;

            public string MarkerType;       // NEW
            public string MarkerSymbolName; // NEW

            public string Notes;
        }


        // -------------------------
        // Export Logic
        // -------------------------
        private static List<Row> ExportAllPointStylesToRows(Database db, Transaction tr, CivilDocument civDoc, Editor ed)
        {
            var rows = new List<Row>();
            int scanned = 0, exported = 0, failed = 0;
            int patched = 0;

            var descKeyLayers = BuildDescriptorKeyLayerMap(tr, db, ed);
            RcsLog.Info(ed, $"Built Descriptor Key Map: found {descKeyLayers.Count} styles with assigned layers.");

            foreach (ObjectId id in SafeEnumerateObjectIds(civDoc.Styles.PointStyles))
            {
                scanned++;
                try
                {
                    var ps = tr.GetObject(id, OpenMode.ForRead, false) as PointStyle;
                    if (ps == null) continue;

                    string name = ps.Name; // Direct access
                    if (string.IsNullOrWhiteSpace(name)) name = "(unnamed)";

                    var r = new Row();
                    r.StyleName = name;

                    // Layers
                    r.PlanMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Plan", "Marker"));
                    r.PlanLabelLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Plan", "Label"));
                    r.ModelMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Model", "Marker"));
                    r.ModelLabelLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Model", "Label"));
                    r.ProfileMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Profile", "Marker"));
                    r.ProfileLabelLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Profile", "Label"));
                    r.SectionMarkerLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Section", "Marker"));
                    r.SectionLabelLayer = NormalizeLayerName(ReadComponentLayer(db, tr, ps, "Section", "Label"));

                    // NEW: Marker Info
                    GetMarkerInfo(tr, ps, out string mType, out string mName);
                    r.MarkerType = mType;
                    r.MarkerSymbolName = mName;

                    // Descriptor Key Enriched Logic
                    if (descKeyLayers.TryGetValue(id, out string keyLayer))
                    {
                        string targetLayer = NormalizeLayerName(keyLayer);
                        bool wasPatched = false;

                        if (IsLayerZero(r.PlanMarkerLayer)) { r.PlanMarkerLayer = targetLayer; wasPatched = true; }
                        if (IsLayerZero(r.PlanLabelLayer)) { r.PlanLabelLayer = targetLayer; wasPatched = true; }

                        if (IsLayerZero(r.ModelMarkerLayer)) r.ModelMarkerLayer = targetLayer;
                        if (IsLayerZero(r.ModelLabelLayer)) r.ModelLabelLayer = targetLayer;

                        if (IsLayerZero(r.ProfileMarkerLayer)) r.ProfileMarkerLayer = targetLayer;
                        if (IsLayerZero(r.ProfileLabelLayer)) r.ProfileLabelLayer = targetLayer;

                        if (IsLayerZero(r.SectionMarkerLayer)) r.SectionMarkerLayer = targetLayer;
                        if (IsLayerZero(r.SectionLabelLayer)) r.SectionLabelLayer = targetLayer;

                        if (wasPatched)
                        {
                            patched++;
                            r.Notes = $"Enriched via DescKey: {targetLayer}";
                        }
                    }

                    // Fallback stabilization
                    if (string.IsNullOrWhiteSpace(r.PlanMarkerLayer)) r.PlanMarkerLayer = "0";
                    if (string.IsNullOrWhiteSpace(r.PlanLabelLayer)) r.PlanLabelLayer = "0";
                    if (string.IsNullOrWhiteSpace(r.ModelMarkerLayer)) r.ModelMarkerLayer = r.PlanMarkerLayer;
                    if (string.IsNullOrWhiteSpace(r.ModelLabelLayer)) r.ModelLabelLayer = r.PlanLabelLayer;
                    if (string.IsNullOrWhiteSpace(r.ProfileMarkerLayer)) r.ProfileMarkerLayer = r.PlanMarkerLayer;
                    if (string.IsNullOrWhiteSpace(r.ProfileLabelLayer)) r.ProfileLabelLayer = r.PlanLabelLayer;
                    if (string.IsNullOrWhiteSpace(r.SectionMarkerLayer)) r.SectionMarkerLayer = r.PlanMarkerLayer;
                    if (string.IsNullOrWhiteSpace(r.SectionLabelLayer)) r.SectionLabelLayer = r.PlanLabelLayer;

                    rows.Add(r);
                    exported++;
                }
                catch (System.Exception ex)
                {
                    failed++;
                    RcsLog.Warn(ed, $"Export failed for style id={id.Handle}: {ex.Message}");
                }
            }

            RcsLog.Info(ed, $"EXPORT_POINTSTYLES: scanned={scanned} exported={exported} failed={failed}. Enriched via Keys={patched}");
            return rows;
        }

        // Add this helper method inside the RcsPointStyleExportDeleteImport class (private section)
        private static IEnumerable<ObjectId> SafeEnumerateObjectIds(IEnumerable<ObjectId> ids)
        {
            if (ids == null)
                yield break;
            foreach (var id in ids)
            {
                if (!id.IsNull && !id.IsErased)
                    yield return id;
            }
        }

        // -------------------------
        // NEW: Marker Info Extraction (FIXED API USAGE)
        // -------------------------
        // -------------------------
        // NEW: Marker Info Extraction
        // -------------------------
        private static void GetMarkerInfo(Transaction tr, PointStyle ps, out string type, out string name)
        {
            type = "";
            name = "";
            try
            {
                // Use PLAN view as the definition
                // FIX: Use the correct method for marker style in Plan view
                var marker = ps.GetMarkerDisplayStylePlan();
                if (marker == null) return;

                // Try to infer type from available properties
                var blockIdProp = marker.GetType().GetProperty("BlockId");
                if (blockIdProp != null)
                {
                    var blockId = blockIdProp.GetValue(marker, null) as ObjectId?;
                    if (blockId.HasValue && !blockId.Value.IsNull)
                    {
                        type = "UseBlockSymbol";
                        var btr = tr.GetObject(blockId.Value, OpenMode.ForRead) as BlockTableRecord;
                        if (btr != null) name = btr.Name;
                    }
                }
                else if (marker.GetType().GetProperty("MarkerType") != null)
                {
                    var markerTypeProp = marker.GetType().GetProperty("MarkerType");
                    var markerTypeValue = markerTypeProp.GetValue(marker, null);
                    if (markerTypeValue != null && markerTypeValue.ToString() == "UseCustomMarker")
                    {
                        // No CustomMarkerStyle property exists on DisplayStyle, so just set type
                        type = "UseCustomMarker";
                        name = "";
                    }
                    else if (markerTypeValue != null && markerTypeValue.ToString() == "UseAutoCADPoint")
                    {
                        type = "UseAutoCADPoint";
                        name = "AutoCAD Point";
                    }
                }
                else if (marker.GetType().GetProperty("MarkerType") != null)
                {
                    name = "AutoCAD Point";
                }
            }
            catch { }
        }
        // -------------------------
        // Descriptor Key Lookup Helper (Direct iteration)
        // -------------------------
        private static Dictionary<ObjectId, string> BuildDescriptorKeyLayerMap(Transaction tr, Database db, Editor ed)
        {
            var map = new Dictionary<ObjectId, string>();

            try
            {
                var setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
                if (setIds == null || setIds.Count == 0)
                {
                    RcsLog.Warn(ed, "No Point Description Key Sets found in drawing.");
                    return map;
                }

                foreach (ObjectId setId in setIds)
                {
                    var set = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                    if (set == null) continue;

                    ObjectIdCollection keyIds = null;
                    try { keyIds = set.GetPointDescriptionKeyIds(); }
                    catch { continue; }

                    foreach (ObjectId keyId in keyIds)
                    {
                        try
                        {
                            var key = tr.GetObject(keyId, OpenMode.ForRead) as PointDescriptionKey;
                            if (key == null) continue;

                            // only keys that assign BOTH a style and a layer are useful for enrichment
                            if (key.StyleId.IsNull || key.LayerId.IsNull) continue;

                            // LayerId might be erased/null in bad drawings
                            if (key.LayerId.IsErased) continue;

                            var ltr = tr.GetObject(key.LayerId, OpenMode.ForRead, false) as LayerTableRecord;
                            if (ltr == null) continue;

                            var layerName = ltr.Name ?? "";
                            if (string.IsNullOrWhiteSpace(layerName)) continue;

                            // first match wins (avoids fighting multiple keys)
                            if (!map.ContainsKey(key.StyleId))
                                map[key.StyleId] = layerName;
                        }
                        catch { /* ignore per-key */ }
                    }
                }
            }
            catch (System.Exception ex)
            {
                RcsLog.Warn(ed, "BuildDescriptorKeyLayerMap failed: " + ex.Message);
            }

            return map;
        }
        private static bool IsLayerZero(string layerName)
        {
            return string.IsNullOrWhiteSpace(layerName) ||
                   string.Equals(layerName, "0", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(layerName, "ByLayer", StringComparison.OrdinalIgnoreCase);
        }

        // -------------------------
        // Reading helpers
        // -------------------------

        private static string ReadComponentLayer(Database db, Transaction tr, PointStyle ps, string viewName, string componentHint)
        {
            try
            {
                object col = GetDisplayCollection(ps, viewName);
                if (col == null && string.Equals(viewName, "Model", StringComparison.OrdinalIgnoreCase))
                    col = GetDisplayCollection(ps, "Plan");
                if (col == null) return "0";

                object comp = ResolveComponentFromDisplayCollection(col, componentHint);
                if (comp == null) return "0";

                string ln = ReadLayerNameProperty(comp);
                if (!string.IsNullOrWhiteSpace(ln)) return ln;

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
            PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            try
            {
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
                RcsLog.Warn(ed, $"Set layer failed: view='{viewName}' comp='{componentHint}' style='{ps.Name}' :: {ex.Message}");
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

        
// -------------------------
// Safety wrappers (fallback to Layer 0)
// -------------------------

private static string NormalizeLabelLayerName(string layerName)
{
    // Rule: any Label column must be TXT-*, unless 0.
    string ln = NormalizeLayerName(layerName);
    if (string.Equals(ln, "0", StringComparison.OrdinalIgnoreCase)) return "0";
    if (!ln.StartsWith("TXT-", StringComparison.OrdinalIgnoreCase))
        ln = "TXT-" + ln.TrimStart('-', '_', ' ');
    return ln;
}

private static ObjectId GetLayer0Id(Database db, Transaction tr)
{
    try
    {
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
        if (lt.Has("0")) return lt["0"];
    }
    catch { }
    return ObjectId.Null;
}

private static ObjectId SafeLayerId(Database db, Transaction tr, string layerName, bool createIfMissing, Editor ed, string context)
{
    try
    {
        if (createIfMissing)
        {
            ObjectId id = EnsureLayer(db, tr, layerName);
            if (!id.IsNull && !id.IsErased) return id;
        }
        else
        {
            ObjectId id = GetLayerIdIfExists(db, tr, layerName);
            if (!id.IsNull && !id.IsErased) return id;
        }
    }
    catch (System.Exception ex)
    {
        RcsLog.Warn(ed, $"Layer resolve failed ({context}) layer='{layerName}': {ex.Message}");
    }

    // fallback to Layer 0
    ObjectId zero = GetLayer0Id(db, tr);
    return zero;
}

private static bool SafeSetComponentLayer(
    Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
    string viewName,
    string componentHint,
    string desiredLayerName,
    ObjectId desiredLayerId,
    Database db,
    Transaction tr,
    Editor ed)
{
    try
    {
        // Enforce label layer naming rule at import-time
        string ln = string.Equals(componentHint, "Label", StringComparison.OrdinalIgnoreCase)
            ? NormalizeLabelLayerName(desiredLayerName)
            : NormalizeLayerName(desiredLayerName);

        ObjectId lid = desiredLayerId;
        if (lid.IsNull || lid.IsErased)
            lid = SafeLayerId(db, tr, ln, true, ed, $"SafeSetComponentLayer {viewName}/{componentHint}");

        bool ok = SetPointStyleComponentLayer(ps, viewName, componentHint, ln, lid, ed);
        if (ok) return true;
    }
    catch (System.Exception ex)
    {
        RcsLog.Warn(ed, $"Set layer threw ({viewName}/{componentHint}) style='{GetStyleName(ps)}' layer='{desiredLayerName}': {ex.Message}");
    }

    // fallback set to 0
    try
    {
        ObjectId zero = GetLayer0Id(db, tr);
        return SetPointStyleComponentLayer(ps, viewName, componentHint, "0", zero, ed);
    }
    catch { return false; }
}


// With this correct implementation:
private static string GetStyleName(PointStyle ps)
{
    if (ps == null)
        return "(null)";
    try
    {
        return string.IsNullOrWhiteSpace(ps.Name) ? "(unnamed)" : ps.Name;
    }
    catch
    {
        return "(error)";
    }
}

private static string SanitizeStyleName(string name)
{
    if (string.IsNullOrWhiteSpace(name)) return "UNNAMED";
    string s = name.Trim();

    // Civil style names are more permissive than layer names, but certain chars can still trigger "Unknown Error".
    // Replace common troublemakers.
    char[] bad = { '<', '>', '/', '\\', ':', ';', '"', '?', '*', '|', ',', '=', '`', '+' };
    foreach (char ch in bad) s = s.Replace(ch, '_');

    // Collapse whitespace
    while (s.Contains("  ")) s = s.Replace("  ", " ");
    s = s.Trim();
    return string.IsNullOrWhiteSpace(s) ? "UNNAMED" : s;
}

private static ObjectId AddPointStyleSafe(CivilDocument civDoc, Transaction tr, Editor ed, string requestedName)
{
    string name1 = requestedName ?? "";
    try
    {
        return civDoc.Styles.PointStyles.Add(name1);
    }
    catch (System.Exception ex1)
    {
        string name2 = SanitizeStyleName(name1);
        if (!string.Equals(name1, name2, StringComparison.Ordinal))
        {
            try
            {
                RcsLog.Warn(ed, $"PointStyle name '{name1}' failed; retrying as '{name2}'. Error: {ex1.Message}");
                return civDoc.Styles.PointStyles.Add(name2);
            }
            catch (System.Exception ex2)
            {
                RcsLog.Warn(ed, $"PointStyle create failed for '{name2}': {ex2.Message}");
                return ObjectId.Null;
            }
        }
        RcsLog.Warn(ed, $"PointStyle create failed for '{name1}': {ex1.Message}");
        return ObjectId.Null;
    }
}

private static string NormalizeLayerName(string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return "0";
            string s = layerName.Trim();

            if (string.Equals(s, "NONE", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "NULL", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "(NULL)", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "0", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "ByLayer", StringComparison.OrdinalIgnoreCase)) return "0";

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
                    var ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
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

            foreach (ObjectId id in SafeEnumerateObjectIds(civDoc.Styles.PointStyles))
            {
                scanned++;
                if (id.IsNull || id.IsErased) continue;

                Autodesk.Civil.DatabaseServices.Styles.PointStyle ps = null;
                try { ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle; }
                catch { continue; }
                if (ps == null) continue;

                string name = ps.Name;
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

        // -------------------------
        // CSV IO
        // -------------------------

        private static int ConvertJsonExportToCsv(string jsonPath, string csvPath, Editor ed)
        {
            try
            {
                var ser = new DataContractJsonSerializer(typeof(PointStyleExportFile));
                PointStyleExportFile file;
                using (var fs = File.OpenRead(jsonPath))
                {
                    file = ser.ReadObject(fs) as PointStyleExportFile;
                }

                if (file == null || file.PointStyles == null || file.PointStyles.Count == 0)
                    throw new InvalidOperationException("JSON has no PointStyles.");

                var rows = new List<ExportRow>();
                foreach (var ps in file.PointStyles)
                {
                    if (ps == null) continue;

                    var dl = ps.DisplayLayers ?? new DisplayLayers();

                    rows.Add(new ExportRow
                    {
                        Name = ps.Name ?? "",
                        PlanMarker = dl.PlanMarker ?? "",
                        PlanLabel = dl.PlanLabel ?? "",
                        ModelMarker = dl.ModelMarker ?? "",
                        ModelLabel = dl.ModelLabel ?? "",
                        ProfileMarker = dl.ProfileMarker ?? "",
                        SectionMarker = dl.SectionMarker ?? "",
                        Display3dType = GetProp(ps, "Display3dType"),
                        CustomMarkerStyle = GetProp(ps, "CustomMarkerStyle"),
                        CustomMarkerSuperimposeStyle = GetProp(ps, "CustomMarkerSuperimposeStyle"),
                        Notes = ""
                    });
                }

                WriteCsv(rows.ConvertAll(er => new Row
                {
                    StyleName = er.Name,
                    PlanMarkerLayer = er.PlanMarker,
                    PlanLabelLayer = er.PlanLabel,
                    ModelMarkerLayer = er.ModelMarker,
                    ModelLabelLayer = er.ModelLabel,
                    ProfileMarkerLayer = er.ProfileMarker,
                    SectionMarkerLayer = er.SectionMarker,
                    Notes = er.Notes
                }), csvPath);

                return rows.Count;
            }
            catch (System.Exception ex)
            {
                RcsLog.Warn(ed, "JSON->CSV convert failed: " + ex.Message);
                throw;
            }
        }

        // Helper for intermediate export
        private class ExportRow
        {
            public string Name;
            public string PlanMarker;
            public string PlanLabel;
            public string ModelMarker;
            public string ModelLabel;
            public string ProfileMarker;
            public string SectionMarker;
            public string Display3dType;
            public string CustomMarkerStyle;
            public string CustomMarkerSuperimposeStyle;
            public string Notes;
        }

        private static string GetProp(PointStyleExport ps, string key)
        {
            try
            {
                if (ps?.Props == null || string.IsNullOrWhiteSpace(key)) return "";
                string v;
                if (ps.Props.TryGetValue(key, out v)) return v ?? "";
            }
            catch { }
            return "";
        }

        [DataContract]
        private class PointStyleExportFile
        {
            [DataMember] public string CivilVersion;
            [DataMember] public string ExportedUtc;
            [DataMember] public List<PointStyleExport> PointStyles;
        }

        [DataContract]
        private class PointStyleExport
        {
            [DataMember] public string Name;
            [DataMember] public DisplayLayers DisplayLayers;
            [DataMember] public Dictionary<string, string> Props;
        }

        [DataContract]
        private class DisplayLayers
        {
            [DataMember] public string PlanMarker;
            [DataMember] public string PlanLabel;
            [DataMember] public string ModelMarker;
            [DataMember] public string ModelLabel;
            [DataMember] public string ProfileMarker;
            [DataMember] public string SectionMarker;
        }

        private static void WriteCsv(List<Row> rows, string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var sw = new StreamWriter(path, false))
            {
                sw.WriteLine("StyleName,PlanMarkerLayer,PlanLabelLayer,ModelMarkerLayer,ModelLabelLayer,ProfileMarkerLayer,ProfileLabelLayer,SectionMarkerLayer,SectionLabelLayer,MarkerType,MarkerSymbolName,Notes");
                foreach (var r in rows)
                {
                    sw.WriteLine(
                        Csv(r.StyleName) + "," +
                        Csv(r.PlanMarkerLayer) + "," +
                        Csv(r.PlanLabelLayer) + "," +
                        Csv(r.ModelMarkerLayer) + "," +
                        Csv(r.ModelLabelLayer) + "," +
                        Csv(r.ProfileMarkerLayer) + "," +
                        Csv(r.ProfileLabelLayer) + "," +
                        Csv(r.SectionMarkerLayer) + "," +
                        Csv(r.SectionLabelLayer) + "," +
                        Csv(r.MarkerType) + "," +
                        Csv(r.MarkerSymbolName) + "," +
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
                    if (cols.Count < 9) continue; // adjusted for new cols

                    rows.Add(new Row
                    {
                        StyleName = Col(cols, 0),
                        PlanMarkerLayer = Col(cols, 1),
                        PlanLabelLayer = Col(cols, 2),
                        ModelMarkerLayer = Col(cols, 3),
                        ModelLabelLayer = Col(cols, 4),
                        ProfileMarkerLayer = Col(cols, 5),
                        ProfileLabelLayer = Col(cols, 6),
                        SectionMarkerLayer = Col(cols, 7),
                        SectionLabelLayer = Col(cols, 8),
                        MarkerType = Col(cols, 9),
                        MarkerSymbolName = Col(cols, 10),
                        Notes = Col(cols, 11)
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

            public static void Info(Editor ed, string msg) => Write(ed, "INFO", msg);
            public static void Warn(Editor ed, string msg) => Write(ed, "WARN", msg);
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