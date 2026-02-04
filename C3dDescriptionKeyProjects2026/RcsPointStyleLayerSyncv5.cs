using System;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Assembly = System.Reflection.Assembly;

namespace RCS.C3D2025
{
    /// <summary>
    /// v6.1 (NET Framework 4.8-safe)
    /// Fixes compile errors caused by string.Contains(..., StringComparison) which is NOT available in .NET Framework 4.8.
    /// Uses IndexOf instead.
    ///
    /// Also keeps the v6 key-layer resolver (string or ObjectId) + RCS_DIAG_DESC_KEYS.
    /// </summary>
    public class RcsPointStyleLayerSyncv6
    {
        [CommandMethod("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv6")]
        public static void RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv6", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;
                Database db = doc.Database;

                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    CivilDocument civDoc = CivilApplication.ActiveDocument;
                    if (civDoc == null) throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    Dictionary<string, ObjectId> pointStyleNameToId = BuildPointStyleNameMap(tr, civDoc, ed);
                    if (pointStyleNameToId.Count == 0)
                    {
                        RcsError.Warn(ed, "No PointStyles found in drawing.");
                        return;
                    }

                    Dictionary<string, string> descKeyNameToLayer = BuildDescKeyNameToLayerMap(tr, db, ed);
                    if (descKeyNameToLayer.Count == 0)
                    {
                        RcsError.Warn(ed, "No Description Key code/name -> layer mappings found. Run RCS_DIAG_DESC_KEYS to inspect key layer fields.");
                        return;
                    }

                    int updated = 0, skipped = 0, noKeyMatch = 0, notLayer0 = 0;

                    foreach (KeyValuePair<string, ObjectId> psPair in pointStyleNameToId)
                    {
                        string psName = psPair.Key;
                        ObjectId psId = psPair.Value;

                        string keyLayer;
                        if (!descKeyNameToLayer.TryGetValue(psName, out keyLayer))
                        {
                            noKeyMatch++;
                            continue;
                        }

                        string markerLayer;
                        markerLayer = ComputeMarkerLayer(keyLayer);
                        string labelLayer = ComputeLabelLayerFromMarker(markerLayer);

                        if (!IsValidLayerName(markerLayer) || !IsValidLayerName(labelLayer))
                        {
                            RcsError.Warn(ed, "Invalid computed layer(s) for '" + psName + "': keyLayer='" + keyLayer + "' marker='" + markerLayer + "' label='" + labelLayer + "'. Skipping.");
                            skipped++;
                            continue;
                        }

                        ObjectId markerLayerId = EnsureLayer(db, tr, markerLayer);
                        ObjectId labelLayerId = EnsureLayer(db, tr, labelLayer);

                        Autodesk.Civil.DatabaseServices.Styles.PointStyle ps =
                            tr.GetObject(psId, OpenMode.ForWrite) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;

                        if (ps == null)
                        {
                            skipped++;
                            continue;
                        }

                        // Gate: only update if CURRENT Plan Marker is Layer 0 or blank
                        string currentPlanMarker = GetCurrentComponentLayer(ps, "Plan", "Marker", ed);
                        if (!string.IsNullOrWhiteSpace(currentPlanMarker) &&
                            !string.Equals(currentPlanMarker, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            notLayer0++;
                            continue;
                        }

                        bool markerOk = SetPointStyleDisplayLayerAllViews(ps, "Marker", markerLayer, markerLayerId, ed);
                        bool labelOk = SetPointStyleDisplayLayerAllViews(ps, "Label", labelLayer, labelLayerId, ed);

                        if (markerOk || labelOk)
                        {
                            updated++;
                            RcsError.Info(ed, "UPDATED '" + psName + "': keyLayer='" + keyLayer + "' -> marker='" + markerLayer + "', label='" + labelLayer + "'");
                        }
                        else
                        {
                            skipped++;
                            RcsError.Warn(ed, "No compatible setter found for '" + psName + "'.");
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage("\nRCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS: Updated=" + updated +
                                    ", Skipped=" + skipped +
                                    ", NoKeyMatch=" + noKeyMatch +
                                    ", NotLayer0=" + notLayer0 +
                                    ". See " + RcsError.LogPath);
                }
            });
        }

        [CommandMethod("RCS_DIAG_DESC_KEYS")]
        public static void RCS_DIAG_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_DIAG_DESC_KEYS", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;
                Database db = doc.Database;

                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    PointDescriptionKeySetCollection setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
                    if (setIds == null || setIds.Count == 0)
                    {
                        RcsError.Warn(ed, "No description key sets found.");
                        return;
                    }

                    int dumped = 0;
                    foreach (ObjectId setId in setIds)
                    {
                        PointDescriptionKeySet set = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                        if (set == null) continue;

                        ObjectIdCollection keyIds;
                        try { keyIds = set.GetPointDescriptionKeyIds(); }
                        catch { continue; }

                        foreach (ObjectId keyId in keyIds)
                        {
                            if (dumped >= 20) break;

                            PointDescriptionKey key = tr.GetObject(keyId, OpenMode.ForRead) as PointDescriptionKey;
                            if (key == null) continue;

                            string rawCode = TryGetStringProperty(key, "Code") ?? TryGetStringProperty(key, "Key") ?? "";
                            string cleaned = CleanDescKeyCode(rawCode);
                            if (string.IsNullOrWhiteSpace(cleaned))
                                cleaned = CleanDescKeyCode(SafeName(key));

                            string layer = ResolveKeyLayerName(tr, key);
                            RcsError.Info(ed, "KEY_SAMPLE #" + (dumped + 1) +
                                             ": Code='" + cleaned +
                                             "' RawCode='" + rawCode +
                                             "' Name='" + SafeName(key) +
                                             "' Layer='" + layer + "'");
                            dumped++;
                        }

                        if (dumped >= 20) break;
                    }

                    RcsError.Info(ed, "RCS_DIAG_DESC_KEYS dumped=" + dumped + " (see " + RcsError.LogPath + ")");
                    tr.Commit();
                }
            });
        }

        // ---------------------- DESC KEY MAP ----------------------

        private static Dictionary<string, string> BuildDescKeyNameToLayerMap(Transaction tr, Database db, Editor ed)
        {
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            PointDescriptionKeySetCollection setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
            if (setIds == null || setIds.Count == 0) return map;

            int keysScanned = 0, mapped = 0, missingLayer = 0, missingCode = 0;

            foreach (ObjectId setId in setIds)
            {
                PointDescriptionKeySet set = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                if (set == null) continue;

                ObjectIdCollection keyIds;
                try { keyIds = set.GetPointDescriptionKeyIds(); }
                catch { continue; }

                foreach (ObjectId keyId in keyIds)
                {
                    PointDescriptionKey key = tr.GetObject(keyId, OpenMode.ForRead) as PointDescriptionKey;
                    if (key == null) continue;

                    keysScanned++;

                    string rawCode = TryGetStringProperty(key, "Code") ?? TryGetStringProperty(key, "Key") ?? "";
                    string code = CleanDescKeyCode(rawCode);

                    if (string.IsNullOrWhiteSpace(code))
                        code = CleanDescKeyCode(SafeName(key));

                    if (string.IsNullOrWhiteSpace(code))
                    {
                        missingCode++;
                        continue;
                    }

                    string layer = ResolveKeyLayerName(tr, key);
                    if (string.IsNullOrWhiteSpace(layer) || string.Equals(layer, "0", StringComparison.OrdinalIgnoreCase))
                    {
                        missingLayer++;
                        continue;
                    }

                    if (!map.ContainsKey(code))
                    {
                        map[code] = layer;
                        mapped++;
                    }
                }
            }

            RcsError.Info(ed, "DescKey mappings: keysScanned=" + keysScanned +
                             ", mapped=" + mapped +
                             ", missingCode=" + missingCode +
                             ", missingLayer=" + missingLayer);
            return map;
        }

        // Resolve layer name from DescriptionKey (supports string or ObjectId properties)
        private static string ResolveKeyLayerName(Transaction tr, PointDescriptionKey key)
        {
            // 1) Known names first
            string[] known = new[] { "Layer", "LayerName", "PointLayer" };

            foreach (string prop in known)
            {
                string s = TryGetStringProperty(key, prop);
                if (!string.IsNullOrWhiteSpace(s)) return s.Trim();

                ObjectId oid = TryGetObjectIdProperty(key, prop);
                string name = LayerNameFromId(tr, oid);
                if (!string.IsNullOrWhiteSpace(name)) return name;
            }

            // 2) Scan any public property containing "Layer"
            PropertyInfo[] props = key.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public);
            foreach (PropertyInfo pi in props)
            {
                if (pi.Name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                try
                {
                    if (pi.PropertyType == typeof(string))
                    {
                        string s = pi.GetValue(key, null) as string;
                        if (!string.IsNullOrWhiteSpace(s)) return s.Trim();
                    }
                    else if (pi.PropertyType == typeof(ObjectId))
                    {
                        ObjectId oid = (ObjectId)pi.GetValue(key, null);
                        string name = LayerNameFromId(tr, oid);
                        if (!string.IsNullOrWhiteSpace(name)) return name;
                    }
                }
                catch
                {
                    // ignore and continue scanning
                }
            }

            return string.Empty;
        }

        private static string LayerNameFromId(Transaction tr, ObjectId id)
        {
            try
            {
                if (id.IsNull || id.IsErased) return string.Empty;
                LayerTableRecord ltr = tr.GetObject(id, OpenMode.ForRead, false) as LayerTableRecord;
                return ltr != null ? ltr.Name : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string CleanDescKeyCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return string.Empty;
            return code.Replace("*", "").Trim();
        }

        // ---------------------- RULES ----------------------

        private static string ComputeMarkerLayer(string keyLayer)
        {
            string markerLayer = (keyLayer ?? string.Empty).Trim();
            if (markerLayer.IndexOf('-') >= 0 && markerLayer.Length >= 4)
                markerLayer = markerLayer.Substring(4);
            return markerLayer;
        }

        private static string ComputeLabelLayerFromMarker(string markerLayer)
        {
            string labelBase = (markerLayer ?? string.Empty).Trim();
            if (labelBase.IndexOf('-') >= 0 && labelBase.Length >= 4)
                labelBase = labelBase.Substring(4);
            return "txt-" + labelBase;
        }

        // ---------------------- DISPLAY SETTERS (same approach as v5) ----------------------

        private static bool SetPointStyleDisplayLayerAllViews(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            bool any = false;
            any |= TrySetViewInternal(ps, "Plan", componentHint, layerName, layerId, ed);
            any |= TrySetViewInternal(ps, "Model", componentHint, layerName, layerId, ed);
            any |= TrySetViewInternal(ps, "Section", componentHint, layerName, layerId, ed);
            any |= TrySetViewInternal(ps, "Profile", componentHint, layerName, layerId, ed);
            return any;
        }

        private static bool TrySetViewInternal(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            if (TrySetViaDisplayStyleCollection(ps, viewName, componentHint, layerName, layerId, ed))
                return true;

            if (TrySetViaGetSet(ps, viewName, componentHint, layerName, layerId, ed))
                return true;

            return false;
        }

        private static bool TrySetViaDisplayStyleCollection(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            try
            {
                PropertyInfo pi = ps.GetType().GetProperty("DisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (pi == null) return false;

                object displayCollection = pi.GetValue(ps, null);
                if (displayCollection == null) return false;

                PropertyInfo compProp = displayCollection.GetType().GetProperty(componentHint, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (compProp != null)
                {
                    object compObj = compProp.GetValue(displayCollection, null);
                    if (compObj != null && TrySetDisplayLayer(compObj, layerName, layerId))
                        return true;
                }

                // Indexer fallback
                MemberInfo[] defaults = displayCollection.GetType().GetDefaultMembers();
                foreach (MemberInfo dm in defaults)
                {
                    PropertyInfo idx = dm as PropertyInfo;
                    if (idx == null) continue;

                    ParameterInfo[] pars = idx.GetIndexParameters();
                    if (pars.Length != 1) continue;

                    Type indexType = pars[0].ParameterType;
                    object key = null;

                    if (indexType == typeof(string))
                        key = componentHint;
                    else if (indexType.IsEnum)
                    {
                        try { key = Enum.Parse(indexType, componentHint, true); } catch { key = null; }
                    }

                    if (key != null)
                    {
                        object compObj = idx.GetValue(displayCollection, new object[] { key });
                        if (compObj != null && TrySetDisplayLayer(compObj, layerName, layerId))
                            return true;
                    }
                }
            }
            catch (System.Exception ex)
            {
                RcsError.Warn(ed, "DisplayStyleCollection set failed: " + SafeName(ps) + " " + viewName + "/" + componentHint + " -> " + layerName + " :: " + ex.Message);
            }

            return false;
        }

        private static bool TrySetViaGetSet(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            try
            {
                MethodInfo getMi = ps.GetType().GetMethod("GetDisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (getMi == null) return false;

                Type enumType = ResolvePointDisplayEnumType(ps);
                if (enumType == null) return false;

                object enumVal;
                try { enumVal = Enum.Parse(enumType, componentHint, true); }
                catch { return false; }

                object displayObj = getMi.Invoke(ps, new object[] { enumVal });
                if (displayObj == null) return false;

                if (!TrySetDisplayLayer(displayObj, layerName, layerId))
                    return false;

                MethodInfo setMi = ps.GetType().GetMethod("SetDisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (setMi != null)
                {
                    setMi.Invoke(ps, new object[] { enumVal, displayObj });
                    return true;
                }

                return true;
            }
            catch (System.Exception ex)
            {
                RcsError.Warn(ed, "Get/SetDisplayStyle failed: " + SafeName(ps) + " " + viewName + "/" + componentHint + " -> " + layerName + " :: " + ex.Message);
                return false;
            }
        }

        private static Type ResolvePointDisplayEnumType(Autodesk.Civil.DatabaseServices.Styles.PointStyle ps)
        {
            Assembly asm = ps.GetType().Assembly;
            return asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointDisplayStyleType")
                ?? asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointStyleDisplayStyleType");
        }

        private static bool TrySetDisplayLayer(object displayObj, string layerName, ObjectId layerId)
        {
            Type t = displayObj.GetType();

            PropertyInfo? pLayerName = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerName != null && pLayerName.CanWrite && pLayerName.PropertyType == typeof(string))
            {
                pLayerName.SetValue(displayObj, layerName, null);
                return true;
            }

            PropertyInfo? pLayer = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayer != null && pLayer.CanWrite && pLayer.PropertyType == typeof(ObjectId))
            {
                pLayer.SetValue(displayObj, layerId, null);
                return true;
            }

            PropertyInfo? pLayerId = t.GetProperty("LayerId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerId != null && pLayerId.CanWrite && pLayerId.PropertyType == typeof(ObjectId))
            {
                pLayerId.SetValue(displayObj, layerId, null);
                return true;
            }

            return false;
        }

        // Rename this method to avoid CS0111 duplicate definition error
        private static bool TrySetDisplayLayerInternal(object displayObj, string layerName, ObjectId layerId)
        {
            Type t = displayObj.GetType();

            PropertyInfo? pLayerName = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerName != null && pLayerName.CanWrite && pLayerName.PropertyType == typeof(string))
            {
                pLayerName.SetValue(displayObj, layerName, null);
                return true;
            }

            PropertyInfo? pLayer = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayer != null && pLayer.CanWrite && pLayer.PropertyType == typeof(ObjectId))
            {
                pLayer.SetValue(displayObj, layerId, null);
                return true;
            }

            PropertyInfo? pLayerId = t.GetProperty("LayerId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerId != null && pLayerId.CanWrite && pLayerId.PropertyType == typeof(ObjectId))
            {
                pLayerId.SetValue(displayObj, layerId, null);
                return true;
            }

            return false;
        }

        private static string GetCurrentComponentLayer(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            Editor ed)
        {
            try
            {
                PropertyInfo pi = ps.GetType().GetProperty("DisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (pi != null)
                {
                    object displayCollection = pi.GetValue(ps, null);
                    if (displayCollection != null)
                    {
                        PropertyInfo compProp = displayCollection.GetType().GetProperty(componentHint, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                        if (compProp != null)
                        {
                            object compObj = compProp.GetValue(displayCollection, null);
                            string ln = ReadLayerName(compObj);
                            if (!string.IsNullOrWhiteSpace(ln)) return ln;
                        }
                    }
                }

                MethodInfo getMi = ps.GetType().GetMethod("GetDisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (getMi == null) return string.Empty;

                Type enumType = ResolvePointDisplayEnumType(ps);
                if (enumType == null) return string.Empty;

                object enumVal;
                try { enumVal = Enum.Parse(enumType, componentHint, true); }
                catch { return string.Empty; }

                object displayObj = getMi.Invoke(ps, new object[] { enumVal });
                return ReadLayerName(displayObj);
            }
            catch (System.Exception ex)
            {
                RcsError.Warn(ed, "Read current layer failed: " + SafeName(ps) + " " + viewName + "/" + componentHint + " :: " + ex.Message);
                return string.Empty;
            }
        }

        private static string ReadLayerName(object displayObj)
        {
            if (displayObj == null) return string.Empty;

            Type t = displayObj.GetType();
            PropertyInfo p = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (p != null && p.PropertyType == typeof(string))
                return (string)(p.GetValue(displayObj, null) ?? string.Empty);

            p = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (p != null && p.PropertyType == typeof(string))
                return (string)(p.GetValue(displayObj, null) ?? string.Empty);

            return string.Empty;
        }

        // ---------------------- HELPERS ----------------------

        private static Dictionary<string, ObjectId> BuildPointStyleNameMap(Transaction tr, CivilDocument civDoc, Editor ed)
        {
            Dictionary<string, ObjectId> map = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
            int added = 0;

            foreach (ObjectId id in civDoc.Styles.PointStyles)
            {
                if (id.IsNull || id.IsErased) continue;

                Autodesk.Civil.DatabaseServices.Styles.PointStyle ps =
                    tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;

                if (ps == null) continue;

                string name = SafeName(ps);
                if (string.IsNullOrWhiteSpace(name)) continue;

                if (!map.ContainsKey(name))
                {
                    map[name] = id;
                    added++;
                }
            }

            RcsError.Info(ed, "PointStyles mapped by name: " + added);
            return map;
        }

        private static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (lt.Has(layerName))
                return lt[layerName];

            lt.UpgradeOpen();
            LayerTableRecord ltr = new LayerTableRecord { Name = layerName };
            ObjectId id = lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
            return id;
        }

        private static bool IsValidLayerName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (name.Length > 255) return false;

            char[] bad = { '<', '>', '/', '\\', ':', ';', '"', '?', '*', '|', ',', '=', '`' };
            return name.IndexOfAny(bad) < 0;
        }

        private static string TryGetStringProperty(object obj, string propName)
        {
            PropertyInfo pi = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pi == null) return null;
            if (pi.PropertyType != typeof(string)) return null;
            return pi.GetValue(obj, null) as string;
        }

        private static ObjectId TryGetObjectIdProperty(object obj, string propName)
        {
            try
            {
                PropertyInfo pi = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi == null) return ObjectId.Null;
                if (pi.PropertyType != typeof(ObjectId)) return ObjectId.Null;
                object v = pi.GetValue(obj, null);
                return v is ObjectId ? (ObjectId)v : ObjectId.Null;
            }
            catch
            {
                return ObjectId.Null;
            }
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
    }
}
