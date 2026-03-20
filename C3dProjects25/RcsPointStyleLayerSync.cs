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
using Autodesk.Civil.DatabaseServices;

namespace RCS.C3D2025
{
    /// <summary>
    /// Command: RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS
    ///
    /// Reads Description Key Sets, finds PointStyle referenced by keys, reads the key's Layer,
    /// then sets PointStyle display component layers (Marker + Label) for all views:
    ///   Plan, Model, Section, Profile
    ///
    /// Rules requested:
    ///   keyLayer = Description Key layer
    ///   markerLayer = keyLayer; if contains '-', remove first 4 chars (only if length>=4)
    ///   labelBase = markerLayer; if contains '-', remove first 4 chars (only if length>=4)
    ///   labelLayer = "txt-" + labelBase
    ///
    /// Robust error handling + logging through RcsError (C:\temp\c3doutput.txt)
    /// </summary>
    public class RcsPointStyleLayerSync
    {
        [CommandMethod("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS")]
        public static void RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument
                          ?? throw new InvalidOperationException("No active document.");

                var ed = doc.Editor;
                var db = doc.Database;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var civDoc = CivilApplication.ActiveDocument;
                    if (civDoc == null)
                        throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    // Map: PointStyleId -> keyLayer (first-hit wins)
                    var styleToKeyLayer = new Dictionary<ObjectId, string>(new ObjectIdComparer());

                    var setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
                    if (setIds == null || setIds.Count == 0)
                    {
                        RcsError.Warn(ed, "No description key sets found.");
                        return;
                    }

                    int setsScanned = 0, keysScanned = 0, mappingsAdded = 0;

                    foreach (ObjectId setId in setIds)
                    {
                        if (setId.IsNull || setId.IsErased) continue;

                        try
                        {
                            var set = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                            if (set == null) continue;

                            setsScanned++;

                            ObjectIdCollection keyIds;
                            try
                            {
                                keyIds = set.GetPointDescriptionKeyIds();
                            }
                            catch (System.Exception ex)
                            {
                                RcsError.Warn(ed, $"GetPointDescriptionKeyIds failed on set '{SafeName(set)}': {ex.Message}");
                                continue;
                            }

                            foreach (ObjectId keyId in keyIds)
                            {
                                if (keyId.IsNull || keyId.IsErased) continue;

                                try
                                {
                                    var key = tr.GetObject(keyId, OpenMode.ForRead) as PointDescriptionKey;
                                    if (key == null) continue;

                                    keysScanned++;

                                    // Pull PointStyleId (API differs across versions; try a few names + reflection)
                                    var pointStyleId = TryGetObjectIdProperty(key, "PointStyleId")
                                                    ?? TryGetObjectIdProperty(key, "PointStyle")
                                                    ?? ObjectId.Null;

                                    if (pointStyleId.IsNull)
                                        continue;

                                    // Pull key layer (common names differ by version; try a few)
                                    var keyLayer =
                                        TryGetStringProperty(key, "Layer")
                                        ?? TryGetStringProperty(key, "LayerName")
                                        ?? TryGetStringProperty(key, "PointLayer")
                                        ?? string.Empty;

                                    if (string.IsNullOrWhiteSpace(keyLayer)) continue;

                                    keyLayer = keyLayer.Trim();
                                    if (string.Equals(keyLayer, "0", StringComparison.OrdinalIgnoreCase))
                                        continue;

                                    if (!styleToKeyLayer.ContainsKey(pointStyleId))
                                    {
                                        styleToKeyLayer[pointStyleId] = keyLayer;
                                        mappingsAdded++;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    RcsError.Warn(ed, $"Key read failed (KeyId={keyId}): {ex.Message}");
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            RcsError.Warn(ed, $"Set read failed (SetId={setId}): {ex.Message}");
                        }
                    }

                    RcsError.Info(ed, $"Scanned sets={setsScanned}, keys={keysScanned}, mappings={mappingsAdded}");

                    if (styleToKeyLayer.Count == 0)
                    {
                        RcsError.Warn(ed, "No mappings found (no PointStyleId+Layer pairs in description keys).");
                        return;
                    }

                    int updated = 0, skipped = 0, labelUnsupported = 0;

                    foreach (var kvp in styleToKeyLayer)
                    {
                        var psId = kvp.Key;
                        var keyLayer = kvp.Value;

                        // MARKER + LABEL rules
                        var markerLayer = ComputeMarkerLayer(keyLayer);
                        var labelLayer = ComputeLabelLayer(markerLayer);

                        if (!IsValidLayerName(markerLayer))
                        {
                            RcsError.Warn(ed, $"Invalid marker layer '{markerLayer}' (from key layer '{keyLayer}'). Skipping.");
                            skipped++;
                            continue;
                        }
                        if (!IsValidLayerName(labelLayer))
                        {
                            RcsError.Warn(ed, $"Invalid label layer '{labelLayer}' (from marker '{markerLayer}'). Skipping.");
                            skipped++;
                            continue;
                        }

                        // Ensure layers exist
                        var markerLayerId = EnsureLayer(db, tr, markerLayer);
                        var labelLayerId = EnsureLayer(db, tr, labelLayer);

                        try
                        {
                            var ps = tr.GetObject(psId, OpenMode.ForWrite) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (ps == null)
                            {
                                RcsError.Warn(ed, $"ObjectId is not a PointStyle: {psId}");
                                skipped++;
                                continue;
                            }

                            bool markerOk = SetPointStyleDisplayLayerAllViews(ps, componentHint: "Marker", markerLayer, markerLayerId, ed);
                            bool labelOk  = SetPointStyleDisplayLayerAllViews(ps, componentHint: "Label",  labelLayer,  labelLayerId,  ed);

                            if (!labelOk) labelUnsupported++;

                            if (markerOk || labelOk) updated++;
                            else skipped++;
                        }
                        catch (System.Exception ex)
                        {
                            RcsError.Warn(ed, $"Update failed for PointStyleId={psId}: {ex.Message}");
                            skipped++;
                        }
                    }

                    tr.Commit();

                    ed.WriteMessage($"\nRCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS: Updated={updated}, Skipped={skipped}, LabelUnsupported={labelUnsupported}. See {RcsError.LogPath}");
                }
            });
        }

        // ---------------------- RULES ----------------------

        private static string ComputeMarkerLayer(string keyLayer)
        {
            string markerLayer = (keyLayer ?? string.Empty).Trim();

            if (markerLayer.Contains("-", StringComparison.Ordinal) && markerLayer.Length >= 4)
                markerLayer = markerLayer.Substring(4);

            return markerLayer;
        }

        private static string ComputeLabelLayer(string markerLayer)
        {
            string labelBase = (markerLayer ?? string.Empty).Trim();

            if (labelBase.Contains("-", StringComparison.Ordinal) && labelBase.Length >= 4)
                labelBase = labelBase.Substring(4);

            return "txt-" + labelBase;
        }

        // ----------------- DISPLAY STYLE SETTERS -----------------
        // Uses reflection because Civil 3D API differs across versions.

        private static bool SetPointStyleDisplayLayerAllViews(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            bool any = false;
            any |= TrySetView(ps, "Plan", componentHint, layerName, layerId, ed);
            any |= TrySetView(ps, "Model", componentHint, layerName, layerId, ed);
            any |= TrySetView(ps, "Section", componentHint, layerName, layerId, ed);
            any |= TrySetView(ps, "Profile", componentHint, layerName, layerId, ed);
            return any;
        }

        private static bool TrySetView(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            try
            {
                // 1) Method: GetDisplayStylePlan(enum)
                var mi = ps.GetType().GetMethod("GetDisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (mi != null)
                {
                    var asm = ps.GetType().Assembly;

                    var enumType =
                        asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointDisplayStyleType")
                        ?? asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointStyleDisplayStyleType");

                    if (enumType != null && enumType.IsEnum)
                    {
                        object? enumVal = null;
                        try { enumVal = Enum.Parse(enumType, componentHint, ignoreCase: true); }
                        catch { enumVal = null; }

                        if (enumVal != null)
                        {
                            var displayObj = mi.Invoke(ps, new[] { enumVal });
                            if (displayObj != null && TrySetDisplayLayer(displayObj, layerName, layerId))
                                return true;
                        }
                    }
                }

                // 2) Property: DisplayStylePlan (collection-like)
                var pi = ps.GetType().GetProperty("DisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (pi != null)
                {
                    var displayCollection = pi.GetValue(ps, null);
                    if (displayCollection != null)
                    {
                        // 2a) Property inside collection: Marker/Label
                        var compProp = displayCollection.GetType().GetProperty(componentHint, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                        if (compProp != null)
                        {
                            var compObj = compProp.GetValue(displayCollection, null);
                            if (compObj != null && TrySetDisplayLayer(compObj, layerName, layerId))
                                return true;
                        }

                        // 2b) Indexer on collection: this[enum] or this[string]
                        var defaultMembers = displayCollection.GetType().GetDefaultMembers();
                        foreach (var dm in defaultMembers)
                        {
                            if (dm is PropertyInfo prop && prop.GetIndexParameters().Length == 1)
                            {
                                var indexType = prop.GetIndexParameters()[0].ParameterType;

                                object? key = null;
                                if (indexType == typeof(string))
                                    key = componentHint;
                                else if (indexType.IsEnum)
                                {
                                    try { key = Enum.Parse(indexType, componentHint, ignoreCase: true); }
                                    catch { key = null; }
                                }

                                if (key != null)
                                {
                                    var compObj = prop.GetValue(displayCollection, new[] { key });
                                    if (compObj != null && TrySetDisplayLayer(compObj, layerName, layerId))
                                        return true;
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                RcsError.Warn(ed, $"Set layer failed: PointStyle='{SafeName(ps)}' View={viewName} Component={componentHint} -> {layerName} :: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Tries to set layer on a Civil display-style object.
        /// Many expose either LayerName(string) or Layer(ObjectId)/LayerId(ObjectId).
        /// </summary>
        private static bool TrySetDisplayLayer(object displayObj, string layerName, ObjectId layerId)
        {
            var t = displayObj.GetType();

            var pLayerName = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerName != null && pLayerName.CanWrite && pLayerName.PropertyType == typeof(string))
            {
                pLayerName.SetValue(displayObj, layerName, null);
                return true;
            }

            var pLayer = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayer != null && pLayer.CanWrite && pLayer.PropertyType == typeof(ObjectId))
            {
                pLayer.SetValue(displayObj, layerId, null);
                return true;
            }

            var pLayerId = t.GetProperty("LayerId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerId != null && pLayerId.CanWrite && pLayerId.PropertyType == typeof(ObjectId))
            {
                pLayerId.SetValue(displayObj, layerId, null);
                return true;
            }

            return false;
        }

        // ---------------------- HELPERS ----------------------

        private static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            if (lt.Has(layerName))
                return lt[layerName];

            lt.UpgradeOpen();
            var ltr = new LayerTableRecord { Name = layerName };
            var id = lt.Add(ltr);
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

        private static string? TryGetStringProperty(object obj, string propName)
        {
            var pi = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pi == null) return null;
            if (pi.PropertyType != typeof(string)) return null;
            return pi.GetValue(obj, null) as string;
        }

        private static ObjectId? TryGetObjectIdProperty(object obj, string propName)
        {
            var pi = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pi == null) return null;
            if (pi.PropertyType != typeof(ObjectId)) return null;
            return (ObjectId)pi.GetValue(obj, null)!;
        }

        private static string SafeName(object? o)
        {
            if (o == null) return "(null)";
            try
            {
                var pi = o.GetType().GetProperty("Name", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi != null && pi.PropertyType == typeof(string))
                    return (string)(pi.GetValue(o, null) ?? "(unnamed)");
            }
            catch { }
            return o.GetType().Name;
        }

        private class ObjectIdComparer : IEqualityComparer<ObjectId>
        {
            public bool Equals(ObjectId x, ObjectId y) => x == y;
            public int GetHashCode(ObjectId obj) => obj.GetHashCode();
        }
    }
}
