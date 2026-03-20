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

namespace RCS.C3D2025
{
    /// <summary>
    /// v4 UPDATE:
    /// - Description Key CODE matching now STRIPS ANY '*' characters before matching.
    ///   Example: '*ABS', 'ABS*', '*ABS*' -> 'ABS'
    /// - Matching is case-insensitive.
    /// - PointStyle NAME is matched to cleaned DescKey CODE (preferred) or key NAME.
    /// - Uses DescKey Layer to drive PointStyle display Marker + Label layers.
    /// </summary>
    public class RcsPointStyleLayerSyncv3
    {
        [CommandMethod("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv3")]
        public static void RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv3", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active document.");
                var ed = doc.Editor;
                var db = doc.Database;

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var civDoc = CivilApplication.ActiveDocument ?? throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    var pointStyleNameToId = BuildPointStyleNameMap(tr, civDoc, ed);
                    if (pointStyleNameToId.Count == 0)
                    {
                        RcsError.Warn(ed, "No PointStyles found in drawing.");
                        return;
                    }

                    var descKeyNameToLayer = BuildDescKeyNameToLayerMap(tr, db, ed);
                    if (descKeyNameToLayer.Count == 0)
                    {
                        RcsError.Warn(ed, "No Description Key code/name -> layer mappings found.");
                        return;
                    }

                    int updated = 0, skipped = 0, noKeyMatch = 0;

                    foreach (var psPair in pointStyleNameToId)
                    {
                        var psName = psPair.Key;
                        var psId = psPair.Value;

                        if (!descKeyNameToLayer.TryGetValue(psName, out var keyLayer))
                        {
                            noKeyMatch++;
                            continue;
                        }

                        var markerLayer = ComputeMarkerLayer(keyLayer);
                        var labelLayer = ComputeLabelLayer(markerLayer);

                        var markerLayerId = EnsureLayer(db, tr, markerLayer);
                        var labelLayerId = EnsureLayer(db, tr, labelLayer);

                        try
                        {
                            var ps = tr.GetObject(psId, OpenMode.ForWrite) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (ps == null)
                            {
                                skipped++;
                                continue;
                            }

                            SetPointStyleDisplayLayerAllViews(ps, "Marker", markerLayer, markerLayerId, ed);
                            SetPointStyleDisplayLayerAllViews(ps, "Label", labelLayer, labelLayerId, ed);

                            updated++;
                            RcsError.Info(ed, $"UPDATED PointStyle '{psName}' -> marker='{markerLayer}', label='{labelLayer}'");
                        }
                        catch (System.Exception ex)
                        {
                            skipped++;
                            RcsError.Warn(ed, $"Update failed for PointStyle '{psName}': {ex.Message}");
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nRCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS: Updated={updated}, Skipped={skipped}, NoKeyMatch={noKeyMatch}. See {RcsError.LogPath}");
                }
            });
        }

        // ---------------------- DESC KEY MAP ----------------------

        private static Dictionary<string, string> BuildDescKeyNameToLayerMap(Transaction tr, Database db, Editor ed)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
            if (setIds == null || setIds.Count == 0) return map;

            foreach (ObjectId setId in setIds)
            {
                var set = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                if (set == null) continue;

                ObjectIdCollection keyIds;
                try { keyIds = set.GetPointDescriptionKeyIds(); }
                catch { continue; }

                foreach (ObjectId keyId in keyIds)
                {
                    var key = tr.GetObject(keyId, OpenMode.ForRead) as PointDescriptionKey;
                    if (key == null) continue;

                    var rawCode = TryGetStringProperty(key, "Code") ?? TryGetStringProperty(key, "Key") ?? "";
                    var cleanedCode = CleanDescKeyCode(rawCode);

                    if (string.IsNullOrWhiteSpace(cleanedCode))
                        cleanedCode = CleanDescKeyCode(SafeName(key));

                    if (string.IsNullOrWhiteSpace(cleanedCode))
                        continue;

                    var keyLayer = (TryGetStringProperty(key, "Layer")
                                   ?? TryGetStringProperty(key, "LayerName")
                                   ?? TryGetStringProperty(key, "PointLayer")
                                   ?? "").Trim();

                    if (string.IsNullOrWhiteSpace(keyLayer) || keyLayer == "0")
                        continue;

                    if (!map.ContainsKey(cleanedCode))
                        map[cleanedCode] = keyLayer;
                }
            }

            RcsError.Info(ed, $"DescKey mappings created: {map.Count}");
            return map;
        }

        // ---------------------- RULES ----------------------

        private static string CleanDescKeyCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return string.Empty;
            return code.Replace("*", "").Trim();
        }

        private static string ComputeMarkerLayer(string keyLayer)
        {
            string markerLayer = keyLayer.Trim();
            if (markerLayer.Contains("-", StringComparison.Ordinal) && markerLayer.Length >= 4)
                markerLayer = markerLayer.Substring(4);
            return markerLayer;
        }

        private static string ComputeLabelLayer(string markerLayer)
        {
            string labelBase = markerLayer.Trim();
            if (labelBase.Contains("-", StringComparison.Ordinal) && labelBase.Length >= 4)
                labelBase = labelBase.Substring(4);
            return "txt-" + labelBase;
        }

        // ---------------------- DISPLAY SETTERS ----------------------

        private static void SetPointStyleDisplayLayerAllViews(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            TrySetView(ps, "Plan", componentHint, layerName, layerId, ed);
            TrySetView(ps, "Model", componentHint, layerName, layerId, ed);
            TrySetView(ps, "Section", componentHint, layerName, layerId, ed);
            TrySetView(ps, "Profile", componentHint, layerName, layerId, ed);
        }

        private static void TrySetView(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            try
            {
                var mi = ps.GetType().GetMethod("GetDisplayStyle" + viewName, BindingFlags.Instance | BindingFlags.Public);
                if (mi == null) return;

                var asm = ps.GetType().Assembly;
                var enumType = asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointDisplayStyleType")
                               ?? asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointStyleDisplayStyleType");
                if (enumType == null || !enumType.IsEnum) return;

                var enumVal = Enum.Parse(enumType, componentHint, true);
                var displayObj = mi.Invoke(ps, new[] { enumVal });
                if (displayObj == null) return;

                TrySetDisplayLayer(displayObj, layerName, layerId);
            }
            catch (System.Exception ex)
            {
                RcsError.Warn(ed, $"Layer set failed: {SafeName(ps)} {viewName}/{componentHint} -> {layerName} :: {ex.Message}");
            }
        }

        private static void TrySetDisplayLayer(object displayObj, string layerName, ObjectId layerId)
        {
            var t = displayObj.GetType();

            var p = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (p != null && p.CanWrite)
            {
                p.SetValue(displayObj, layerName);
                return;
            }

            p = t.GetProperty("LayerId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (p != null && p.CanWrite)
            {
                p.SetValue(displayObj, layerId);
            }
        }

        // ---------------------- HELPERS ----------------------

        private static Dictionary<string, ObjectId> BuildPointStyleNameMap(Transaction tr, CivilDocument civDoc, Editor ed)
        {
            var map = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId id in civDoc.Styles.PointStyles)
            {
                var ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                if (ps == null) continue;
                var name = SafeName(ps);
                if (!map.ContainsKey(name))
                    map[name] = id;
            }
            return map;
        }

        private static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName)) return lt[layerName];

            lt.UpgradeOpen();
            var ltr = new LayerTableRecord { Name = layerName };
            var id = lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
            return id;
        }

        private static string? TryGetStringProperty(object obj, string propName)
        {
            var pi = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            return pi?.GetValue(obj) as string;
        }

        private static string SafeName(object? o)
        {
            if (o == null) return string.Empty;
            var pi = o.GetType().GetProperty("Name", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            return pi?.GetValue(o)?.ToString() ?? string.Empty;
        }
    }
}
