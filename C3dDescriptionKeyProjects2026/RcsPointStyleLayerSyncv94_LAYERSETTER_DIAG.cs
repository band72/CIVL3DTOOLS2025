using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using RCS.C3D2025;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using static System.Net.WebRequestMethods;
using DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject;
using Entity = Autodesk.AutoCAD.DatabaseServices.Entity;

namespace RCS.C3D2025
{
    /// <summary>
    /// v94 (NET Framework 4.8-safe) (NET Framework 4.8-safe)
    ///
    /// Goals:
    /// - Keep the working "direct setter" approach (no Get/SetDisplayStyle* calls -> avoids Parameter count mismatch).
    /// - Update Plan view (required) and optionally Model/Profile/Section by reading DisplayStyle{View} collections
    ///   and setting the component (Marker/Label) layer via direct properties (LayerName / LayerId / Layer).
    /// - Keep DescKey layer resolver (string/ObjectId/field scan) + forced-to-0 CSV output.
    /// </summary>
    public class RcsPointStyleLayerSyncv94
    {
        private const string ForcedLayer0CsvPath = @"C:\temp\descKeys_forced_layer0.csv";


        private static int _layerSetterDiagCount = 0;
        private const int _layerSetterDiagMax = 15;

        private class ForcedLayer0Row
        {
            public string SetName;
            public string Code;
            public string RawCode;
            public string KeyName;
            public string ResolvedRaw;
            public string Normalized;
            public string Reason;
        }

        [CommandMethod("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv94")]
        public static void RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv94", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;
                Database db = doc.Database;

                bool forceAll = false;
                bool ensureKeyLayers = false;
                bool setAllViews = false;

                try
                {
                    var pko = new PromptKeywordOptions("\nUpdate ALL PointStyles (ignore Layer0 gate)?");
                    pko.AllowNone = true;
                    pko.Keywords.Add("No");
                    pko.Keywords.Add("Yes");
                    pko.Keywords.Default = "No";
                    var pr = ed.GetKeywords(pko);
                    if (pr.Status == PromptStatus.OK && string.Equals(pr.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                        forceAll = true;

                    var pko2 = new PromptKeywordOptions("\nAuto-create missing KEY layers (non-zero) too?");
                    pko2.AllowNone = true;
                    pko2.Keywords.Add("No");
                    pko2.Keywords.Add("Yes");
                    pko2.Keywords.Default = "No";
                    var pr2 = ed.GetKeywords(pko2);
                    if (pr2.Status == PromptStatus.OK && string.Equals(pr2.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                        ensureKeyLayers = true;

                    var pko3 = new PromptKeywordOptions("\nSet Model/Profile/Section display layers too? (best-effort)");
                    pko3.AllowNone = true;
                    pko3.Keywords.Add("No");
                    pko3.Keywords.Add("Yes");
                    pko3.Keywords.Default = "Yes";
                    var pr3 = ed.GetKeywords(pko3);
                    if (pr3.Status == PromptStatus.OK && string.Equals(pr3.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                        setAllViews = true;
                }
                catch { }

                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    CivilDocument civDoc = CivilApplication.ActiveDocument;
                    if (civDoc == null) throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    // Preflight:
                    // - one-time cleanup: auto-rename lowercase txt-/pnt- layers to canonical TXT-/PNT-
                    // - audit: FAIL if any lowercase txt-/pnt- still exists
                    PreflightCanonicalLayers(db, tr, ed);


                    Dictionary<string, ObjectId> pointStyleNameToId = BuildPointStyleNameMap(tr, civDoc, ed);
                    if (pointStyleNameToId.Count == 0)
                    {
                        RcsError.Warn(ed, "No PointStyles found in drawing.");
                        return;
                    }

                    List<ForcedLayer0Row> forcedRows;
                    Dictionary<string, string> descKeyNameToLayer = BuildDescKeyNameToLayerMap(tr, db, ed, out forcedRows);
                    WriteForcedLayer0Csv(forcedRows, ed);

                    if (descKeyNameToLayer.Count == 0)
                    {
                        RcsError.Warn(ed, "No Description Key code/name -> layer mappings found. Run RCS_DIAG_DESC_KEYS to inspect key layer fields.");
                        return;
                    }

                    int updated = 0, skipped = 0, noKeyMatch = 0, notLayer0 = 0, secondaryMatch = 0;

                    foreach (var psPair in pointStyleNameToId)
                    {
                        string psName = psPair.Key;
                        ObjectId psId = psPair.Value;

                        try
                        {
                            string keyLayer;
                            string matchedKey;
                            bool usedSecondary;

                            if (!TryGetDescKeyLayerWithSecondaryRule(descKeyNameToLayer, psName, out keyLayer, out matchedKey, out usedSecondary))
                            {
                                noKeyMatch++;
                                continue;
                            }
                            if (usedSecondary) secondaryMatch++;

                            string markerLayer = NormalizeLayerName(ComputeMarkerLayer(keyLayer));
                            string labelLayer = NormalizeLayerName(ComputeLabelLayerFromMarker(markerLayer));

                            if (ensureKeyLayers && !string.Equals(keyLayer, "0", StringComparison.OrdinalIgnoreCase))
                            {
                                try { EnsureLayer(db, tr, keyLayer); } catch { }
                            }

                            if (!IsValidLayerName(markerLayer) || !IsValidLayerName(labelLayer))
                            {
                                skipped++;
                                continue;
                            }

                            ObjectId markerLayerId = EnsureLayer(db, tr, markerLayer);
                            ObjectId labelLayerId = EnsureLayer(db, tr, labelLayer);

                            var ps = tr.GetObject(psId, OpenMode.ForWrite) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (ps == null) { skipped++; continue; }

                            // Gate: only update if CURRENT Plan Marker is Layer 0 or blank (unless FORCEALL enabled)
                            if (!forceAll)
                            {
                                string currentPlanMarker = GetCurrentComponentLayerFromCollection(ps, "Plan", "Marker");
                                if (!string.IsNullOrWhiteSpace(currentPlanMarker) &&
                                    !string.Equals(currentPlanMarker, "0", StringComparison.OrdinalIgnoreCase))
                                {
                                    notLayer0++;
                                    continue;
                                }
                            }

                            bool markerOk = SetPointStyleComponentLayer(ps, "Plan", "Marker", markerLayer, markerLayerId, ed);
                            bool labelOk = SetPointStyleComponentLayer(ps, "Plan", "Label", labelLayer, labelLayerId, ed);

                            if (setAllViews)
                            {
                                // Model often shares Plan collection, but we try anyway.
                                TrySetAllViews(ps, "Marker", markerLayer, markerLayerId, ed);
                                TrySetAllViews(ps, "Label", labelLayer, labelLayerId, ed);
                            }

                            if (markerOk || labelOk)
                            {
                                updated++;
                                RcsError.Info(ed, "UPDATED '" + psName + "': key='" + matchedKey + "' layer='" + keyLayer + "' -> marker='" + markerLayer + "', label='" + labelLayer + "'" + (usedSecondary ? " (SECONDARY)" : ""));
                            }
                            else
                            {
                                skipped++;
                                RcsError.Warn(ed, "No compatible setter found for '" + psName + "'.");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            skipped++;
                            RcsError.Warn(ed, "PointStyle update failed: '" + psName + "' :: " + ex.Message);
                        }
                    }

                    // Prevent regression: lock canonical layers used for point markers/labels


                    // Optional: re-apply description keys to existing points so layers update in-place
                    bool autoReapply = false;
                    try
                    {
                        var pko4 = new PromptKeywordOptions("\nRe-apply Description Keys to EXISTING points now? (moves points to correct layers)");
                        pko4.AllowNone = true;
                        pko4.Keywords.Add("No");
                        pko4.Keywords.Add("Yes");
                        pko4.Keywords.Default = "Yes";
                        var pr4 = ed.GetKeywords(pko4);
                        if (pr4.Status == PromptStatus.OK && string.Equals(pr4.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                            autoReapply = true;
                    }
                    catch { }

                    if (autoReapply)
                    {
                        int scannedPts = 0, appliedPts = 0, failedPts = 0;
                        foreach (ObjectId pid in civDoc.CogoPoints)
                        {
                            scannedPts++;
                            try
                            {
                                var pt = tr.GetObject(pid, OpenMode.ForWrite, false) as CogoPoint;
                                if (pt == null) continue;
                                pt.ApplyDescriptionKeys();
                                appliedPts++;
                            }
                            catch { failedPts++; }
                        }
                        RcsError.Info(ed, $"AUTO_REAPPLY_DESC_KEYS: scanned={scannedPts} applied={appliedPts} failed={failedPts}");
                    }

                    LockCanonicalLayers(db, tr, ed);

                    tr.Commit();
                    ed.WriteMessage("\nRCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS: Updated=" + updated +
                                    ", Skipped=" + skipped +
                                    ", NoKeyMatch=" + noKeyMatch +
                                    ", NotLayer0=" + notLayer0 +
                                    ", SecondaryMatch=" + secondaryMatch +
                                    ". See " + RcsError.LogPath + "\nForced-to-0 CSV: " + ForcedLayer0CsvPath + "\n");
                }
            });
        }
        private static void TrySetAllViews(Autodesk.Civil.DatabaseServices.Styles.PointStyle ps, string component, string layerName, ObjectId layerId, Editor ed)
        {
            // What your UI exposes:
            // - Plan: Marker + Label
            // - Model: Marker + Label
            // - Profile: Marker only
            // - Section: Marker only
            //
            // So we ONLY attempt Label on Plan/Model. Profile/Section Label should not be auto-forced.
            try { SetPointStyleComponentLayer(ps, "Model", component, layerName, layerId, ed); } catch { }

            if (!string.Equals(component, "Label", StringComparison.OrdinalIgnoreCase))
            {
                try { SetPointStyleComponentLayer(ps, "Profile", component, layerName, layerId, ed); } catch { }
                try { SetPointStyleComponentLayer(ps, "Section", component, layerName, layerId, ed); } catch { }
            }
        }



        private static bool SetPointStyleComponentLayer(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            object col = GetDisplayCollection(ps, viewName);
            if (col == null)
            {
                // If Model doesn't exist, treat as Plan
                if (string.Equals(viewName, "Model", StringComparison.OrdinalIgnoreCase))
                    col = GetDisplayCollection(ps, "Plan");
            }

            if (col == null) return false;

            // 1) Try targeted component (Marker/Label)
            object comp = ResolveComponentFromDisplayCollection(col, componentHint);
            bool ok = false;

            if (comp != null)
                ok = TrySetLayerOnDisplayStyle(comp, layerName, layerId);

            if (!ok)
                DumpLayerSetterDiagnostics(comp, viewName, componentHint, ed);

            if (ok) return true;

            // 2) Fallback: attempt to set layer on any display style objects in the collection.
            // IMPORTANT: Profile/Section do NOT expose Label in your template (marker only).
            if (string.Equals(componentHint, "Label", StringComparison.OrdinalIgnoreCase) &&
                (string.Equals(viewName, "Profile", StringComparison.OrdinalIgnoreCase) || string.Equals(viewName, "Section", StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            bool any = TrySetLayerOnAllDisplayStyles(col, layerName, layerId);
            if (!any)
                DumpLayerSetterDiagnostics(col, viewName, componentHint + " (collection)", ed);

            return any;
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

            // common aliases (varies by Civil 3D build)
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

            // Try readable properties that look like display style components
            try
            {
                foreach (var pi in ct.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    if (!pi.CanRead) continue;
                    if (pi.GetIndexParameters().Length != 0) continue;

                    object val = null;
                    try { val = pi.GetValue(displayCollection, null); } catch { val = null; }
                    if (val == null) continue;

                    // If it has a layer property we can set, try it
                    if (TrySetLayerOnDisplayStyle(val, layerName, layerId))
                        any = true;
                }
            }
            catch { }

            // Try default members/indexer with common keys if present
            try
            {
                MemberInfo[] defaults = ct.GetDefaultMembers();
                foreach (MemberInfo dm in defaults)
                {
                    var idx = dm as PropertyInfo;
                    if (idx == null) continue;
                    var pars = idx.GetIndexParameters();
                    if (pars == null || pars.Length != 1) continue;

                    Type indexType = pars[0].ParameterType;

                    // Probe a few likely keys
                    object[] keys = null;
                    if (indexType == typeof(string))
                        keys = new object[] { "Marker", "Label", "Text", "Symbol", "Point", "Tick" };
                    else if (indexType != null && indexType.IsEnum)
                    {
                        var list = new List<object>();
                        foreach (var name in new[] { "Marker", "Label", "Text", "Symbol", "Point", "Tick" })
                        {
                            try { list.Add(Enum.Parse(indexType, name, true)); } catch { }
                        }
                        keys = list.ToArray();
                    }

                    if (keys == null || keys.Length == 0) continue;

                    foreach (var k in keys)
                    {
                        object comp = null;
                        try { comp = idx.GetValue(displayCollection, new object[] { k }); } catch { comp = null; }
                        if (comp == null) continue;

                        if (TrySetLayerOnDisplayStyle(comp, layerName, layerId))
                            any = true;
                    }
                }
            }
            catch { }

            return any;
        }

        private static bool TrySetLayerOnDisplayStyle(object displayStyle, string layerName, ObjectId layerId)
        {
            if (displayStyle == null) return false;
            Type t = displayStyle.GetType();

            // 1) LayerId/ObjectId
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

            // 2) LayerName string
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

            // 3) Method setters
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

            // 4) Last resort: any writable property containing "Layer"
            try
            {
                foreach (var pi in t.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    if (pi == null || !pi.CanWrite) continue;
                    string pn = pi.Name ?? "";
                    if (pn.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    try
                    {
                        if (pi.PropertyType == typeof(string))
                        {
                            pi.SetValue(displayStyle, layerName, null);
                            return true;
                        }
                        if (pi.PropertyType == typeof(ObjectId))
                        {
                            pi.SetValue(displayStyle, layerId, null);
                            return true;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return false;
        }

        private static string GetCurrentComponentLayerFromCollection(Autodesk.Civil.DatabaseServices.Styles.PointStyle ps, string viewName, string componentHint)
        {
            try
            {
                object col = GetDisplayCollection(ps, viewName);
                if (col == null) return string.Empty;

                object comp = ResolveComponentFromDisplayCollection(col, componentHint);
                if (comp == null) return string.Empty;

                // Read LayerName or Layer(string)
                Type t = comp.GetType();
                PropertyInfo p = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (p != null && p.PropertyType == typeof(string))
                    return (string)(p.GetValue(comp, null) ?? string.Empty);

                p = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (p != null && p.PropertyType == typeof(string))
                    return (string)(p.GetValue(comp, null) ?? string.Empty);
            }
            catch { }
            return string.Empty;
        }

        // ---------------------- DESC KEY MAP ----------------------

        private static Dictionary<string, string> BuildDescKeyNameToLayerMap(Transaction tr, Database db, Editor ed, out List<ForcedLayer0Row> forcedRows)
        {
            forcedRows = new List<ForcedLayer0Row>();
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            PointDescriptionKeySetCollection setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
            if (setIds == null || setIds.Count == 0) return map;

            int keysScanned = 0, mapped = 0, forcedLayer0 = 0, missingCode = 0;

            foreach (ObjectId setId in setIds)
            {
                PointDescriptionKeySet set = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                if (set == null) continue;

                string setName = SafeName(set);

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

                    string resolvedRaw = ResolveKeyLayerName(tr, key);
                    string layer = NormalizeLayerName(resolvedRaw);

                    // If the key resolves to a non-zero layer that does NOT exist in the drawing,
                    // treat it as unusable and force to Layer 0 (per RCS rule).
                    bool forcedBecauseMissingLayer = false;
                    if (!string.Equals(layer, "0", StringComparison.OrdinalIgnoreCase) && !LayerExists(db, tr, layer))
                    {
                        forcedBecauseMissingLayer = true;
                        forcedLayer0++;
                        forcedRows.Add(new ForcedLayer0Row
                        {
                            SetName = setName,
                            Code = code,
                            RawCode = rawCode,
                            KeyName = SafeName(key),
                            ResolvedRaw = resolvedRaw ?? string.Empty,
                            Normalized = "0",
                            Reason = "LayerMissing"
                        });
                        layer = "0";
                    }

                    if (!forcedBecauseMissingLayer && string.Equals(layer, "0", StringComparison.OrdinalIgnoreCase))
                    {
                        forcedLayer0++;
                        forcedRows.Add(new ForcedLayer0Row
                        {
                            SetName = setName,
                            Code = code,
                            RawCode = rawCode,
                            KeyName = SafeName(key),
                            ResolvedRaw = resolvedRaw ?? string.Empty,
                            Normalized = layer,
                            Reason = string.IsNullOrWhiteSpace(resolvedRaw) ? "Blank/Unresolved" : "NONE/NULL/0 token"
                        });
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
                             ", forcedLayer0=" + forcedLayer0);

            return map;
        }

        private static string ResolveKeyLayerName(Transaction tr, PointDescriptionKey key)
        {
            try
            {
                string[] known = new[] { "Layer", "LayerName", "PointLayer", "PointLayerId", "LayerId" };

                foreach (string prop in known)
                {
                    string s = TryGetStringProperty(key, prop);
                    if (!string.IsNullOrWhiteSpace(s)) return NormalizeLayerName(s);

                    ObjectId oid = TryGetObjectIdProperty(key, prop);
                    string name = LayerNameFromId(tr, oid);
                    if (!string.IsNullOrWhiteSpace(name)) return NormalizeLayerName(name);

                    PropertyInfo piAny = key.GetType().GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (piAny != null)
                    {
                        object v = null;
                        try { v = piAny.GetValue(key, null); } catch { v = null; }
                        string ln = ExtractLayerNameFromAny(tr, v);
                        if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                    }
                }

                PropertyInfo[] props = key.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (PropertyInfo pi in props)
                {
                    if (pi == null || pi.Name == null) continue;
                    if (pi.Name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    object v = null;
                    try { v = pi.GetValue(key, null); } catch { v = null; }

                    string ln = ExtractLayerNameFromAny(tr, v);
                    if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                }

                var fields = key.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (var fi in fields)
                {
                    if (fi == null || fi.Name == null) continue;
                    if (fi.Name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0) continue;

                    object v = null;
                    try { v = fi.GetValue(key); } catch { v = null; }

                    string ln = ExtractLayerNameFromAny(tr, v);
                    if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                }
            }
            catch { }

            return string.Empty;
        }

        private static string ExtractLayerNameFromAny(Transaction tr, object val)
        {
            try
            {
                if (val == null) return string.Empty;

                if (val is string)
                    return NormalizeLayerName((string)val);

                if (val is ObjectId)
                {
                    ObjectId oid = (ObjectId)val;
                    string ln = LayerNameFromId(tr, oid);
                    return string.IsNullOrWhiteSpace(ln) ? string.Empty : NormalizeLayerName(ln);
                }

                LayerTableRecord ltr = val as LayerTableRecord;
                if (ltr != null)
                    return NormalizeLayerName(ltr.Name);

                DBObject dbo = val as DBObject;
                if (dbo != null)
                {
                    string ln = LayerNameFromId(tr, dbo.ObjectId);
                    if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                }

                Type t = val.GetType();

                foreach (string pn in new[] { "LayerName", "Name", "Layer" })
                {
                    PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null)
                    {
                        object v = null;
                        try { v = pi.GetValue(val, null); } catch { v = null; }
                        if (v is string && !string.IsNullOrWhiteSpace((string)v))
                            return NormalizeLayerName((string)v);
                    }
                }

                foreach (string pn in new[] { "LayerId", "Id", "LayerObjectId" })
                {
                    PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null)
                    {
                        object v = null;
                        try { v = pi.GetValue(val, null); } catch { v = null; }
                        if (v is ObjectId)
                        {
                            string ln = LayerNameFromId(tr, (ObjectId)v);
                            if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                        }
                    }
                }
            }
            catch { }

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
            catch { return string.Empty; }
        }

        private static string CleanDescKeyCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return string.Empty;
            return code.Replace("*", "").Trim();
        }

        private static string NormalizeLayerName(string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return "0";
            string s = layerName.Trim();

            // Canonical remaps (keep your template consistent)
            // Legacy: some drawings store Control Points without the PNT- bucket
            if (string.Equals(s, "CONTROL POINTS", StringComparison.OrdinalIgnoreCase))
                s = "PNT-CONTROL POINTS";

            // Null-ish / none
            if (string.Equals(s, "NONE", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "NULL", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "(NULL)", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "0", StringComparison.OrdinalIgnoreCase)) return "0";

            // RCS rule: enforce prefix capitalization so matching is stable
            if (s.StartsWith("txt-", StringComparison.OrdinalIgnoreCase))
                s = "TXT-" + s.Substring(4);
            if (s.StartsWith("pnt-", StringComparison.OrdinalIgnoreCase))
                s = "PNT-" + s.Substring(4);

            return s;
        }

        // ---------------------- RULES ----------------------

        private static string ComputeMarkerLayer(string keyLayer)
        {
            string markerLayer = NormalizeLayerName((keyLayer ?? string.Empty).Trim());
            if (string.Equals(markerLayer, "0", StringComparison.OrdinalIgnoreCase)) return "0";
            if (markerLayer.IndexOf('-') >= 0 && markerLayer.Length >= 4)
                markerLayer = markerLayer.Substring(4);
            return markerLayer;
        }

        private static string ComputeLabelLayerFromMarker(string markerLayer)
        {
            string labelBase = (markerLayer ?? string.Empty).Trim();
            if (labelBase.IndexOf('-') >= 0 && labelBase.Length >= 4)
                labelBase = labelBase.Substring(4);
            if (string.Equals(labelBase, "0", StringComparison.OrdinalIgnoreCase)) return "0";
            return "TXT-" + labelBase;
        }

        // Secondary match rule (strip 2-char prefix for common discipline buckets)
        private static bool TryGetDescKeyLayerWithSecondaryRule(
            Dictionary<string, string> descKeyNameToLayer,
            string pointStyleName,
            out string keyLayer,
            out string matchedKey,
            out bool usedSecondary)
        {
            keyLayer = string.Empty;
            matchedKey = string.Empty;
            usedSecondary = false;

            if (string.IsNullOrWhiteSpace(pointStyleName))
                return false;

            string psNorm = pointStyleName.Trim().Replace("*", "").Trim();

            if (descKeyNameToLayer.TryGetValue(psNorm, out keyLayer))
            {
                matchedKey = psNorm;
                return true;
            }

            if (psNorm.Length > 2)
            {
                string prefix = psNorm.Substring(0, 2).ToUpperInvariant();
                switch (prefix)
                {
                    case "BC":
                    case "CP":
                    case "EL":
                    case "FC":
                    case "ML":
                    case "GS":
                    case "RR":
                    case "SD":
                    case "SL":
                    case "SN":
                    case "SS":
                    case "TP":
                    case "TL":
                    case "TR":
                    case "TS":
                    case "TV":
                    case "WA":
                    case "WL":
                        string stripped = psNorm.Substring(2).TrimStart('-', '_', ' ', '.');
                        if (descKeyNameToLayer.TryGetValue(stripped, out keyLayer))
                        {
                            matchedKey = stripped;
                            usedSecondary = true;
                            return true;
                        }
                        break;
                }
            }

            return false;
        }

        // ---------------------- HELPERS ----------------------

        private static Dictionary<string, ObjectId> BuildPointStyleNameMap(Transaction tr, CivilDocument civDoc, Editor ed)
        {
            Dictionary<string, ObjectId> map = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);

            int scanned = 0;
            int added = 0;
            int nullOrErased = 0;
            int openFail = 0;
            int unnamed = 0;

            foreach (ObjectId id in civDoc.Styles.PointStyles)
            {
                scanned++;

                if (id.IsNull || id.IsErased)
                {
                    nullOrErased++;
                    continue;
                }

                Autodesk.Civil.DatabaseServices.Styles.PointStyle ps = null;
                try { ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle; }
                catch { openFail++; continue; }

                if (ps == null) continue;

                string name = GetStyleName(ps);
                if (string.IsNullOrWhiteSpace(name)) name = SafeName(ps);

                if (string.IsNullOrWhiteSpace(name) || name == "(unnamed)" || string.Equals(name, ps.GetType().Name, StringComparison.OrdinalIgnoreCase))
                {
                    unnamed++;
                    try { name = "UNNAMED_" + id.Handle.ToString(); } catch { name = string.Empty; }
                }

                if (string.IsNullOrWhiteSpace(name)) continue;

                if (!map.ContainsKey(name))
                {
                    map[name] = id;
                    added++;
                }
            }

            RcsError.Info(ed, "PointStyles mapped by name: " + added + " (scanned=" + scanned +
                             ", nullOrErased=" + nullOrErased + ", openFail=" + openFail + ", unnamed=" + unnamed + ")");
            return map;
        }

        private static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            try
            {
                string ln = string.IsNullOrWhiteSpace(layerName) ? "0" : layerName.Trim();
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
            catch
            {
                try
                {
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has("0")) return lt["0"];
                }
                catch { }
                return ObjectId.Null;
            }
        }

        private static bool IsValidLayerName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            if (name.Length > 255) return false;

            char[] bad = { '<', '>', '/', '\\', ':', ';', '"', '?', '*', '|', ',', '=', '`' };
            return name.IndexOfAny(bad) < 0;
        }


        private static bool LayerExists(Database db, Transaction tr, string layerName)
        {
            try
            {
                if (db == null || tr == null) return false;
                string ln = string.IsNullOrWhiteSpace(layerName) ? "0" : layerName.Trim();
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
                if (db == null || tr == null) return null;
                string ln = string.IsNullOrWhiteSpace(layerName) ? "0" : layerName.Trim();

                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt == null || !lt.Has(ln)) return null;

                ObjectId id = lt[ln];
                if (id.IsNull || id.IsErased) return null;

                return (LayerTableRecord)tr.GetObject(id, mode);
            }
            catch { return null; }
        }

        private static string TryGetStringProperty(object obj, string propName)
        {
            try
            {
                PropertyInfo pi = obj.GetType().GetProperty(propName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi == null) return null;
                if (pi.PropertyType != typeof(string)) return null;
                return pi.GetValue(obj, null) as string;
            }
            catch { return null; }
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
            catch { return ObjectId.Null; }
        }

        private static string GetStyleName(object o)
        {
            if (o == null) return string.Empty;

            try
            {
                var sb = o as Autodesk.Civil.DatabaseServices.Styles.StyleBase;
                if (sb != null)
                {
                    string n = sb.Name;
                    if (!string.IsNullOrWhiteSpace(n)) return n;
                }
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

        private static void WriteForcedLayer0Csv(List<ForcedLayer0Row> rows, Editor ed)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(ForcedLayer0CsvPath));

                using (var sw = new StreamWriter(ForcedLayer0CsvPath, false))
                {
                    sw.WriteLine("SetName,Code,RawCode,KeyName,ResolvedRaw,Normalized,Reason");

                    if (rows != null)
                    {
                        foreach (var r in rows)
                        {
                            sw.WriteLine(
                                Csv(r.SetName) + "," +
                                Csv(r.Code) + "," +
                                Csv(r.RawCode) + "," +
                                Csv(r.KeyName) + "," +
                                Csv(r.ResolvedRaw) + "," +
                                Csv(r.Normalized) + "," +
                                Csv(r.Reason));
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                try { RcsError.Warn(ed, "Failed to write forced-to-0 CSV: " + ex.Message); } catch { }
            }
        }

        private static string Csv(string s)
        {
            if (s == null) return "\"\"";
            string v = s.Replace("\"", "\"\"");
            return "\"" + v + "\"";
        }

        // ---------------------- DIAG COMMAND ----------------------


        // ---------------------- APPLY DESCRIPTION KEYS ----------------------
        [CommandMethod("RCS_REAPPLY_DESC_KEYS")]
        public static void RCS_REAPPLY_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_REAPPLY_DESC_KEYS", () =>
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

                    int scanned = 0, applied = 0, failed = 0;

                    foreach (ObjectId id in civDoc.CogoPoints)
                    {
                        scanned++;
                        try
                        {
                            var pt = tr.GetObject(id, OpenMode.ForWrite, false) as CogoPoint;
                            if (pt == null) continue;

                            pt.ApplyDescriptionKeys();
                            applied++;
                        }
                        catch
                        {
                            failed++;
                        }
                    }

                    tr.Commit();

                    RcsError.Info(ed, "RCS_REAPPLY_DESC_KEYS complete. scanned=" + scanned + " applied=" + applied + " failed=" + failed);
                    ed.WriteMessage("\nRCS_REAPPLY_DESC_KEYS: scanned=" + scanned + " applied=" + applied + " failed=" + failed + ". See " + RcsError.LogPath + "\n");
                }
            });
        }



        [CommandMethod("RCS_DIAG_DESC_KEYS_V89")]
        public static void RCS_DIAG_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_DIAG_DESC_KEYS_V89", () =>
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

                    RcsError.Info(ed, "RCS_DIAG_DESC_KEYS_V89 dumped=" + dumped + " (see " + RcsError.LogPath + ")");
                    tr.Commit();
                }
            });
        }

        // ------------------------------
        // Preflight / One-time cleanup
        // ------------------------------
        static void PreflightCanonicalLayers(Database db, Transaction tr, Editor ed)
        {
            // 1) Auto-rename lowercase txt-/pnt- to canonical TXT-/PNT- (one-time cleanup)
            int renamed = AutoRenameLowercasePrefixedLayers(db, tr, ed);

            // 2) Audit: FAIL if any lowercase txt-/pnt- still exists (prevents regressions)
            var remaining = FindLowercasePrefixedLayers(db, tr);
            if (remaining.Count > 0)
            {
                string msg = "FAIL: Found non-canonical lowercase layer prefixes (must be TXT- / PNT-). " +
                             "Remaining=" + remaining.Count + " Example(s): " + string.Join(", ", remaining.Take(10));
                throw new InvalidOperationException(msg);
            }

            if (renamed > 0)
                RcsError.Info(ed, "Preflight: renamed lowercase TXT-/PNT- layers to canonical uppercase. Renamed=" + renamed);
            else
                RcsError.Info(ed, "Preflight: no lowercase TXT-/PNT- layers found.");
        }

        private static int AutoRenameLowercasePrefixedLayers(Database db, Transaction tr, Editor ed)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            // collect first (renaming while iterating LayerTable can be touchy)
            var renames = new List<(string oldName, string newName)>();

            foreach (ObjectId lid in lt)
            {
                LayerTableRecord ltr = tr.GetObject(lid, OpenMode.ForRead) as LayerTableRecord;
                if (ltr == null) continue;

                string name = ltr.Name ?? "";
                if (name.StartsWith("txt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("TXT-"))
                    renames.Add((name, "TXT-" + name.Substring(4)));
                else if (name.StartsWith("pnt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("PNT-"))
                    renames.Add((name, "PNT-" + name.Substring(4)));
            }

            int renamed = 0;
            foreach (var pair in renames)
            {
                if (string.Equals(pair.oldName, pair.newName, StringComparison.Ordinal)) continue;

                try
                {
                    RenameLayerWithMerge(db, tr, ed, pair.oldName, pair.newName);
                    renamed++;
                }
                catch (System.Exception ex)
                {
                    RcsError.Warn(ed, "Preflight: failed to rename layer '" + pair.oldName + "' â†’ '" + pair.newName + "'. " + ex.Message);
                }
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
                if (name.StartsWith("txt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("TXT-"))
                    bad.Add(name);
                else if (name.StartsWith("pnt-", StringComparison.OrdinalIgnoreCase) && !name.StartsWith("PNT-"))
                    bad.Add(name);
            }

            return bad;
        }

        private static void RenameLayerWithMerge(Database db, Transaction tr, Editor ed, string oldName, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) return;

            // If old doesn't exist, nothing to do
            if (!LayerExists(db, tr, oldName)) return;

            // Ensure current layer isn't the old layer
            try
            {
                LayerTableRecord cur = tr.GetObject(db.Clayer, OpenMode.ForRead) as LayerTableRecord;
                if (cur != null && string.Equals(cur.Name, oldName, StringComparison.Ordinal))
                {
                    // switch to 0 if possible
                    if (LayerExists(db, tr, "0"))
                    {
                        LayerTableRecord zero = GetLayer(db, tr, "0", OpenMode.ForRead);
                        if (zero != null) db.Clayer = zero.ObjectId;
                    }
                }
            }
            catch { /* ignore */ }

            // If target already exists, merge: move entities, then erase old
            if (LayerExists(db, tr, newName))
            {
                ReassignAllEntitiesLayer(db, tr, oldName, newName);

                // erase old layer record if possible
                LayerTableRecord oldLtr = GetLayer(db, tr, oldName, OpenMode.ForWrite);
                if (oldLtr != null && !oldLtr.IsErased)
                {
                    try { oldLtr.Erase(true); }
                    catch { /* may be in use; ignore */ }
                }

                return;
            }

            // Otherwise, rename in place
            LayerTableRecord ltrWrite = GetLayer(db, tr, oldName, OpenMode.ForWrite);
            if (ltrWrite == null) return;

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
                    Entity ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                    if (ent == null) continue;

                    if (string.Equals(ent.Layer, fromLayer, StringComparison.Ordinal))
                        ent.Layer = toLayer;
                }
            }
        }

        // ------------------------------
        // Lock canonical layers
        // ------------------------------
        private static int LockCanonicalLayers(Database db, Transaction tr, Editor ed)
        {
            int locked = 0;
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            foreach (ObjectId lid in lt)
            {
                LayerTableRecord ltr = tr.GetObject(lid, OpenMode.ForWrite) as LayerTableRecord;
                if (ltr == null) continue;

                string name = ltr.Name ?? "";
                if (string.Equals(name, "0", StringComparison.OrdinalIgnoreCase)) continue;
                if (string.Equals(name, "Defpoints", StringComparison.OrdinalIgnoreCase)) continue;

                // Only lock canonical uppercase layers
                if (name.StartsWith("PNT-", StringComparison.Ordinal) || name.StartsWith("TXT-", StringComparison.Ordinal))
                {
                    if (!ltr.IsLocked)
                    {
                        ltr.IsLocked = true;
                        locked++;
                    }
                }
            }

            if (locked > 0)
                RcsError.Info(ed, "Locked canonical layers (PNT-/TXT-). Locked=" + locked);

            return locked;
        }

        // Add this helper method anywhere in the class (preferably near similar diagnostics helpers)
        private static void DumpLayerSetterDiagnostics(object comp, string viewName, string componentHint, Editor ed, string pointStyleName = null)
        {
            if (_layerSetterDiagCount >= _layerSetterDiagMax) return;
            _layerSetterDiagCount++;

            try
            {
                string compType = comp == null ? "(null)" : comp.GetType().FullName;
                string ps = string.IsNullOrWhiteSpace(pointStyleName) ? "" : $" PointStyle='{pointStyleName}'";
                RcsError.Warn(ed, $"[DIAG] Layer setter failed:{ps} View='{viewName}', Component='{componentHint}', Type='{compType}'");
            }
            catch { }
        }

    }
}

