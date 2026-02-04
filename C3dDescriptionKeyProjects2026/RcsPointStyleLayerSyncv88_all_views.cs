
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
using System.Reflection;
using Assembly = System.Reflection.Assembly;
using DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject;

namespace RCS.C3D2025
{
    /// <summary>
    /// RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv88
    /// - Maps PointStyle NAME -> DescKey CODE (case-insensitive)
    /// - Derives Marker layer from DescKey Layer (strips "PNT-" prefix; Layer0 passes through)
    /// - Derives Label layer as "txt-" + MarkerBase
    /// - Updates Plan view always; optionally tries Model/Profile/Section too (best-effort)
    ///
    /// Notes:
    /// - .NET Framework 4.8 safe (no string.Contains(StringComparison))
    /// - DescKey layer resolver supports string, ObjectId, and reflection fallbacks
    /// </summary>
    public class RcsPointStyleLayerSyncv7
    {
        private const string ForcedLayer0CsvPath = @"C:\temp\descKeys_forced_layer0.csv";

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

        [CommandMethod("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv88")]
        public static void RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv88", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;
                Database db = doc.Database;

                bool forceAll = false;         // update styles even if current Plan/Marker isn't 0
                bool ensureKeyLayers = false;  // also ensure DescKey layer exists (non-zero)
                bool setAllViews = false;      // attempt Model/Profile/Section too

                try
                {
                    PromptKeywordOptions pko = new PromptKeywordOptions("\nUpdate ALL PointStyles (ignore Layer0 gate)?");
                    pko.AllowNone = true;
                    pko.Keywords.Add("No");
                    pko.Keywords.Add("Yes");
                    pko.Keywords.Default = "No";
                    PromptResult pr = ed.GetKeywords(pko);
                    if (pr.Status == PromptStatus.OK && string.Equals(pr.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                        forceAll = true;

                    PromptKeywordOptions pko2 = new PromptKeywordOptions("\nAuto-create missing KEY layers (non-zero) too?");
                    pko2.AllowNone = true;
                    pko2.Keywords.Add("No");
                    pko2.Keywords.Add("Yes");
                    pko2.Keywords.Default = "No";
                    PromptResult pr2 = ed.GetKeywords(pko2);
                    if (pr2.Status == PromptStatus.OK && string.Equals(pr2.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                        ensureKeyLayers = true;

                    PromptKeywordOptions pko3 = new PromptKeywordOptions("\nSet Model/Profile/Section display layers too? (best-effort)");
                    pko3.AllowNone = true;
                    pko3.Keywords.Add("No");
                    pko3.Keywords.Add("Yes");
                    pko3.Keywords.Default = "No";
                    PromptResult pr3 = ed.GetKeywords(pko3);
                    if (pr3.Status == PromptStatus.OK && string.Equals(pr3.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
                        setAllViews = true;
                }
                catch
                {
                    // If prompts fail, keep defaults.
                }

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

                    List<ForcedLayer0Row> forcedRows;
                    Dictionary<string, string> descKeyNameToLayer = BuildDescKeyNameToLayerMap(tr, db, ed, out forcedRows);
                    WriteForcedLayer0Csv(forcedRows, ed);

                    if (descKeyNameToLayer.Count == 0)
                    {
                        RcsError.Warn(ed, "No Description Key code/name -> layer mappings found. Run RCS_DIAG_DESC_KEYS to inspect key layer fields.");
                        return;
                    }

                    int updated = 0, skipped = 0, noKeyMatch = 0, notLayer0 = 0, secondaryMatch = 0;

                    foreach (KeyValuePair<string, ObjectId> psPair in pointStyleNameToId)
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

                            string markerLayer = ComputeMarkerLayer(keyLayer);
                            string labelLayer = ComputeLabelLayerFromMarker(markerLayer);

                            if (ensureKeyLayers && !string.Equals(keyLayer, "0", StringComparison.OrdinalIgnoreCase))
                            {
                                try { EnsureLayer(db, tr, keyLayer); } catch { }
                            }

                            if (!IsValidLayerName(markerLayer) || !IsValidLayerName(labelLayer))
                            {
                                RcsError.Warn(ed, "Invalid computed layer(s) for '" + psName + "': keyLayer='" + keyLayer + "' marker='" + markerLayer + "' label='" + labelLayer + "'. Skipping.");
                                skipped++;
                                continue;
                            }

                            ObjectId markerLayerId = EnsureLayer(db, tr, markerLayer);
                            ObjectId labelLayerId = EnsureLayer(db, tr, labelLayer);

                            var ps = tr.GetObject(psId, OpenMode.ForWrite) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                            if (ps == null)
                            {
                                skipped++;
                                continue;
                            }

                            // Gate: only update if CURRENT Plan/Marker is Layer 0 or blank (unless FORCEALL enabled)
                            if (!forceAll)
                            {
                                string currentPlanMarker = GetCurrentComponentLayer(ps, "Plan", "Marker", ed);
                                if (!string.IsNullOrWhiteSpace(currentPlanMarker) &&
                                    !string.Equals(currentPlanMarker, "0", StringComparison.OrdinalIgnoreCase))
                                {
                                    notLayer0++;
                                    continue;
                                }
                            }

                            bool markerOk = SetPointStyleDisplayLayerAllViews(ps, "Marker", markerLayer, markerLayerId, ed, setAllViews);
                            bool labelOk = SetPointStyleDisplayLayerAllViews(ps, "Label", labelLayer, labelLayerId, ed, setAllViews);

                            if (markerOk || labelOk)
                            {
                                updated++;
                                RcsError.Info(ed, "UPDATED '" + psName + "': keyLayer='" + keyLayer +
                                                "' -> Plan" + (setAllViews ? "/Model/Profile/Section" : "") +
                                                " Marker='" + markerLayer + "', Label='" + labelLayer + "'" +
                                                (usedSecondary ? " (SECONDARY key match '" + matchedKey + "')" : ""));
                            }
                            else
                            {
                                skipped++;
                                RcsError.Warn(ed, "No compatible layer setter found for '" + psName + "'.");
                            }
                        }
                        catch (System.Exception exPs)
                        {
                            skipped++;
                            RcsError.Warn(ed, "PointStyle update failed: '" + psName + "' :: " + exPs.Message);
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage(
                        "\nRCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS: Updated=" + updated +
                        ", Skipped=" + skipped +
                        ", NoKeyMatch=" + noKeyMatch +
                        ", NotLayer0=" + notLayer0 +
                        ", SecondaryMatch=" + secondaryMatch +
                        ". See " + RcsError.LogPath +
                        "\nForced-to-0 CSV: " + ForcedLayer0CsvPath + "\n");
                }
            });
        }

        [CommandMethod("RCS_DIAG_POINTSTYLES")]
        public static void RCS_DIAG_POINTSTYLES()
        {
            RcsError.RunCommandSafe("RCS_DIAG_POINTSTYLES", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;

                using (doc.LockDocument())
                {
                    CivilDocument civDoc = CivilApplication.ActiveDocument;
                    if (civDoc == null) throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    Database db = doc.Database;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        int i = 0;
                        foreach (ObjectId id in civDoc.Styles.PointStyles)
                        {
                            i++;
                            try
                            {
                                if (id.IsNull || id.IsErased) continue;
                                var ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                                if (ps == null) continue;
                                RcsError.Info(ed, "POINTSTYLE #" + i + ": Name='" + GetStyleName(ps) + "' Id=" + id.ToString());
                            }
                            catch (System.Exception ex)
                            {
                                RcsError.Warn(ed, "POINTSTYLE #" + i + ": FAILED open/read: " + ex.Message);
                            }
                        }
                        RcsError.Info(ed, "POINTSTYLES_DIAG: iterated=" + i);
                        tr.Commit();
                    }
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
                                cleaned = CleanDescKeyCode(GetStyleName(key));

                            string layer = ResolveKeyLayerName(tr, key);
                            RcsError.Info(ed, "KEY_SAMPLE #" + (dumped + 1) +
                                             ": Code='" + cleaned +
                                             "' RawCode='" + rawCode +
                                             "' Name='" + GetStyleName(key) +
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

                string setName = GetStyleName(set);

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
                        code = CleanDescKeyCode(GetStyleName(key));

                    if (string.IsNullOrWhiteSpace(code))
                    {
                        missingCode++;
                        continue;
                    }

                    string resolvedRaw = ResolveKeyLayerName(tr, key);
                    string layer = NormalizeLayerName(resolvedRaw);

                    if (string.Equals(layer, "0", StringComparison.OrdinalIgnoreCase))
                    {
                        forcedLayer0++;
                        forcedRows.Add(new ForcedLayer0Row
                        {
                            SetName = setName,
                            Code = code,
                            RawCode = rawCode,
                            KeyName = GetStyleName(key),
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

                // Scan any property containing "Layer"
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

                // Scan fields too
                FieldInfo[] fields = key.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (FieldInfo fi in fields)
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

                foreach (string pn in new[] { "Layer", "LayerStyle", "PointLayer", "PointLayerStyle" })
                {
                    PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (pi != null)
                    {
                        object v = null;
                        try { v = pi.GetValue(val, null); } catch { v = null; }
                        string ln = ExtractLayerNameFromAny(tr, v);
                        if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
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
            if (string.Equals(s, "NONE", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "NULL", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "(NULL)", StringComparison.OrdinalIgnoreCase)) return "0";
            if (string.Equals(s, "0", StringComparison.OrdinalIgnoreCase)) return "0";
            return s;
        }

        // ---------------------- RULES ----------------------

        private static string ComputeMarkerLayer(string keyLayer)
        {
            string markerLayer = NormalizeLayerName((keyLayer ?? string.Empty).Trim());
            if (string.Equals(markerLayer, "0", StringComparison.OrdinalIgnoreCase)) return "0";

            // Strip common "PNT-" prefix from layer names like "PNT-TREES"
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
            return "txt-" + labelBase;
        }

        // ---------------------- DISPLAY SETTERS ----------------------

        private static bool SetPointStyleDisplayLayerAllViews(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed,
            bool setAllViews)
        {
            bool planOk = TrySetDisplayLayerRobust(ps, "Plan", componentHint, layerName, layerId, ed);

            if (!setAllViews) return planOk;

            // best-effort for other view representations
            try { TrySetDisplayLayerRobust(ps, "Model", componentHint, layerName, layerId, ed); } catch { }
            try { TrySetDisplayLayerRobust(ps, "Profile", componentHint, layerName, layerId, ed); } catch { }
            try { TrySetDisplayLayerRobust(ps, "Section", componentHint, layerName, layerId, ed); } catch { }

            return planOk;
        }

        private static bool TrySetDisplayLayerRobust(
            Autodesk.Civil.DatabaseServices.Styles.PointStyle ps,
            string viewName,
            string componentHint,
            string layerName,
            ObjectId layerId,
            Editor ed)
        {
            object displayStyle = GetDisplayStyleObject(ps, viewName, componentHint);
            if (displayStyle == null) return false;
            return TrySetLayerOnDisplayStyle(displayStyle, layerName, layerId);
        }

        private static object GetDisplayStyleObject(object pointStyle, string viewName, string componentHint)
        {
            if (pointStyle == null) return null;

            string v = (viewName ?? "").Trim();
            string c = (componentHint ?? "").Trim();

            string[] viewHints = (string.Equals(v, "Model", StringComparison.OrdinalIgnoreCase))
                ? new[] { "Model", "Plan" }
                : new[] { v };

            string[] compHints = new[] { c, "Point" + c, c + "Style" };

            // 1) Property-based retrieval (no args)
            try
            {
                Type t = pointStyle.GetType();
                PropertyInfo[] props = t.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

                foreach (string vh in viewHints)
                {
                    foreach (string ch in compHints)
                    {
                        foreach (PropertyInfo pi in props)
                        {
                            if (pi == null || pi.Name == null) continue;
                            if (!pi.CanRead) continue;

                            string pn = pi.Name;
                            bool looks = (pn.IndexOf("DisplayStyle", StringComparison.OrdinalIgnoreCase) >= 0) ||
                                         (pn.IndexOf("Style", StringComparison.OrdinalIgnoreCase) >= 0);
                            if (!looks) continue;

                            bool viewMatch = string.IsNullOrWhiteSpace(vh) || (pn.IndexOf(vh, StringComparison.OrdinalIgnoreCase) >= 0);
                            bool compMatch = string.IsNullOrWhiteSpace(ch) || (pn.IndexOf(ch, StringComparison.OrdinalIgnoreCase) >= 0);
                            if (!viewMatch || !compMatch) continue;

                            object res = null;
                            try { res = pi.GetValue(pointStyle, null); } catch { res = null; }
                            if (res != null) return res;
                        }
                    }
                }
            }
            catch { }

            // 2) Method-based retrieval (overloads)
            try
            {
                Type t = pointStyle.GetType();
                MethodInfo[] methods = t.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

                foreach (string vh in viewHints)
                {
                    foreach (MethodInfo mi in methods)
                    {
                        if (mi == null) continue;
                        string mn = mi.Name ?? "";
                        if (mn.IndexOf("GetDisplayStyle", StringComparison.OrdinalIgnoreCase) < 0) continue;

                        if (!string.IsNullOrWhiteSpace(vh) && mn.IndexOf(vh, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            if (!mn.Equals("GetDisplayStyle", StringComparison.OrdinalIgnoreCase)) continue;
                        }

                        ParameterInfo[] ps = mi.GetParameters();
                        if (ps == null) continue;

                        try
                        {
                            if (ps.Length == 1)
                            {
                                object a0;
                                if (!TryBuildArg(ps[0].ParameterType, componentHint, out a0)) continue;
                                object res = mi.Invoke(pointStyle, new object[] { a0 });
                                if (res != null) return res;
                            }
                            else if (ps.Length == 2)
                            {
                                object a0, a1;
                                if (!TryBuildArg(ps[0].ParameterType, vh, out a0)) continue;
                                if (!TryBuildArg(ps[1].ParameterType, componentHint, out a1)) continue;
                                object res = mi.Invoke(pointStyle, new object[] { a0, a1 });
                                if (res != null) return res;
                            }
                        }
                        catch
                        {
                            // try next
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private static bool TryBuildArg(Type paramType, string hint, out object arg)
        {
            arg = null;

            if (paramType == typeof(string))
            {
                arg = hint;
                return true;
            }

            if (paramType != null && paramType.IsEnum)
            {
                try { arg = Enum.Parse(paramType, hint, true); return true; } catch { }

                string h = (hint ?? "").Trim();

                if (string.Equals(h, "Plan", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(h, "Profile", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(h, "Section", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(h, "Model", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (string cand in new[] { h, h + "View", "PlanView", "ModelView" })
                    {
                        try { arg = Enum.Parse(paramType, cand, true); return true; } catch { }
                    }
                }

                if (string.Equals(h, "Marker", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(h, "Label", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (string cand in new[] { h, h + "Style", "Point" + h })
                    {
                        try { arg = Enum.Parse(paramType, cand, true); return true; } catch { }
                    }
                }

                return false;
            }

            return false;
        }

        private static bool TrySetLayerOnDisplayStyle(object displayStyle, string layerName, ObjectId layerId)
        {
            if (displayStyle == null) return false;

            Type t = displayStyle.GetType();

            foreach (string pn in new[] { "LayerId", "LayerObjectId" })
            {
                PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (pi != null && pi.CanWrite && pi.PropertyType == typeof(ObjectId))
                {
                    try { pi.SetValue(displayStyle, layerId, null); return true; } catch { }
                }
            }

            foreach (string pn in new[] { "Layer", "LayerName" })
            {
                PropertyInfo pi = t.GetProperty(pn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (pi != null && pi.CanWrite)
                {
                    try
                    {
                        if (pi.PropertyType == typeof(string)) { pi.SetValue(displayStyle, layerName, null); return true; }
                        if (pi.PropertyType == typeof(ObjectId)) { pi.SetValue(displayStyle, layerId, null); return true; }
                    }
                    catch { }
                }
            }

            foreach (string mn in new[] { "SetLayerId", "SetLayer", "SetLayerName" })
            {
                MethodInfo mi = t.GetMethod(mn, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                if (mi == null) continue;

                ParameterInfo[] ps = mi.GetParameters();
                if (ps == null || ps.Length != 1) continue;

                try
                {
                    if (ps[0].ParameterType == typeof(ObjectId)) { mi.Invoke(displayStyle, new object[] { layerId }); return true; }
                    if (ps[0].ParameterType == typeof(string)) { mi.Invoke(displayStyle, new object[] { layerName }); return true; }
                }
                catch { }
            }

            return false;
        }

        // ---------------------- READ CURRENT LAYER (for Layer0 gating) ----------------------

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
                RcsError.Warn(ed, "Read current layer failed: " + GetStyleName(ps) + " " + viewName + "/" + componentHint + " :: " + ex.Message);
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

        private static Type ResolvePointDisplayEnumType(Autodesk.Civil.DatabaseServices.Styles.PointStyle ps)
        {
            Assembly asm = ps.GetType().Assembly;
            return asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointDisplayStyleType")
                ?? asm.GetType("Autodesk.Civil.DatabaseServices.Styles.PointStyleDisplayStyleType");
        }

        // ---------------------- CSV ----------------------

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
                        foreach (ForcedLayer0Row r in rows)
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

        // ---------------------- MATCH RULE ----------------------

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
                catch (System.Exception ex)
                {
                    openFail++;
                    RcsError.Warn(ed, "PointStyle open failed (id=" + id.ToString() + "): " + ex.Message);
                    continue;
                }

                if (ps == null) continue;

                string name = GetStyleName(ps);
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

            RcsError.Info(ed, "PointStyles mapped by name: " + added +
                             " (scanned=" + scanned +
                             ", nullOrErased=" + nullOrErased +
                             ", openFail=" + openFail +
                             ", unnamed=" + unnamed + ")");
            return map;
        }

        private static ObjectId EnsureLayer(Database db, Transaction tr, string layerName)
        {
            try
            {
                string ln = string.IsNullOrWhiteSpace(layerName) ? "0" : layerName.Trim();
                if (!IsValidLayerName(ln)) ln = "0";

                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (lt.Has(ln))
                    return lt[ln];

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
            if (o == null) return "(null)";

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
    }
}
