using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
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
    public class RcsPointStyleLayerSyncv7
    {
        // --- Output files ---
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


        [CommandMethod("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv83")]
        public static void RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS()
        {
            RcsError.RunCommandSafe("RCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYSv83", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;
                Database db = doc.Database;

                // One-time override switches (per run)
                bool forceAll = false;            // update styles even if current layer is not 0
                bool ensureKeyLayers = false;     // also ensure keyLayer exists (in addition to marker/label layers)

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
                }
                catch
                {
                    // If prompts fail for any reason, keep defaults (safe behavior)
                }


                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    CivilDocument civDoc = CivilApplication.ActiveDocument;
                    if (civDoc == null) throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    Dictionary<string, ObjectId> pointStyleNameToId = BuildPointStyleNameMap(tr, civDoc, ed);

                    // Diagnostic: dump first 50 PointStyle names so we know what we're matching
                    try
                    {
                        int n = 0;
                        foreach (var kv in pointStyleNameToId)
                        {
                            if (n >= 50) break;
                            RcsError.Info(ed, "POINTSTYLE_SAMPLE #" + (n + 1) + ": '" + kv.Key + "'");
                            n++;
                        }
                    }
                    catch { }

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
                        try
                        {
                        string psName = psPair.Key;
                        ObjectId psId = psPair.Value;

                        string keyLayer;
                        string matchedKey;
                        bool usedSecondary;

                        if (!TryGetDescKeyLayerWithSecondaryRule(descKeyNameToLayer, psName, out keyLayer, out matchedKey, out usedSecondary))
                        {
                            noKeyMatch++;
                            try
                            {
                                // Debug: show what we attempted
                                string psNormDbg = (psName ?? string.Empty).Trim().Replace("*", "");
                                string psStripDbg = psNormDbg.Length > 2 ? psNormDbg.Substring(2).TrimStart('-', '_', ' ', '.') : "";
                                RcsError.Info(ed, "NO_KEY_MATCH: PointStyle='" + psName + "' norm='" + psNormDbg + "' strip2='" + psStripDbg + "'");
                            }
                            catch { }
                            continue;
                        }

                        if (usedSecondary) secondaryMatch++;
                        try
                        {
                            RcsError.Info(ed, "MATCH: PointStyle='" + psName + "' -> matchedKey='" + matchedKey + "'" + (usedSecondary ? " (SECONDARY)" : "") + " keyLayer='" + keyLayer + "'");
                        }
                        catch { }



                        string markerLayer;
                        markerLayer = ComputeMarkerLayer(keyLayer);
                        string labelLayer = ComputeLabelLayerFromMarker(markerLayer);

                        // Optional: ensure the DescKey's own layer exists too (non-zero only)
                        if (ensureKeyLayers && !string.Equals(keyLayer, "0", StringComparison.OrdinalIgnoreCase))
                        {
                            try { EnsureLayer(db, tr, keyLayer); } catch { /* ignore */ }
                        }


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

                        // Diagnostic: read current layer before any changes
                        string currentPlanMarkerDbg = GetCurrentComponentLayer(ps, "Plan", "Marker", ed);
                        try { RcsError.Info(ed, "CURRENT_LAYER: PointStyle='" + psName + "' Plan/Marker='" + currentPlanMarkerDbg + "' ForceAll=" + (forceAll ? "Yes" : "No")); } catch { }

// Gate: only update if CURRENT Plan Marker is Layer 0 or blank (unless FORCEALL enabled)
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

                        bool markerOk = SetPointStyleDisplayLayerAllViews(ps, "Marker", markerLayer, markerLayerId, ed);
                        bool labelOk = SetPointStyleDisplayLayerAllViews(ps, "Label", labelLayer, labelLayerId, ed);

                        try
                        {
                            RcsError.Info(ed, "SET_RESULT: PointStyle='" + psName + "' markerOk=" + (markerOk ? "Yes" : "No") + " labelOk=" + (labelOk ? "Yes" : "No"));
                        }
                        catch { }


                        if (markerOk || labelOk)
                        {
                            updated++;
                            RcsError.Info(ed, "UPDATED '" + psName + "': matchedKey='" + matchedKey + "' keyLayer='" + keyLayer + "' -> marker='" + markerLayer + "', label='" + labelLayer + "'" + (usedSecondary ? " (SECONDARY)" : ""));
                        }
                        else
                        {
                            skipped++;
                            RcsError.Warn(ed, "No compatible setter found for '" + psName + "'.");
                        }
                        }
                        catch (System.Exception exPs)
                        {
                            skipped++;
                            RcsError.Warn(ed, "PointStyle update failed: '" + psPair.Key + "' :: " + exPs.Message);
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage("\nRCS_SYNC_POINTSTYLE_LAYERS_FROM_DESC_KEYS: Updated=" + updated +
                                    ", Skipped=" + skipped +
                                    ", NoKeyMatch=" + noKeyMatch +
                                    ", NotLayer0=" + notLayer0 +
                                    ". See " + RcsError.LogPath + "\nForced-to-0 CSV: " + ForcedLayer0CsvPath);
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

                    int total = 0;
                    try
                    {
                        var col = civDoc.Styles.PointStyles;
                        var pi = col.GetType().GetProperty("Count");
                        if (pi != null)
                        {
                            object v = pi.GetValue(col, null);
                            if (v is int) total = (int)v;
                        }
                    }
                    catch { }

                    RcsError.Info(ed, "POINTSTYLES_DIAG: starting. (CountMaybe=" + total + ")");

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
                                RcsError.Info(ed, "POINTSTYLE #" + i + ": Name='" + SafeName(ps) + "' Id=" + id.ToString());
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
                    if (string.Equals(layer, "0", StringComparison.OrdinalIgnoreCase))
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

        // Resolve layer name from DescriptionKey (supports string or ObjectId properties)

        // Best-effort extraction of a layer name from ANY value coming back from Civil properties.
        // Handles: string, ObjectId, LayerTableRecord, and objects that expose Name/LayerName/LayerId/Layer properties.
        private static string ExtractLayerNameFromAny(Transaction tr, object val)
        {
            try
            {
                if (val == null) return string.Empty;

                // Direct string
                if (val is string)
                    return NormalizeLayerName((string)val);

                // ObjectId -> layer table record
                if (val is ObjectId)
                {
                    ObjectId oid = (ObjectId)val;
                    string ln = LayerNameFromId(tr, oid);
                    return string.IsNullOrWhiteSpace(ln) ? string.Empty : NormalizeLayerName(ln);
                }

                // LayerTableRecord
                LayerTableRecord ltr = val as LayerTableRecord;
                if (ltr != null)
                    return NormalizeLayerName(ltr.Name);

                // If it's a DBObject with an ObjectId we can resolve
                Autodesk.AutoCAD.DatabaseServices.DBObject dbo = val as Autodesk.AutoCAD.DatabaseServices.DBObject;
                if (dbo != null)
                {
                    string ln = LayerNameFromId(tr, dbo.ObjectId);
                    if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                }

                // Reflection fallback: Name, LayerName, Layer, LayerId
                Type t = val.GetType();

                // 1) Name or LayerName
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

                // 2) LayerId / Id-like object id
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

                // 3) If there's an embedded object that might have Name/LayerName
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
            catch
            {
                // swallow
            }

            return string.Empty;
        }

        private static string ResolveKeyLayerName(Transaction tr, PointDescriptionKey key)
        {
            try
            {
                // 1) Known names first (string/ObjectId/other)
                string[] known = new[] { "Layer", "LayerName", "PointLayer", "PointLayerId", "LayerId" };

                foreach (string prop in known)
                {
                    // string fast-path (existing behavior)
                    string s = TryGetStringProperty(key, prop);
                    if (!string.IsNullOrWhiteSpace(s)) return NormalizeLayerName(s);

                    // ObjectId fast-path (existing behavior)
                    ObjectId oid = TryGetObjectIdProperty(key, prop);
                    string name = LayerNameFromId(tr, oid);
                    if (!string.IsNullOrWhiteSpace(name)) return NormalizeLayerName(name);

                    // Any-type fallback
                    PropertyInfo piAny = key.GetType().GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase);
                    if (piAny != null)
                    {
                        object v = null;
                        try { v = piAny.GetValue(key, null); } catch { v = null; }
                        string ln = ExtractLayerNameFromAny(tr, v);
                        if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                    }
                }

                // 2) Scan ANY property containing "Layer" (public + nonpublic)
                PropertyInfo[] props = key.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (PropertyInfo pi in props)
                {
                    if (pi == null) continue;
                    if (pi.Name == null) continue;
                    if (pi.Name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    object v = null;
                    try { v = pi.GetValue(key, null); } catch { v = null; }

                    string ln = ExtractLayerNameFromAny(tr, v);
                    if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                }

                // 3) Scan fields too (some Civil builds store layer info in fields)
                var fields = key.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                foreach (var fi in fields)
                {
                    if (fi == null || fi.Name == null) continue;
                    if (fi.Name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    object v = null;
                    try { v = fi.GetValue(key); } catch { v = null; }

                    string ln = ExtractLayerNameFromAny(tr, v);
                    if (!string.IsNullOrWhiteSpace(ln)) return NormalizeLayerName(ln);
                }
            }
            catch
            {
                // ignore
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


        private static string NormalizeLayerName(string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return "0";

            string s = layerName.Trim();

            // Common "no layer override" tokens seen in some templates/builds
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

            PropertyInfo pLayerName = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerName != null && pLayerName.CanWrite && pLayerName.PropertyType == typeof(string))
            {
                pLayerName.SetValue(displayObj, layerName, null);
                return true;
            }

            PropertyInfo pLayer = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayer != null && pLayer.CanWrite && pLayer.PropertyType == typeof(ObjectId))
            {
                pLayer.SetValue(displayObj, layerId, null);
                return true;
            }

            PropertyInfo pLayerId = t.GetProperty("LayerId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
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

            PropertyInfo pLayerName = t.GetProperty("LayerName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayerName != null && pLayerName.CanWrite && pLayerName.PropertyType == typeof(string))
            {
                pLayerName.SetValue(displayObj, layerName, null);
                return true;
            }

            PropertyInfo pLayer = t.GetProperty("Layer", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (pLayer != null && pLayer.CanWrite && pLayer.PropertyType == typeof(ObjectId))
            {
                pLayer.SetValue(displayObj, layerId, null);
                return true;
            }

            PropertyInfo pLayerId = t.GetProperty("LayerId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
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


        // Secondary match rule:
        // If no initial DescKey match for a PointStyle name, and the PointStyle name starts with any of:
        // BC,CP,EL,FC,ML,GS,RR,SD,SL,SN,SS,TP,TL,TR,TS,TV,WA,WL
        // then remove the first two characters and attempt the match again.
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

            // Normalize point style name for matching
            string psNorm = pointStyleName.Trim();
            psNorm = psNorm.Replace("*", "").Trim();

            // 1) direct match
            if (descKeyNameToLayer.TryGetValue(psNorm, out keyLayer))
            {
                matchedKey = psNorm;
                return true;
            }

            // 2) secondary match by stripping a 2-char prefix
            if (psNorm.Length > 2)
            {
                string ps = psNorm;
                if (ps.Length > 2)
                {
                    string prefix = ps.Substring(0, 2).ToUpperInvariant();
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
                            string stripped = ps.Substring(2);
                            stripped = stripped.TrimStart('-', '_', ' ', '.');
                            if (descKeyNameToLayer.TryGetValue(stripped, out keyLayer))
                            {
                                matchedKey = stripped;
                                usedSecondary = true;
                                return true;
                            }
                            break;
                    }
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

            // Best-effort: log apparent collection count
            try
            {
                object col = civDoc.Styles.PointStyles;
                var pi = col.GetType().GetProperty("Count");
                if (pi != null)
                {
                    object v = pi.GetValue(col, null);
                    if (v is int) RcsError.Info(ed, "PointStyles collection Count=" + ((int)v));
                }
            }
            catch { }

            foreach (ObjectId id in civDoc.Styles.PointStyles)
            {
                scanned++;

                if (id.IsNull || id.IsErased)
                {
                    nullOrErased++;
                    continue;
                }

                Autodesk.Civil.DatabaseServices.Styles.PointStyle ps = null;
                try
                {
                    ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                }
                catch (System.Exception ex)
                {
                    openFail++;
                    RcsError.Warn(ed, "PointStyle open failed (id=" + id.ToString() + "): " + ex.Message);
                    continue;
                }

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

                if (!IsValidLayerName(ln))
                    ln = "0";

                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (lt.Has(ln))
                    return lt[ln];

                // Layer 0 should always exist; if we can't create, fall back to 0
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

        
        private static string GetStyleName(object o)
        {
            if (o == null) return "(null)";

            // 1) Civil StyleBase (most Civil styles inherit this)
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

            // 2) Reflection on public + nonpublic properties
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
    }
}
