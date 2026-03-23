// RCS_PointStyle_DescKeyLayerApply.cs
// Civil 3D / AutoCAD .NET 2025+
//
// Command:
//   RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES
//
// Summary:
//   - Map DescKey Code -> Layer across ALL Desc Key Sets.
//   - For each PointStyle, if PointStyle.Name matches DescKey Code (case-insensitive):
//       * Apply layer to PointStyle display styles (Marker + Label) in Plan/Model/Profile.
//       * Label layer rule (exact): if layer starts with TXT/PNT/SYM => trim first 4 chars.
//   - Create missing layers.
//   - Output CSV + log.
//
// Output:
//   CSV: C:\temp\PointStyle_DescKeyLayer_Results.csv
//   Log: C:\temp\c3d_pointstyle_descKeyLayer.log

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;

namespace RCS.C3D2025
{
    public class RCS_PointStyle_DescKeyLayerApply
    {
        private const string CsvPath = @"C:\temp\PointStyle_DescKeyLayer_Results.csv";
        private const string LogPath = @"C:\temp\c3d_pointstyle_descKeyLayer.log";

        // Priority rule for duplicate codes across multiple DescKey sets:
        //   true  -> first match wins (keeps earliest found)
        //   false -> last match wins (overwrites prior)
        private static readonly bool FirstMatchWins = true;

        [CommandMethod("RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES")]
        public void ApplyDescKeyLayersToPointStyles()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;
            var db = doc?.Database;

            Directory.CreateDirectory(@"C:\temp");
            Log("=== START RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES ===");

            if (doc == null || db == null)
            {
                Log("FATAL: ActiveDocument/Database is null.");
                return;
            }

            CivilDocument civDoc;
            try
            {
                civDoc = CivilApplication.ActiveDocument;
                if (civDoc == null)
                {
                    Log("FATAL: CivilApplication.ActiveDocument is null.");
                    return;
                }
            }
            catch (System.Exception ex)
            {
                Log("FATAL: Failed to get CivilApplication.ActiveDocument.", ex);
                return;
            }

            int scanned = 0, matched = 0, updated = 0, skipped = 0, failed = 0, layersCreated = 0;
            int layerSetOps = 0;

            using (var sw = new StreamWriter(CsvPath, false))
            {
                sw.WriteLine("PointStyleName,MatchedDescKeyCode,DescKeyLayer,LabelLayerComputed,View,Part,BeforeLayer,AfterLayer,Status,Message");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    Dictionary<string, string> descKeyLayerMap;
                    try
                    {
                        // Build DescKey Code -> Layer map
                        descKeyLayerMap = BuildDescKeyCodeToLayerMap(civDoc, tr);
                        Log($"DescKey map built. Entries={descKeyLayerMap.Count} FirstMatchWins={FirstMatchWins}");
                    }
                    catch (System.Exception ex)
                    {
                        Log("FATAL: Failed to build description key map.", ex);
                        return;
                    }

                    // LayerTable open once
                    var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                    foreach (ObjectId psId in civDoc.Styles.PointStyles)
                    {
                        scanned++;

                        try
                        {
                            var ps = tr.GetObject(psId, OpenMode.ForWrite, false) as PointStyle;
                            if (ps == null)
                            {
                                failed++;
                                sw.WriteLine($",,,,{Csv("")},{Csv("")},{Csv("")},{Csv("")},FAIL,{Csv("Null PointStyle object")}");
                                continue;
                            }

                            string psName = ps.Name ?? "";
                            string key = psName.Trim();

                            if (!descKeyLayerMap.TryGetValue(key, out string markerLayer) || string.IsNullOrWhiteSpace(markerLayer))
                            {
                                skipped++;
                                sw.WriteLine($"{Csv(psName)},,,,{Csv("")},{Csv("")},{Csv("")},{Csv("")},SKIP,{Csv("No DescKey Code match")}");
                                continue;
                            }

                            matched++;

                            string labelLayer = ComputeLabelLayer(markerLayer);

                            // Ensure layers exist (create if missing)
                            layersCreated += EnsureLayerExists(db, tr, lt, markerLayer);
                            if (!string.Equals(labelLayer, markerLayer, StringComparison.OrdinalIgnoreCase))
                                layersCreated += EnsureLayerExists(db, tr, lt, labelLayer);

                            // Apply to display styles (Marker + Label) for Plan, Model, Profile
                            layerSetOps += ApplyForAllViews(ps, markerLayer, labelLayer, sw, psName, key, markerLayer, labelLayer);

                            updated++;
                        }
                        catch (System.Exception exStyle)
                        {
                            failed++;
                            Log($"ERROR: PointStyle loop failed at scanned={scanned}", exStyle);
                            sw.WriteLine($",,,,{Csv("")},{Csv("")},{Csv("")},{Csv("")},FAIL,{Csv(exStyle.Message)}");
                        }
                    }

                    tr.Commit();
                }
            }

            ed?.WriteMessage(
                $"\nRCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES complete." +
                $"\nScanned: {scanned} | Matched: {matched} | Updated: {updated} | Skipped: {skipped} | Failed: {failed} | LayersCreated: {layersCreated}" +
                $"\nLayerSetOps: {layerSetOps}" +
                $"\nCSV: {CsvPath}" +
                $"\nLog: {LogPath}\n"
            );

            Log($"SUMMARY scanned={scanned}, matched={matched}, updated={updated}, skipped={skipped}, failed={failed}, layersCreated={layersCreated}, layerSetOps={layerSetOps}");
            Log("=== END RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES ===");
        }

        // --------------------- DescKey map ---------------------

        private static Dictionary<string, string> BuildDescKeyCodeToLayerMap(CivilDocument civDoc, Transaction tr)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc?.Database;
            if (db == null) return map;

            ObjectIdCollection setIds;
            try
            {
                // Fix: Use PointDescriptionKeySetCollection as an enumerable, not ObjectIdCollection
                var keySetCollection = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);
                setIds = new ObjectIdCollection();
                foreach (ObjectId id in keySetCollection)
                {
                    setIds.Add(id);
                }
            }
            catch (System.Exception ex)
            {
                Log("WARN: GetPointDescriptionKeySets failed.", ex);
                return map;
            }

            if (setIds == null || setIds.Count == 0)
                return map;

            foreach (ObjectId setId in setIds)
            {
                PointDescriptionKeySet set = null;

                try
                {
                    set = tr.GetObject(setId, OpenMode.ForRead, false) as PointDescriptionKeySet;
                    if (set == null) continue;

                    // Iterate keys safely
                    foreach (ObjectId keyId in set.GetPointDescriptionKeyIds())
                    {
                        try
                        {
                            var keyObj = tr.GetObject(keyId, OpenMode.ForRead, false) as PointDescriptionKey;
                            if (keyObj == null) continue;

                            string code = (keyObj.Code ?? "").Trim();
                            string layer = (GetDescKeyLayerName(keyObj, tr) ?? "").Trim();

                            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(layer))
                                continue;

                            if (FirstMatchWins)
                            {
                                if (!map.ContainsKey(code))
                                    map[code] = layer;
                            }
                            else
                            {
                                map[code] = layer; // last match wins
                            }
                        }
                        catch (System.Exception exKey)
                        {
                            Log($"WARN: Failed reading key in set '{set.Name}'.", exKey);
                        }
                    }
                }
                catch (System.Exception exSet)
                {
                    Log($"WARN: Failed reading description key set '{set?.Name ?? "<unknown>"}'.", exSet);
                }
            }

            return map;
        }

        // --------------------- Apply view layers ---------------------

        private static int ApplyForAllViews(PointStyle ps, string markerLayer, string labelLayer, StreamWriter sw,
            string psName, string matchedCode, string descKeyLayer, string labelLayerComputed)
        {
            int sets = 0;

            // Plan
            sets += ApplyOne(sw, psName, matchedCode, descKeyLayer, labelLayerComputed, "Plan", "Marker", ps.GetDisplayStylePlan(PointDisplayStyleType.Marker), markerLayer);
            sets += ApplyOne(sw, psName, matchedCode, descKeyLayer, labelLayerComputed, "Plan", "Label", ps.GetDisplayStylePlan(PointDisplayStyleType.Label), labelLayer);

            // Model
            sets += ApplyOne(sw, psName, matchedCode, descKeyLayer, labelLayerComputed, "Model", "Marker", ps.GetDisplayStyleModel(PointDisplayStyleType.Marker), markerLayer);
            sets += ApplyOne(sw, psName, matchedCode, descKeyLayer, labelLayerComputed, "Model", "Label", ps.GetDisplayStyleModel(PointDisplayStyleType.Label), labelLayer);

            // Profile
            sets += ApplyOne(sw, psName, matchedCode, descKeyLayer, labelLayerComputed, "Profile", "Marker", ps.GetDisplayStyleProfile(), markerLayer);
            sets += ApplyOne(sw, psName, matchedCode, descKeyLayer, labelLayerComputed, "Profile", "Label", ps.GetDisplayStyleProfile(), labelLayer);

            // Note: PointStyle does not have a distinct "Section" view style in the .NET API. 
            // Points in sections are usually projections that use Plan/Model styles or projection styles.

            return sets;
        }

        private static int ApplyOne(StreamWriter sw, string psName, string matchedCode, string descKeyLayer, string labelLayerComputed,
            string view, string part, DisplayStyle ds, string targetLayer)
        {
            if (ds == null)
            {
                sw.WriteLine($"{Csv(psName)},{Csv(matchedCode)},{Csv(descKeyLayer)},{Csv(labelLayerComputed)},{Csv(view)},{Csv(part)},,,WARN,{Csv("DisplayStyle is null (API returned null)")}");
                return 0;
            }

            string before = "";
            string after = "";
            int sets = 0;

            try { before = ds.Layer ?? ""; } catch { before = ""; }

            try
            {
                if (!string.Equals(before, targetLayer, StringComparison.OrdinalIgnoreCase))
                {
                    ds.Layer = targetLayer;
                    sets++;
                }

                // Ensure visibility is TRUE
                if (!ds.Visible)
                {
                    ds.Visible = true;
                    sets++;
                }
            }
            catch (System.Exception ex)
            {
                Log($"WARN: Failed setting DisplayStyle.Layer='{targetLayer}' ({view}/{part}) for PointStyle='{psName}'.", ex);
                sw.WriteLine($"{Csv(psName)},{Csv(matchedCode)},{Csv(descKeyLayer)},{Csv(labelLayerComputed)},{Csv(view)},{Csv(part)},{Csv(before)},,FAIL,{Csv(ex.Message)}");
                return 0;
            }

            try { after = ds.Layer ?? ""; } catch { after = ""; }

            sw.WriteLine($"{Csv(psName)},{Csv(matchedCode)},{Csv(descKeyLayer)},{Csv(labelLayerComputed)},{Csv(view)},{Csv(part)},{Csv(before)},{Csv(after)},OK,{Csv($"Sets={sets}")}");
            return sets;
        }

        // --------------------- Label rule ---------------------

        private static string ComputeLabelLayer(string layer)
        {
            if (string.IsNullOrWhiteSpace(layer)) return "";

            if (layer.StartsWith("TXT", StringComparison.OrdinalIgnoreCase) ||
                layer.StartsWith("PNT", StringComparison.OrdinalIgnoreCase) ||
                layer.StartsWith("SYM", StringComparison.OrdinalIgnoreCase))
            {
                if (layer.Length > 4) return layer.Substring(4);
                return layer; // too short to trim safely
            }

            return layer;
        }

        // --------------------- DescKey layer getter ---------------------

        private static string GetDescKeyLayerName(PointDescriptionKey keyObj, Transaction tr)
        {
            if (keyObj == null) return "";

            try
            {
                // Use robust API property if possible
                ObjectId layerId = keyObj.LayerId;

                if (layerId != ObjectId.Null)
                {
                    var ltr = tr.GetObject(layerId, OpenMode.ForRead, false) as LayerTableRecord;
                    return ltr?.Name ?? "";
                }
            }
            catch
            {
                // Fallback or ignore if API version mismatch
            }

            return "";
        }

        // --------------------- Layer utilities ---------------------

        private static int EnsureLayerExists(Database db, Transaction tr, LayerTable lt, string layerName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(layerName)) return 0;
                if (lt.Has(layerName)) return 0;

                // Upgrade only if needed
                if (!lt.IsWriteEnabled)
                    lt.UpgradeOpen();

                var ltr = new LayerTableRecord { Name = layerName };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);

                return 1;
            }
            catch (System.Exception ex)
            {
                Log($"WARN: EnsureLayerExists failed for '{layerName}'.", ex);
                return 0;
            }
        }

        // --------------------- Logging helpers ---------------------

        private static void Log(string msg, System.Exception ex = null)
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
            catch { /* never throw from logger */ }
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