// RCS_DescKeySet_ImportExport_PATCHED.cs
// Civil 3D .NET (C#) - safer Description Key Set export (and scaffold for import)
// Commands:
//   RCS_EXPORT_DESC_KEYSETS
//
// Notes:
// - Uses reflection + fallbacks to support Civil 3D builds where CivilDocument.PointDescriptionKeySets is unavailable.
// - Writes CSV to C:\temp\DescKeySets_All.csv
// - Logs to C:\temp\c3doutput.txt

using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025
{
    public class RCS_DescKeySet_ImportExport
    {
        private const string LogPath = @"C:\temp\c3doutput.txt";
        private const string ExportCsv = @"C:\temp\DescKeySets_All.csv";

        [CommandMethod("RCS_EXPORT_DESC_KEYSETS")]
        public void ExportAllDescKeySets()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            try
            {
                EnsureFolder(@"C:\temp");
                Log("==== RCS_EXPORT_DESC_KEYSETS START ====");

                var civDoc = CivilApplication.ActiveDocument;

                using (var tr = doc.Database.TransactionManager.StartTransaction())
                using (var sw = new StreamWriter(ExportCsv, false))
                {
                    // Header (single CSV for all sets)
                    sw.WriteLine(string.Join(",",
                        Csv("KeySetName"),
                        Csv("Key"),
                        Csv("FullDescription"),
                        Csv("Layer"),
                        Csv("PointStyle"),
                        Csv("PointLabelStyle"),
                        Csv("BlockOrMarker"),
                        Csv("Scale"),
                        Csv("Rotation")
                    ));

                    var setIds = GetDescKeySetIdsSafe(civDoc);
                    Log($"Found DescKeySet IDs: {setIds.Count}");

                    int rowCount = 0;

                    foreach (ObjectId setId in setIds)
                    {
                        try
                        {
                            var setObj = tr.GetObject(setId, OpenMode.ForRead);
                            if (setObj == null) continue;

                            string setName = GetPropString(setObj, "Name");
                            if (string.IsNullOrWhiteSpace(setName))
                                setName = GetPropString(setObj, "DescriptionKeySetName");
                            if (string.IsNullOrWhiteSpace(setName))
                                setName = "(UnnamedSet)";

                            var keys = GetKeysEnumerableSafe(setObj);
                            int keyCount = 0;

                            foreach (var keyObj in keys)
                            {
                                keyCount++;
                                string key = FirstNonEmpty(
                                    GetPropString(keyObj, "Key"),
                                    GetPropString(keyObj, "Code"),
                                    GetPropString(keyObj, "DescriptionKey")
                                );

                                string desc = FirstNonEmpty(
                                    GetPropString(keyObj, "FullDescription"),
                                    GetPropString(keyObj, "Description"),
                                    GetPropString(keyObj, "Text")
                                );

                                string layer = FirstNonEmpty(
                                    GetPropString(keyObj, "Layer"),
                                    GetPropString(keyObj, "LayerName")
                                );

                                // These are often ObjectIds in some builds; if they come out empty, that's OK for export.
                                string ps = FirstNonEmpty(
                                    GetPropString(keyObj, "PointStyle"),
                                    GetPropString(keyObj, "PointStyleName")
                                );
                                string ls = FirstNonEmpty(
                                    GetPropString(keyObj, "PointLabelStyle"),
                                    GetPropString(keyObj, "LabelStyleName"),
                                    GetPropString(keyObj, "PointLabelStyleName")
                                );

                                string blk = FirstNonEmpty(
                                    GetPropString(keyObj, "MarkerSymbolName"),
                                    GetPropString(keyObj, "BlockName"),
                                    GetPropString(keyObj, "SymbolName"),
                                    GetPropString(keyObj, "PointSymbolBlockName")
                                );

                                string scale = FirstNonEmpty(
                                    GetPropInvariantDoubleString(keyObj, "Scale"),
                                    GetPropInvariantDoubleString(keyObj, "MarkerScale"),
                                    GetPropInvariantDoubleString(keyObj, "FixedScaleFactor")
                                );

                                string rot = FirstNonEmpty(
                                    GetPropInvariantDoubleString(keyObj, "Rotation"),
                                    GetPropInvariantDoubleString(keyObj, "MarkerRotation")
                                );

                                sw.WriteLine(string.Join(",",
                                    Csv(setName),
                                    Csv(key),
                                    Csv(desc),
                                    Csv(layer),
                                    Csv(ps),
                                    Csv(ls),
                                    Csv(blk),
                                    Csv(scale),
                                    Csv(rot)
                                ));
                                rowCount++;
                            }

                            Log($"Exported Set '{setName}' keys={keyCount}");
                        }
                        catch (System.Exception exSet)
                        {
                            Log($"ERROR exporting setId={setId}: {exSet}");
                        }
                    }

                    tr.Commit();
                    Log($"EXPORT COMPLETE. Rows={rowCount}. File={ExportCsv}");
                    ed.WriteMessage($"\nExport complete: {ExportCsv}\nRows: {rowCount}\nLog: {LogPath}");
                }

                Log("==== RCS_EXPORT_DESC_KEYSETS END ====");
            }
            catch (System.Exception ex)
            {
                Log("FATAL EXPORT: " + ex);
                ed.WriteMessage($"\nFATAL ERROR: {ex.Message}\nSee log: {LogPath}");
            }
        }

        // -------------------------
        // SAFE SET ID ACCESS (PATCH)
        // -------------------------

        private static ObjectIdCollection GetDescKeySetIdsSafe(CivilDocument civDoc)
        {
            // 1) Try civDoc.PointDescriptionKeySets
            try
            {
                var p = civDoc.GetType().GetProperty("PointDescriptionKeySets", BindingFlags.Public | BindingFlags.Instance);
                if (p != null)
                {
                    var v = p.GetValue(civDoc);
                    if (v is ObjectIdCollection col && col.Count > 0)
                        return col;
                }
            }
            catch (System.Exception ex)
            {
                Log("WARN: PointDescriptionKeySets access failed: " + ex.Message);
            }

            // 2) Try civDoc.Styles.DescriptionKeySets (some builds)
            try
            {
                var stylesProp = civDoc.GetType().GetProperty("Styles", BindingFlags.Public | BindingFlags.Instance);
                var styles = stylesProp?.GetValue(civDoc);
                if (styles != null)
                {
                    var dkProp = styles.GetType().GetProperty("DescriptionKeySets", BindingFlags.Public | BindingFlags.Instance);
                    var v = dkProp?.GetValue(styles);
                    if (v is ObjectIdCollection col2 && col2.Count > 0)
                        return col2;
                }
            }
            catch (System.Exception ex)
            {
                Log("WARN: Styles.DescriptionKeySets access failed: " + ex.Message);
            }

            // 3) Fail-safe empty
            return new ObjectIdCollection();
        }

        private static System.Collections.Generic.IEnumerable<object> GetKeysEnumerableSafe(object descKeySetObj)
        {
            if (descKeySetObj == null) yield break;

            object keysObj = null;
            try
            {
                keysObj = GetPropObject(descKeySetObj, "DescriptionKeys") ?? GetPropObject(descKeySetObj, "Keys");
            }
            catch { keysObj = null; }

            if (keysObj is System.Collections.IEnumerable en)
            {
                foreach (var item in en)
                    yield return item;
            }
        }

        // -------------------------
        // Reflection helpers
        // -------------------------

        private static object GetPropObject(object obj, string prop)
        {
            if (obj == null) return null;
            var pi = obj.GetType().GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
            if (pi == null) return null;
            return pi.GetValue(obj, null);
        }

        private static string GetPropString(object obj, string prop)
        {
            try
            {
                var v = GetPropObject(obj, prop);
                if (v == null) return "";
                return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
            }
            catch { return ""; }
        }

        private static string GetPropInvariantDoubleString(object obj, string prop)
        {
            try
            {
                var v = GetPropObject(obj, prop);
                if (v == null) return "";
                if (v is double d) return d.ToString("0.########", CultureInfo.InvariantCulture);
                if (double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture),
                    NumberStyles.Any, CultureInfo.InvariantCulture, out var dd))
                    return dd.ToString("0.########", CultureInfo.InvariantCulture);
                return "";
            }
            catch { return ""; }
        }

        private static string FirstNonEmpty(params string[] vals)
        {
            foreach (var s in vals)
            {
                if (!string.IsNullOrWhiteSpace(s)) return s.Trim();
            }
            return "";
        }

        // -------------------------
        // CSV / Logging
        // -------------------------

        private static string Csv(string s)
        {
            s ??= "";
            bool needs = s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r");
            if (!needs) return s;
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private static void EnsureFolder(string dir)
        {
            if (string.IsNullOrWhiteSpace(dir)) return;
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        }

        private static void Log(string msg)
        {
            try
            {
                EnsureFolder(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {msg}{Environment.NewLine}");
            }
            catch { /* swallow */ }
        }
    }
}
