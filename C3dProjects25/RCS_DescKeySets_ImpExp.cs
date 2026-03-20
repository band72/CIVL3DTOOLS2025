using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;

//[assembly: CommandClass(typeof(RCS.C3D2025.RCS_DescKeySet_ImpExp))]

namespace RCS.C3D2025
{
    public class RCS_DescKeySet_ImpExp
    {
        private const string CsvPath = @"C:\temp\DescKeySets_Master.csv";
        private const string LogPath = @"C:\temp\DescKey_Log.txt";

        // ===================================================================================
        // COMMAND 1: EXPORT (Fixed Crash on Deleted Layers)
        // ===================================================================================
        [CommandMethod("RCS_EXPORT_DESC_KEYSETSV2")]
        public void ExportDescKeys()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            CivilDocument civilDoc = CivilApplication.ActiveDocument;

            try
            {
                EnsureFolder(@"C:\temp");
                Log("=== EXPORT V2 STARTED ===");

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                using (StreamWriter sw = new StreamWriter(CsvPath, false, Encoding.UTF8))
                {
                    // HEADER
                    sw.WriteLine("Set_Name,Code,Format,Layer,PointStyle,LabelStyle,ScaleParam,RotateParam,MarkerBlock");

                    // 1. GET SETS (Robust Collection Access)
                    PointDescriptionKeySetCollection keySets = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(doc.Database);
                    Log($"Found {keySets.Count} sets.");

                    int totalKeys = 0;

                    foreach (ObjectId setId in keySets)
                    {
                        var setObj = tr.GetObject(setId, OpenMode.ForRead) as PointDescriptionKeySet;
                        if (setObj == null) continue;

                        string setName = setObj.Name;

                        // 2. GET KEYS
                        foreach (ObjectId keyId in setObj.GetPointDescriptionKeyIds())
                        {
                            var key = tr.GetObject(keyId, OpenMode.ForRead) as PointDescriptionKey;
                            if (key == null) continue;

                            // 3. EXTRACT PROPERTIES (Safe Wrappers)
                            string code = key.Code;
                            string format = key.Format;

                            // CRITICAL FIX: Access .LayerId inside a try-catch to prevent crashes
                            string layerName = SafeGetLayerName(tr, key);

                            string ptStyle = ResolveStyleName(tr, key.StyleId);

                            // FIX: Add missing lblStyle assignment
                            string lblStyle = ResolveStyleName(tr, key.LabelStyleId);

                            // Bool Properties
                            string scale = key.ApplyScaleParameter ? "TRUE" : "FALSE";
                            string rot = key.ApplyMarkerRotationParameter ? "TRUE" : "FALSE";

                            // Marker (often not directly exposed on key, explicit null check)
                            string marker = "";

                            // ... then use lblStyle in the WriteLine call:
                            sw.WriteLine($"{Csv(setName)},{Csv(code)},{Csv(format)},{Csv(layerName)},{Csv(ptStyle)},{Csv(lblStyle)},{Csv(scale)},{Csv(rot)},{Csv(marker)}");
                            totalKeys++;
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nSuccess: {totalKeys} keys exported to {CsvPath}");
                }
            }
            catch (System.Exception ex)
            {
                string err = $"\nFATAL ERROR: {ex.Message}\nTrace: {ex.StackTrace}";
                ed.WriteMessage(err);
                Log(err);
            }
        }

        // ===================================================================================
        // COMMAND 2: IMPORT (Fixed Property Assignments)
        // ===================================================================================
        [CommandMethod("RCS_IMPORT_DESC_KEYSETSV2")]
        public void ImportDescKeys()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            CivilDocument civilDoc = CivilApplication.ActiveDocument;

            if (!File.Exists(CsvPath))
            {
                ed.WriteMessage($"\nError: File not found: {CsvPath}");
                return;
            }

            try
            {
                Log("=== IMPORT V2 STARTED ===");

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var lines = File.ReadAllLines(CsvPath);
                    var dataLines = lines.Skip(1).Where(l => !string.IsNullOrWhiteSpace(l));

                    var lt = tr.GetObject(doc.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                    PointDescriptionKeySetCollection keySets = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(doc.Database);

                    int updatedCount = 0;
                    int createdCount = 0;

                    foreach (string line in dataLines)
                    {
                        string[] cols = ParseCsvLine(line);
                        if (cols.Length < 6) continue;

                        string setName = cols[0];
                        string code = cols[1];
                        string format = cols[2];
                        string layer = cols[3];
                        string ptStyle = cols[4];
                        string lblStyle = cols[5];
                        bool useScale = cols[6].Trim().ToUpper() == "TRUE";
                        bool useRot = cols[7].Trim().ToUpper() == "TRUE";

                        // 1. Get/Create Set
                        ObjectId setId;
                        if (keySets.Contains(setName))
                        {
                            setId = keySets[setName];
                        }
                        else
                        {
                            // Create new set if missing
                            setId = keySets.Add(setName);
                            var newSetObj = tr.GetObject(setId, OpenMode.ForWrite) as PointDescriptionKeySet;
                            newSetObj.Name = setName;
                        }

                        var setObj = tr.GetObject(setId, OpenMode.ForWrite) as PointDescriptionKeySet;

                        // 2. Get/Create Key
                        ObjectId keyId = ObjectId.Null;

                        // Iterate explicitly to find key by Code
                        foreach (ObjectId existingId in setObj.GetPointDescriptionKeyIds())
                        {
                            var k = tr.GetObject(existingId, OpenMode.ForRead) as PointDescriptionKey;
                            if (k.Code.Equals(code, StringComparison.OrdinalIgnoreCase))
                            {
                                keyId = existingId;
                                break;
                            }
                        }

                        if (keyId == ObjectId.Null)
                        {
                            keyId = setObj.Add(code);
                            createdCount++;
                        }
                        else updatedCount++;

                        // 3. Update Properties
                        var keyObj = tr.GetObject(keyId, OpenMode.ForWrite) as PointDescriptionKey;

                        keyObj.Format = format;

                        if (lt.Has(layer)) keyObj.LayerId = lt[layer];

                        // Style assignments
                        ObjectId pStyleId = GetStyleId(civilDoc, ptStyle, "Point");
                        if (pStyleId != ObjectId.Null) keyObj.StyleId = pStyleId;

                        ObjectId lStyleId = GetStyleId(civilDoc, lblStyle, "Label");
                        if (lStyleId != ObjectId.Null) keyObj.LabelStyleId = lStyleId;

                        // FIX: Assign BOOLEAN properties, not integers
                        keyObj.ApplyScaleParameter = useScale;
                        keyObj.ApplyMarkerRotationParameter = useRot;
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nImport Complete.\nCreated: {createdCount}\nUpdated: {updatedCount}");
                }
            }
            catch (System.Exception ex)
            {
                Log($"IMPORT ERROR: {ex.Message}");
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        // ===================================================================================
        // HELPERS
        // ===================================================================================

        // --- THE CRASH FIX ---
        private static string SafeGetLayerName(Transaction tr, PointDescriptionKey key)
        {
            try
            {
                // Accessing .LayerId throws 'eKeyNotFound' if the layer is deleted.
                // We must try-catch the PROPERTY ACCESS itself.
                ObjectId id = key.LayerId;

                if (id.IsNull) return "0";
                var l = tr.GetObject(id, OpenMode.ForRead) as LayerTableRecord;
                return l?.Name ?? "0";
            }
            catch
            {
                return "0"; // Default to layer 0 if broken
            }
        }

        private static string ResolveStyleName(Transaction tr, ObjectId styleId)
        {
            if (styleId.IsNull) return "<default>";
            try
            {
                var s = tr.GetObject(styleId, OpenMode.ForRead) as StyleBase;
                return s?.Name ?? "<default>";
            }
            catch { return "<default>"; }
        }

        private static ObjectId GetStyleId(CivilDocument doc, string name, string type)
        {
            if (string.IsNullOrEmpty(name) || name == "<default>") return ObjectId.Null;
            try
            {
                if (type == "Point")
                    return doc.Styles.PointStyles.Contains(name) ? doc.Styles.PointStyles[name] : ObjectId.Null;
                if (type == "Label")
                    return doc.Styles.LabelStyles.PointLabelStyles.LabelStyles.Contains(name) ? doc.Styles.LabelStyles.PointLabelStyles.LabelStyles[name] : ObjectId.Null;
            }
            catch { }
            return ObjectId.Null;
        }

        private static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.Contains(",") || s.Contains("\"")) return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static string[] ParseCsvLine(string line)
        {
            List<string> result = new List<string>();
            bool inQuotes = false;
            StringBuilder sb = new StringBuilder();
            foreach (char c in line)
            {
                if (c == '\"') inQuotes = !inQuotes;
                else if (c == ',' && !inQuotes) { result.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result.ToArray();
        }

        private static void Log(string msg)
        {
            try { File.AppendAllText(LogPath, $"{DateTime.Now}: {msg}{Environment.NewLine}"); } catch { }
        }

        private static void EnsureFolder(string p) { if (!Directory.Exists(p)) Directory.CreateDirectory(p); }
    }
}