using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject;

// [assembly: CommandClass(typeof(RcsTools.ExportDescriptionKeysCommand))]

namespace RcsTools
{
    public class ExportDescriptionKeysCommand
    {
        [CommandMethod("RCS_EXPORT_DESCKEY_CODE_BLOCKS")]
        public void ExportDescriptionKeyCodesAndBlockNames()
        {
            Document acadDoc = AcApp.DocumentManager.MdiActiveDocument;
            if (acadDoc == null)
                return;

            Editor ed = acadDoc.Editor;
            Database db = acadDoc.Database;

            try
            {
                CivilDocument civilDoc = CivilApplication.ActiveDocument;
                if (civilDoc == null)
                {
                    ed.WriteMessage("\nNo active Civil 3D document found.");
                    return;
                }

                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string outPath = Path.Combine(desktop, "DescriptionKey_Code_BlockName.txt");

                List<string> rows = new List<string>
                {
                    "DescriptionKeySet\tCode\tPointStyle\tBlockName"
                };

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    PointDescriptionKeySetCollection keySets =
                        PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);

                    if (keySets == null || keySets.Count == 0)
                    {
                        ed.WriteMessage("\nNo description key sets were found in this drawing.");
                        return;
                    }

                    foreach (ObjectId keySetId in keySets)
                    {
                        if (keySetId == ObjectId.Null || keySetId.IsErased)
                            continue;

                        PointDescriptionKeySet keySet =
                            tr.GetObject(keySetId, OpenMode.ForRead, false) as PointDescriptionKeySet;

                        if (keySet == null)
                            continue;

                        string keySetName = string.IsNullOrWhiteSpace(keySet.Name)
                            ? "<UnnamedKeySet>"
                            : keySet.Name;

                        ObjectIdCollection keyIds = keySet.GetPointDescriptionKeyIds();
                        if (keyIds == null || keyIds.Count == 0)
                            continue;

                        foreach (ObjectId keyId in keyIds)
                        {
                            if (keyId == ObjectId.Null || keyId.IsErased)
                                continue;

                            PointDescriptionKey key =
                                tr.GetObject(keyId, OpenMode.ForRead, false) as PointDescriptionKey;

                            if (key == null)
                                continue;

                            string code = SafeString(key.Code);
                            if (string.IsNullOrWhiteSpace(code))
                                code = SafeString(key.DisplayName);

                            string pointStyleName = string.Empty;
                            string blockName = string.Empty;

                            ObjectId pointStyleId = key.StyleId;
                            if (pointStyleId != ObjectId.Null && !pointStyleId.IsErased)
                            {
                                PointStyle pointStyle =
                                    tr.GetObject(pointStyleId, OpenMode.ForRead, false) as PointStyle;

                                if (pointStyle != null)
                                {
                                    pointStyleName = SafeString(pointStyle.Name);
                                    blockName = TryGetPointStyleBlockName(pointStyle);
                                }
                            }

                            rows.Add(
                                $"{Escape(keySetName)}\t{Escape(code)}\t{Escape(pointStyleName)}\t{Escape(blockName)}");
                        }
                    }

                    tr.Commit();
                }

                File.WriteAllLines(outPath, rows, Encoding.UTF8);
                ed.WriteMessage($"\nExport complete: {outPath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private static string TryGetPointStyleBlockName(PointStyle pointStyle)
        {
            try
            {
                string blockName = TryGetByReflection(
                    pointStyle,
                    "MarkerBlockName",
                    "BlockName",
                    "MarkerSymbolName",
                    "SymbolName");

                if (!string.IsNullOrWhiteSpace(blockName))
                    return blockName;

                object markerStyle = GetPropertyValue(pointStyle, "MarkerStyle");
                if (markerStyle != null)
                {
                    blockName = TryGetByReflection(
                        markerStyle,
                        "BlockName",
                        "MarkerBlockName",
                        "SymbolName",
                        "MarkerSymbolName");

                    if (!string.IsNullOrWhiteSpace(blockName))
                        return blockName;
                }

                object displayStyle3d = GetPropertyValue(pointStyle, "DisplayStyle3d");
                if (displayStyle3d != null)
                {
                    blockName = TryGetByReflection(
                        displayStyle3d,
                        "BlockName",
                        "SymbolName");

                    if (!string.IsNullOrWhiteSpace(blockName))
                        return blockName;
                }

                object displayStylePlan = GetPropertyValue(pointStyle, "DisplayStylePlan");
                if (displayStylePlan != null)
                {
                    blockName = TryGetByReflection(
                        displayStylePlan,
                        "BlockName",
                        "SymbolName");

                    if (!string.IsNullOrWhiteSpace(blockName))
                        return blockName;
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static object GetPropertyValue(object obj, string propertyName)
        {
            if (obj == null || string.IsNullOrWhiteSpace(propertyName))
                return null;

            try
            {
                PropertyInfo pi = obj.GetType().GetProperty(
                    propertyName,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

                return pi != null ? pi.GetValue(obj) : null;
            }
            catch
            {
                return null;
            }
        }

        private static string TryGetByReflection(object obj, params string[] propertyNames)
        {
            if (obj == null)
                return string.Empty;

            foreach (string name in propertyNames)
            {
                try
                {
                    object value = GetPropertyValue(obj, name);
                    string s = SafeString(value);
                    if (!string.IsNullOrWhiteSpace(s))
                        return s;
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        private static string SafeString(object value)
        {
            return value?.ToString()?.Trim() ?? string.Empty;
        }

        private static string Escape(string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;

            return s.Replace("\t", " ")
                    .Replace("\r", " ")
                    .Replace("\n", " ");
        }
    }
}