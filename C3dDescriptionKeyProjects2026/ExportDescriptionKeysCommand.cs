using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using DBObject = Autodesk.AutoCAD.DatabaseServices.DBObject;

[assembly: CommandClass(typeof(RcsTools.ExportDescriptionKeysCommand))]

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
                    object keySetCollection = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(db);

                    if (keySetCollection == null)
                    {
                        ed.WriteMessage("\nCould not access description key sets from the active drawing database.");
                        return;
                    }

                    foreach (object keySetItem in EnumerateUnknownCollection(keySetCollection))
                    {
                        ObjectId keySetId = ToObjectId(keySetItem);
                        if (keySetId == ObjectId.Null || keySetId.IsErased)
                            continue;

                        DBObject keySetObj = tr.GetObject(keySetId, OpenMode.ForRead, false);
                        if (keySetObj == null)
                            continue;

                        string keySetName = SafeString(GetPropertyValue(keySetObj, "Name"));
                        if (string.IsNullOrWhiteSpace(keySetName))
                            keySetName = "<UnnamedKeySet>";

                        object keysObj = GetPropertyValue(keySetObj, "DescriptionKeys");
                        if (keysObj == null)
                            keysObj = GetPropertyValue(keySetObj, "Keys");
                        if (keysObj == null)
                            keysObj = GetPropertyValue(keySetObj, "Items");

                        if (keysObj == null)
                            continue;

                        foreach (object keyItem in EnumerateUnknownCollection(keysObj))
                        {
                            object descKeyObj = TryOpenDbObjectOrReturnRaw(tr, keyItem);
                            if (descKeyObj == null)
                                continue;

                            string code = SafeString(GetPropertyValue(descKeyObj, "Code"));
                            if (string.IsNullOrWhiteSpace(code))
                                code = SafeString(GetPropertyValue(descKeyObj, "Name"));
                            if (string.IsNullOrWhiteSpace(code))
                                code = SafeString(GetPropertyValue(descKeyObj, "RawDesc"));

                            string pointStyleName = string.Empty;
                            string blockName = string.Empty;

                            // Autodesk docs use StyleId on PointDescriptionKey
                            ObjectId pointStyleId = GetObjectIdProperty(descKeyObj, "StyleId");
                            if (pointStyleId == ObjectId.Null)
                                pointStyleId = GetObjectIdProperty(descKeyObj, "PointStyleId");

                            if (pointStyleId != ObjectId.Null && !pointStyleId.IsErased)
                            {
                                DBObject psObj = tr.GetObject(pointStyleId, OpenMode.ForRead, false);
                                if (psObj is PointStyle pointStyle)
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

        private static object TryOpenDbObjectOrReturnRaw(Transaction tr, object item)
        {
            try
            {
                ObjectId id = ToObjectId(item);
                if (id != ObjectId.Null && !id.IsErased)
                    return tr.GetObject(id, OpenMode.ForRead, false);

                return item;
            }
            catch
            {
                return null;
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

        private static IEnumerable<object> EnumerateUnknownCollection(object collection)
        {
            if (collection == null)
                yield break;

            if (collection is IEnumerable enumerable)
            {
                foreach (object item in enumerable)
                    yield return item;

                yield break;
            }

            object countObj = GetPropertyValue(collection, "Count");
            if (countObj is int count)
            {
                for (int i = 0; i < count; i++)
                {
                    object item = GetIndexedValue(collection, i);
                    if (item != null)
                        yield return item;
                }
            }
        }

        private static object GetIndexedValue(object obj, int index)
        {
            try
            {
                PropertyInfo indexer = obj.GetType()
                    .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .FirstOrDefault(p =>
                    {
                        ParameterInfo[] idx = p.GetIndexParameters();
                        return idx.Length == 1 && idx[0].ParameterType == typeof(int);
                    });

                if (indexer != null)
                    return indexer.GetValue(obj, new object[] { index });

                MethodInfo itemMethod = obj.GetType().GetMethod("Item", new[] { typeof(int) });
                if (itemMethod != null)
                    return itemMethod.Invoke(obj, new object[] { index });
            }
            catch
            {
            }

            return null;
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

        private static ObjectId GetObjectIdProperty(object obj, string propertyName)
        {
            try
            {
                object val = GetPropertyValue(obj, propertyName);
                if (val is ObjectId id)
                    return id;
            }
            catch
            {
            }

            return ObjectId.Null;
        }

        private static ObjectId ToObjectId(object obj)
        {
            if (obj is ObjectId id)
                return id;

            try
            {
                object value = GetPropertyValue(obj, "ObjectId");
                if (value is ObjectId oid)
                    return oid;
            }
            catch
            {
            }

            return ObjectId.Null;
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