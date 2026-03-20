// RCS_PointStyle_LayerPolicy_FIXED_NoLayerId.cs
// Civil 3D .NET 2025+
// PATCH NOTES:
// - PointStyle / LabelStyle do NOT reliably expose LayerId in Civil 3D APIs.
// - Styles store Layer as a *string* (layer name). We read/write it via reflection safely.
// - When you need a LayerTableRecord ObjectId, resolve by layer name in the LayerTable.
//
// Commands:
//   RCS_EXPORT_POINTSTYLE_LAYERS
//   RCS_SET_POINT_MARKER_AND_LABEL_BY_OBJECT_LAYER_V1  (renamed to avoid duplicate command registration)
//
// Output:
//   CSV: C:\temp\PointStyles_And_LabelStyles.csv
//   Log: C:\temp\c3doutput.txt

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
using Autodesk.Civil.DatabaseServices.Styles;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace RCS.C3D2025
{
    public class RCS_PointStyleLayerPolicy
    {
        private const string CsvPath = @"C:\temp\PointStyles_And_LabelStyles.csv";
        private const string LogPath = @"C:\temp\c3doutput.txt";

        [CommandMethod("RCS_EXPORT_POINTSTYLE_LAYERS")]
        public void ExportPointStyleLayers()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            Directory.CreateDirectory(@"C:\temp");

            var civDoc = CivilApplication.ActiveDocument;

            try
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                using (var sw = new StreamWriter(CsvPath, false))
                {
                    // Include both layer name + resolved layerId (from layer table) for convenience
                    sw.WriteLine("PointStyleName,PointStyleLayer,PointStyleLayerId,MarkerType,MarkerSymbolName,LabelStyleName,LabelStyleLayer,LabelStyleLayerId");

                    foreach (ObjectId id in civDoc.Styles.PointStyles)
                    {
                        var ps = tr.GetObject(id, OpenMode.ForRead) as PointStyle;
                        if (ps == null) continue;

                        string psLayerName = SafeGet(ps, "Layer");                 // <-- style layer is string
                        ObjectId psLayerId = ResolveLayerId(doc.Database, tr, psLayerName);

                        string markerType = SafeGet(ps, "MarkerType");
                        string markerSymbolName = SafeGet(ps, "MarkerSymbolName");

                        // Replace the problematic line with correct string interpolation and ternary usage
                        sw.WriteLine($"{Csv(ps.Name)},{Csv(psLayerName)},{Csv(psLayerId == ObjectId.Null ? "" : Csv(psLayerId.ToString()))},{Csv(markerType)},{Csv(markerSymbolName)},,,");
                    }

                    foreach (ObjectId labelId in civDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)
                    {
                        var ls = tr.GetObject(labelId, OpenMode.ForRead) as LabelStyle;
                        if (ls == null) continue;

                        string lsLayerName = SafeGet(ls, "Layer");                 // <-- style layer is string
                        ObjectId lsLayerId = ResolveLayerId(doc.Database, tr, lsLayerName);

                        sw.WriteLine($",,,,,{Csv(ls.Name)},{Csv(lsLayerName)},{Csv(lsLayerId == ObjectId.Null ? "" : lsLayerId.ToString())}");
                    }
    

                        tr.Commit();
                        }

                        ed.WriteMessage($"\nExported to {CsvPath}");
                    }
            catch (Exception ex)
            {
                Log("EXPORT ERROR: " + ex);
                ed.WriteMessage($"\nEXPORT ERROR: {ex.Message}\nSee log: {LogPath}");
            }
        }

        [CommandMethod("RCS_SET_POINT_MARKER_AND_LABEL_BY_OBJECT_LAYER_V1")]
        public void ApplyByObjectLayerPolicy()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            Directory.CreateDirectory(@"C:\temp");

            var civDoc = CivilApplication.ActiveDocument;

            try
            {
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // Neutralize style layers to "0" so the CogoPoint object's layer can control display.
                    foreach (ObjectId id in civDoc.Styles.PointStyles)
                    {
                        var ps = tr.GetObject(id, OpenMode.ForWrite) as PointStyle;
                        if (ps == null) continue;

                        SafeSet(ps, "Layer", "0");
                    }

                    foreach (ObjectId labelId in civDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)
                    {
                        var ls = tr.GetObject(labelId, OpenMode.ForWrite) as LabelStyle;
                        if (ls == null) continue;

                        SafeSet(ls, "Layer", "0");
                    }

                    tr.Commit();
                }

                ed.WriteMessage("\nPoint style and label styles set to layer 0 (via style Layer name).");
            }
            catch (Exception ex)
            {
                Log("POLICY ERROR: " + ex);
                ed.WriteMessage($"\nPOLICY ERROR: {ex.Message}\nSee log: {LogPath}");
            }
        }

        // ---------------------------
        // Layer helpers
        // ---------------------------

        /// <summary>
        /// Styles store layer as a name. If you need the ObjectId, resolve it from LayerTable.
        /// </summary>
        private static ObjectId ResolveLayerId(Database db, Transaction tr, string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return ObjectId.Null;

            var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName)) return ObjectId.Null;
            return lt[layerName];
        }

        // ---------------------------
        // Reflection helpers
        // ---------------------------

        private static string SafeGet(object obj, string prop)
        {
            try
            {
                if (obj == null) return "";
                var pi = obj.GetType().GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return "";
                var v = pi.GetValue(obj);
                return v?.ToString() ?? "";
            }
            catch { return ""; }
        }

        private static void SafeSet(object obj, string prop, object value)
        {
            try
            {
                if (obj == null) return;
                var pi = obj.GetType().GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null || !pi.CanWrite) return;
                pi.SetValue(obj, value);
            }
            catch { }
        }

        private static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(@"C:\temp");
                File.AppendAllText(LogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {msg}{Environment.NewLine}");
            }
            catch { }
        }
    }
}
