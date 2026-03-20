// RCS_PointStyle_ByLayerComponents.cs
// Civil 3D .NET 2025+
// Goal:
//   Force Point Style marker components AND Point Label Style components to ByLayer
//   (color, linetype, lineweight) using best-effort reflection across Civil 3D versions.
//
// Commands:
//   RCS_FORCE_MARKER_LABEL_COMPONENTS_BYLAYER   -> best-effort: sets style component display to ByLayer
//   RCS_SET_POINT_MARKER_AND_LABEL_BY_OBJECT_LAYER -> neutralizes style layers to "0" + forces ByLayer displays
//   RCS_EXPORT_POINTSTYLE_LAYERS                -> exports key style info to CSV
//
// Output:
//   CSV: C:\temp\PointStyles_And_LabelStyles.csv
//   Log: C:\temp\c3doutput.txt

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Color = Autodesk.AutoCAD.Colors.Color;

namespace RCS.C3D2025
{
    public class RCS_PointStyle_ByLayerComponents
    {

[CommandMethod("RCS_FORCE_POINTSTYLE_ALL_VIEWS_BYLAYER")]
public void RcsForceAllViewsByLayer()
{
        var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
    var ed = doc.Editor;
    EnsureFolder(@"C:\temp");
    Log("==== RCS_FORCE_POINTSTYLE_ALL_VIEWS_BYLAYER START ====");

    var civDoc = CivilApplication.ActiveDocument;

    using (var tr = doc.Database.TransactionManager.StartTransaction())
    {
        var byLayerColor = Color.FromColorIndex(ColorMethod.ByLayer, 256);
        var byLayerLinetypeId = GetLinetypeId(doc.Database, tr, "ByLayer");

        int psTouched = 0, psDisplayTouched = 0;
        foreach (ObjectId id in civDoc.Styles.PointStyles)
        {
            var ps = tr.GetObject(id, OpenMode.ForWrite) as PointStyle;
            if (ps == null) continue;

            psTouched++;
            psDisplayTouched += ForcePointStyleAllViews(ps, "0", byLayerColor, byLayerLinetypeId);
                    psDisplayTouched += ForcePointStyleAllViews(ps, "0", byLayerColor, byLayerLinetypeId);
                    psDisplayTouched += ForceByLayerOnObjectGraph(ps, byLayerColor, byLayerLinetypeId);
        }

        int lsTouched = 0, lsDisplayTouched = 0;
        foreach (ObjectId id in civDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)
        {
            var ls = tr.GetObject(id, OpenMode.ForWrite) as LabelStyle;
            if (ls == null) continue;

            lsTouched++;
            lsDisplayTouched += ForceByLayerOnObjectGraph(ls, byLayerColor, byLayerLinetypeId);
        }

        tr.Commit();

        ed.WriteMessage($"\nAll-views ByLayer applied.\nPointStyles: {psTouched} (sets: {psDisplayTouched})" +
                        $"\nPoint Label Styles: {lsTouched} (sets: {lsDisplayTouched})\nLog: {LogPath}");
        Log($"SUMMARY: PS={psTouched}, PSsets={psDisplayTouched}, LS={lsTouched}, LSsets={lsDisplayTouched}");
    }

    Log("==== RCS_FORCE_POINTSTYLE_ALL_VIEWS_BYLAYER END ====");
}

        private const string CsvPath = @"C:\temp\PointStyles_And_LabelStyles.csv";
        private const string LogPath = @"C:\temp\c3doutput.txt";

        [CommandMethod("RCS_EXPORT_POINTSTYLE_LAYERSV2")]
        public void ExportPointStyleLayers()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            EnsureFolder(@"C:\temp");

            var civDoc = CivilApplication.ActiveDocument;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            using (var sw = new StreamWriter(CsvPath, false))
            {
                sw.WriteLine("PointStyleName,PointStyleLayer,MarkerType,MarkerSymbolName,LabelStyleName,LabelStyleLayer");

                foreach (ObjectId id in civDoc.Styles.PointStyles)
                {
                    var ps = tr.GetObject(id, OpenMode.ForRead) as PointStyle;
                    if (ps == null) continue;

                    string psLayer = SafeGet(ps, "Layer");
                    string markerType = SafeGet(ps, "MarkerType");
                    string markerSymbol = SafeGet(ps, "MarkerSymbolName");

                    sw.WriteLine($"{Csv(ps.Name)},{Csv(psLayer)},{Csv(markerType)},{Csv(markerSymbol)},,");
                }

                foreach (ObjectId id in civDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)     
                {
                    var ls = tr.GetObject(id, OpenMode.ForRead) as LabelStyle;
                    if (ls == null) continue;

                    string lsLayer = SafeGet(ls, "Layer");
                    sw.WriteLine($",,,,{Csv(ls.Name)},{Csv(lsLayer)}");
                }

                tr.Commit();
            }

            ed.WriteMessage($"\nExported to {CsvPath}");
        }

        [CommandMethod("RCS_FORCE_MARKER_LABEL_COMPONENTS_BYLAYER")]
        public void ForceMarkerAndLabelComponentsByLayer()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            EnsureFolder(@"C:\temp");
            Log("==== RCS_FORCE_MARKER_LABEL_COMPONENTS_BYLAYER START ====");

            var civDoc = CivilApplication.ActiveDocument;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var byLayerColor = Color.FromColorIndex(ColorMethod.ByLayer, 256);
                var byLayerLinetypeId = GetLinetypeId(doc.Database, tr, "ByLayer");

                int psTouched = 0, psDisplayTouched = 0;
                foreach (ObjectId id in civDoc.Styles.PointStyles)
                {
                    var ps = tr.GetObject(id, OpenMode.ForWrite) as PointStyle;
                    if (ps == null) continue;

                    psTouched++;
                    psDisplayTouched += ForcePointStyleAllViews(ps, "0", byLayerColor, byLayerLinetypeId);
                    psDisplayTouched += ForcePointStyleAllViews(ps, "0", byLayerColor, byLayerLinetypeId);
                    psDisplayTouched += ForceByLayerOnObjectGraph(ps, byLayerColor, byLayerLinetypeId);
                }

                int lsTouched = 0, lsDisplayTouched = 0;
                foreach (ObjectId id in civDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)
                {
                    var ls = tr.GetObject(id, OpenMode.ForWrite) as LabelStyle;
                    if (ls == null) continue;

                    lsTouched++;
                    lsDisplayTouched += ForceByLayerOnObjectGraph(ls, byLayerColor, byLayerLinetypeId);
                }

                tr.Commit();

                ed.WriteMessage($"\nDone.\nPointStyles scanned: {psTouched} (display fields updated: {psDisplayTouched})" +
                                $"\nPoint Label Styles scanned: {lsTouched} (display fields updated: {lsDisplayTouched})" +
                                $"\nLog: {LogPath}");
                Log($"SUMMARY: PS={psTouched}, PSdisp={psDisplayTouched}, LS={lsTouched}, LSdisp={lsDisplayTouched}");
            }

            Log("==== RCS_FORCE_MARKER_LABEL_COMPONENTS_BYLAYER END ====");
        }

        /// <summary>
        /// Policy command: neutralize style layers to "0" AND force all display components to ByLayer.
        /// This makes the CogoPoint object's Layer the controlling layer.
        /// </summary>
        [CommandMethod("RCS_SET_POINT_MARKER_AND_LABEL_BY_OBJECT_LAYER")]
        public void ApplyByObjectLayerPolicy()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            EnsureFolder(@"C:\temp");
            Log("==== RCS_SET_POINT_MARKER_AND_LABEL_BY_OBJECT_LAYER START ====");

            var civDoc = CivilApplication.ActiveDocument;

            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                var byLayerColor = Color.FromColorIndex(ColorMethod.ByLayer, 256);
                var byLayerLinetypeId = GetLinetypeId(doc.Database, tr, "ByLayer");

                int psTouched = 0, psDisplayTouched = 0;
                foreach (ObjectId id in civDoc.Styles.PointStyles)
                {
                    var ps = tr.GetObject(id, OpenMode.ForWrite) as PointStyle;
                    if (ps == null) continue;

                    psTouched++;
                    SafeSet(ps, "Layer", "0");
                    psDisplayTouched += ForcePointStyleAllViews(ps, "0", byLayerColor, byLayerLinetypeId);
                    psDisplayTouched += ForcePointStyleAllViews(ps, "0", byLayerColor, byLayerLinetypeId);
                    psDisplayTouched += ForceByLayerOnObjectGraph(ps, byLayerColor, byLayerLinetypeId);
                }

                int lsTouched = 0, lsDisplayTouched = 0;
                foreach (ObjectId id in civDoc.Styles.LabelStyles.PointLabelStyles.LabelStyles)
                {
                    var ls = tr.GetObject(id, OpenMode.ForWrite) as LabelStyle;
                    if (ls == null) continue;

                    lsTouched++;
                    SafeSet(ls, "Layer", "0");
                    lsDisplayTouched += ForceByLayerOnObjectGraph(ls, byLayerColor, byLayerLinetypeId);
                }

                tr.Commit();

                ed.WriteMessage($"\nPolicy applied.\nPointStyles touched: {psTouched} (display updated: {psDisplayTouched})" +
                                $"\nPoint Label Styles touched: {lsTouched} (display updated: {lsDisplayTouched})" +
                                $"\nLog: {LogPath}");

                Log($"SUMMARY: PS={psTouched}, PSdisp={psDisplayTouched}, LS={lsTouched}, LSdisp={lsDisplayTouched}");
            }

            Log("==== RCS_SET_POINT_MARKER_AND_LABEL_BY_OBJECT_LAYER END ====");
        }

        // ------------------------
        // Core: Best-effort ByLayer setter across object graphs
        // ------------------------

        /// <summary>
        /// Walks an object's public instance properties up to a limited depth and tries to set
        /// Color/Linetype/Lineweight (and common synonyms) to ByLayer wherever found.
        /// Returns count of successful property sets.
        /// </summary>
        private static int ForceByLayerOnObjectGraph(object root, Color byLayerColor, ObjectId byLayerLinetypeId)
        {
            if (root == null) return 0;

            var visited = new HashSet<object>(ReferenceEqualityComparer.Instance);
            var q = new Queue<(object obj, int depth)>();
            q.Enqueue((root, 0));
            visited.Add(root);

            int sets = 0;
            const int MaxDepth = 3;
            const int MaxNodes = 250;

            while (q.Count > 0 && visited.Count < MaxNodes)
            {
                var (obj, depth) = q.Dequeue();
                if (obj == null) continue;

                // 1) Attempt display fields on this object
                sets += TryForceByLayerOnSingleObject(obj, byLayerColor, byLayerLinetypeId);

                // 2) Traverse deeper if allowed
                if (depth >= MaxDepth) continue;

                foreach (var child in EnumerateChildObjects(obj))
                {
                    if (child == null) continue;
                    if (visited.Contains(child)) continue;

                    visited.Add(child);
                    q.Enqueue((child, depth + 1));
                }
            }

            return sets;
        }

        private static IEnumerable<object> EnumerateChildObjects(object obj)
        {
            // We only traverse likely "component" objects: DisplayStyle, Components, MarkerDisplayStyle, etc.
            var t = obj.GetType();
            var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var p in props)
            {
                if (!p.CanRead) continue;

                // Avoid huge graphs: skip primitive/string
                if (p.PropertyType.IsPrimitive || p.PropertyType.IsEnum || p.PropertyType == typeof(string))
                    continue;

                string name = p.Name ?? "";

                // Prefer known component-ish properties
                bool likely =
                    name.IndexOf("Display", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Plan", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Model", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Profile", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Section", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Component", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Style", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Marker", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Symbol", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    name.IndexOf("Leader", StringComparison.OrdinalIgnoreCase) >= 0;

                if (!likely) continue;

                object v = null;
                try { v = p.GetValue(obj); } catch { v = null; }
                if (v == null) continue;

                // If enumerable, yield each item
                if (v is System.Collections.IEnumerable en && !(v is Entity) && !(v is DBObject))
                {
                    foreach (var item in en)
                    {
                        if (item == null) continue;
                        // avoid accidental traversal into database objects via enumerables
                        if (item is ObjectId) continue;
                        yield return item;
                    }
                }
                else
                {
                    // Avoid traversing actual AutoCAD DBObjects/entities here; styles are OK.
                    yield return v;
                }
            }
        }

        private static int TryForceByLayerOnSingleObject(object obj, Color byLayerColor, ObjectId byLayerLinetypeId)
        {
            int sets = 0;
            
// --- PATCH: force display ON (fixes "turned off" components) ---
sets += TrySetBool(obj, "Visible", true);
sets += TrySetBool(obj, "IsVisible", true);
sets += TrySetBool(obj, "Display", true);
sets += TrySetBool(obj, "Show", true);
sets += TrySetBool(obj, "Enabled", true);

var t = obj.GetType();

            // Color-like
            sets += TrySetProperty(obj, "Color", byLayerColor);
            sets += TrySetProperty(obj, "Colour", byLayerColor);

            // Linetype-like: could be ObjectId, string, or enum
            sets += TrySetLinetype(obj, "Linetype", byLayerLinetypeId);
            sets += TrySetLinetype(obj, "LineType", byLayerLinetypeId);
            sets += TrySetProperty(obj, "LinetypeName", "ByLayer");
            sets += TrySetProperty(obj, "LineTypeName", "ByLayer");

            // Lineweight-like
            sets += TrySetLineweight(obj, "Lineweight");
            sets += TrySetLineweight(obj, "LineWeight");

            // Some Civil3D display styles use "UseObjectLayer" / "UseByLayer" flags
            sets += TrySetProperty(obj, "UseObjectLayer", true);
            sets += TrySetProperty(obj, "UseByLayer", true);

            return sets;
        }

        
// ---------------------------
// PATCH (Explicit): PointStyle per-view DisplayStyle accessors
// PointStyle does NOT expose DisplayStylePlan/Model/Profile/Section as properties in your build.
// It exposes methods:
//   GetDisplayStylePlan/Model/Profile/Section(PointDisplayStyleType type)
// We explicitly touch Marker + Label display styles in Plan/Model/Profile/Section.
// DisplayStyle.Layer is a string layer name.
// ---------------------------

private static int ForcePointStyleAllViews(PointStyle ps, string layerName, Color byLayerColor, ObjectId byLayerLinetypeId)
{
    int sets = 0;

    // Style-level layer (if exposed as a string)
    sets += TrySetString(ps, "Layer", layerName);
    sets += TrySetString(ps, "LayerName", layerName);

    // Marker + Label display styles in each view
    sets += ForceDisplayStyle(ps.GetDisplayStylePlan(PointDisplayStyleType.Marker), layerName, byLayerColor, byLayerLinetypeId);
    sets += ForceDisplayStyle(ps.GetDisplayStylePlan(PointDisplayStyleType.Label), layerName, byLayerColor, byLayerLinetypeId);

    sets += ForceDisplayStyle(ps.GetDisplayStyleModel(PointDisplayStyleType.Marker), layerName, byLayerColor, byLayerLinetypeId);
    sets += ForceDisplayStyle(ps.GetDisplayStyleModel(PointDisplayStyleType.Label), layerName, byLayerColor, byLayerLinetypeId);
    
    sets += ForceDisplayStyle(ps.GetDisplayStyleProfile(), layerName, byLayerColor, byLayerLinetypeId);
    sets += ForceDisplayStyle(ps.GetDisplayStyleSection(), layerName, byLayerColor, byLayerLinetypeId);
    
    return sets;
}

private static int ForceDisplayStyle(DisplayStyle ds, string layerName, Color byLayerColor, ObjectId byLayerLinetypeId)
{
    if (ds == null) return 0;

    int sets = 0;

    // Layer name
    try { ds.Layer = layerName; sets++; } catch { }

    // Turn ON
    sets += TrySetBool(ds, "Visible", true);
    sets += TrySetBool(ds, "IsVisible", true);
    sets += TrySetBool(ds, "Display", true);
    sets += TrySetBool(ds, "Show", true);
    sets += TrySetBool(ds, "Enabled", true);

    // ByLayer
    try { ds.Color = byLayerColor; sets++; } catch { }

    // Linetype may be string or ObjectId depending on build
    try
    {
        var pi = ds.GetType().GetProperty("Linetype", BindingFlags.Public | BindingFlags.Instance);
        if (pi != null && pi.CanWrite)
        {
            if (pi.PropertyType == typeof(string)) { pi.SetValue(ds, "ByLayer"); sets++; }
            else if (pi.PropertyType == typeof(ObjectId)) { pi.SetValue(ds, byLayerLinetypeId); sets++; }
        }
    }
    catch { }

    sets += TrySetLineweight(ds, "Lineweight");
    sets += TrySetLineweight(ds, "LineWeight");

    return sets;
}

private static int TrySetString(object obj, string propName, string value)
{
    try
    {
        var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
        if (pi == null || !pi.CanWrite) return 0;
        if (pi.PropertyType != typeof(string)) return 0;
        pi.SetValue(obj, value);
        return 1;
    }
    catch { return 0; }
}

private static int TrySetBool(object obj, string propName, bool value)
{
    try
    {
        var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
        if (pi == null || !pi.CanWrite) return 0;

        if (pi.PropertyType == typeof(bool))
        {
            pi.SetValue(obj, value);
            return 1;
        }

        if (pi.PropertyType == typeof(int))
        {
            pi.SetValue(obj, value ? 1 : 0);
            return 1;
        }
        if (pi.PropertyType == typeof(short))
        {
            pi.SetValue(obj, (short)(value ? 1 : 0));
            return 1;
        }

        if (pi.PropertyType.IsEnum)
        {
            string[] tries = value ? new[] { "True", "On", "Yes", "Enabled" } : new[] { "False", "Off", "No", "Disabled" };
            foreach (var s in tries)
            {
                try
                {
                    var ev = Enum.Parse(pi.PropertyType, s, true);
                    pi.SetValue(obj, ev);
                    return 1;
                }
                catch { }
            }
        }

        return 0;
    }
    catch
    {
        return 0;
    }
}

private static int TrySetProperty(object obj, string propName, object value)
{
    try
    {
        var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
        if (pi == null || !pi.CanWrite) return 0;

        // Ensure assignable
        if (value != null && !pi.PropertyType.IsAssignableFrom(value.GetType()))
        {
            // attempt string conversion for string props
            if (pi.PropertyType == typeof(string))
            {
                pi.SetValue(obj, Convert.ToString(value, CultureInfo.InvariantCulture));
                return 1;
            }
            return 0;
        }

        pi.SetValue(obj, value);
        return 1;
    }
    catch
    {
        return 0;
    }
}

private static int TrySetLinetype(object obj, string propName, ObjectId byLayerLinetypeId)
{
    try
    {
        var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
        if (pi == null || !pi.CanWrite) return 0;

        if (pi.PropertyType == typeof(ObjectId))
        {
            pi.SetValue(obj, byLayerLinetypeId);
            return 1;
        }

        if (pi.PropertyType == typeof(string))
        {
            pi.SetValue(obj, "ByLayer");
            return 1;
        }

        if (pi.PropertyType.IsEnum)
        {
            // Try parse "ByLayer"
            object v = Enum.Parse(pi.PropertyType, "ByLayer", true);
            pi.SetValue(obj, v);
            return 1;
        }
    }
    catch
    {
        // ignore
    }
    return 0;
}

private static int TrySetLineweight(object obj, string propName)
{
    try
    {
        var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
        if (pi == null || !pi.CanWrite) return 0;

        if (pi.PropertyType == typeof(LineWeight))
        {
            pi.SetValue(obj, LineWeight.ByLayer);
            return 1;
        }

        if (pi.PropertyType.IsEnum)
        {
            // many APIs use an enum with ByLayer
            object v = Enum.Parse(pi.PropertyType, "ByLayer", true);
            pi.SetValue(obj, v);
            return 1;
        }
    }
    catch
    {
        // ignore
    }
    return 0;
}

private static ObjectId GetLinetypeId(Database db, Transaction tr, string linetypeName)
{
    try
    {
        var ltt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);
        if (ltt.Has(linetypeName))
            return ltt[linetypeName];
    }
    catch
    {
        // ignore
    }
    return ObjectId.Null;
}

        // ------------------------
        // Utilities
        // ------------------------

        private static string SafeGet(object obj, string prop)
        {
            try
            {
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
            catch { }
        }

        /// <summary>
        /// Reference-equality comparer for graph traversal.
        /// </summary>
        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceEqualityComparer Instance = new ReferenceEqualityComparer();
            public new bool Equals(object x, object y) => ReferenceEquals(x, y);
            public int GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
        }
    }
}
