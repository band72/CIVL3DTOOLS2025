using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(RcsTools.CaptureMeasOsmCommand))]

namespace RcsTools
{
    public class CaptureMeasOsmCommand : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("RCS_CAPTURE_MEAS_OSM")]
        public void CaptureMeasOsm()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Database db = doc.Database;
            Editor ed = doc.Editor;

            const string imageLayerName = "image-0";
            const string measLayerName = "measpl";

            string originalLayout = LayoutManager.Current.CurrentLayout;
            string targetPaperLayout = null;

            try
            {
                using (doc.LockDocument())
                {
                    ObjectId imageLayerId = EnsureLayer(db, imageLayerName);
                    SetCurrentLayer(db, imageLayerId);

                    Extents3d measExtents = GetLayerExtents(db, measLayerName);
                    if (measExtents.MinPoint.DistanceTo(measExtents.MaxPoint) < Tolerance.Global.EqualPoint)
                    {
                        ed.WriteMessage($"\nNo measurable extents found on layer \"{measLayerName}\".");
                        return;
                    }

                    ObjectId viewportId = GetLargestViewportAcrossAllLayouts(db, originalLayout, out targetPaperLayout);
                    if (viewportId.IsNull)
                    {
                        ed.WriteMessage("\nNo floating viewports found in any paper space layout.");
                        return;
                    }
                    else
                    {
                        ed.WriteMessage($"\nAutomatically locked onto floating viewport '{viewportId.Handle}' located on layout '{targetPaperLayout}'.");
                    }

                    // Go to model space.
                    LayoutManager.Current.CurrentLayout = "Model";

                    // Force a WCS plan view and zoom to MEAS extents before map capture.
                    SetWorldPlanView(ed);
                    ZoomToExtents(ed, measExtents, 1.08);

                    // Make sure image-0 is current before capture.
                    SetCurrentLayer(db, imageLayerId);

                    try { ed.Command("_.GEOMAP", "openstreetMap"); } catch { ed.Command((object)"\x03"); } // \x03 is Esc to clear prompt if it failed
                    try { ed.Command("_.GEOMAPIMAGE", "_Viewport"); } catch { ed.Command((object)"\x03"); }

                    // Switch to paper space layout.
                    LayoutManager.Current.CurrentLayout = targetPaperLayout;

                    // Activate viewport to freeze layers, explicitly unlock it, and frame/center the object.
                    ActivateViewportAndFreezeAllBut(db, ed, viewportId, imageLayerName, measExtents);

                    // Return to paper space and save.
                    ed.SwitchToPaperSpace();
                    db.SaveAs(doc.Name, true, DwgVersion.Current, doc.Database.SecurityParameters);

                    ed.WriteMessage($"\nDone. Captured map to layer \"{imageLayerName}\", updated viewport, and saved drawing.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
            finally
            {
                // Best-effort return to paperspace if a paper layout exists.
                try
                {
                    if (!string.IsNullOrWhiteSpace(targetPaperLayout))
                    {
                        LayoutManager.Current.CurrentLayout = targetPaperLayout;
                        ed.SwitchToPaperSpace();
                    }
                }
                catch
                {
                    // swallow cleanup errors
                }
            }
        }

        private static ObjectId EnsureLayer(Database db, string layerName)
        {
            using Transaction tr = db.TransactionManager.StartTransaction();

            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName))
            {
                ObjectId existingId = lt[layerName];
                tr.Commit();
                return existingId;
            }

            lt.UpgradeOpen();

            LayerTableRecord ltr = new LayerTableRecord
            {
                Name = layerName
            };

            ObjectId layerId = lt.Add(ltr);
            tr.AddNewlyCreatedDBObject(ltr, true);
            tr.Commit();
            return layerId;
        }

        private static void SetCurrentLayer(Database db, ObjectId layerId)
        {
            using Transaction tr = db.TransactionManager.StartTransaction();
            db.Clayer = layerId;
            tr.Commit();
        }

        private static Extents3d GetLayerExtents(Database db, string layerName)
        {
            bool found = false;
            Extents3d total = default;

            using Transaction tr = db.TransactionManager.StartTransaction();
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            if (!bt.Has(BlockTableRecord.ModelSpace))
            {
                throw new InvalidOperationException("ModelSpace not found.");
            }

            BlockTableRecord ms =
                (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in ms)
            {
                if (!id.IsValid || id.IsErased) continue;

                if (!(tr.GetObject(id, OpenMode.ForRead, false) is Entity ent)) continue;
                if (!string.Equals(ent.Layer, layerName, StringComparison.OrdinalIgnoreCase)) continue;

                try
                {
                    Extents3d ext = ent.GeometricExtents;

                    if (!found)
                    {
                        total = ext;
                        found = true;
                    }
                    else
                    {
                        total.AddExtents(ext);
                    }
                }
                catch
                {
                    // ignore entities without geometric extents
                }
            }

            tr.Commit();

            if (!found)
            {
                throw new InvalidOperationException($"No entities with extents found on layer \"{layerName}\".");
            }

            return total;
        }

        private static ObjectId GetLargestViewportAcrossAllLayouts(Database db, string originalLayout, out string bestLayoutName)
        {
            bestLayoutName = null;
            ObjectId bestVpId = ObjectId.Null;
            double bestArea = -1.0;

            using Transaction tr = db.TransactionManager.StartTransaction();
            DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            // Give strict priority to original layout if they were already in paperspace
            bool preferOriginal = !string.Equals(originalLayout, "Model", StringComparison.OrdinalIgnoreCase) && layoutDict.Contains(originalLayout);

            foreach (DBDictionaryEntry entry in layoutDict)
            {
                if (string.Equals(entry.Key, "Model", StringComparison.OrdinalIgnoreCase)) continue;

                Layout layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                foreach (ObjectId vpId in layout.GetViewports())
                {
                    if (vpId.IsNull || !vpId.IsValid) continue;
                    if (!(tr.GetObject(vpId, OpenMode.ForRead) is Viewport vp)) continue;
                    if (vp.Number <= 1) continue; // PaperSpace container
                    if (vp.Width <= 0 || vp.Height <= 0) continue;

                    double area = vp.Width * vp.Height;
                    
                    // Massive multiplier if this is the layout the user originated in
                    double score = (preferOriginal && string.Equals(entry.Key, originalLayout, StringComparison.OrdinalIgnoreCase)) ? (area * 1000.0) : area;

                    // Directly prefer handle '4F929' if requested by user for diagnostics tracking
                    if (vp.Handle.ToString().Equals("4F929", StringComparison.OrdinalIgnoreCase)) score = double.MaxValue;

                    if (score > bestArea)
                    {
                        bestArea = score;
                        bestVpId = vpId;
                        bestLayoutName = entry.Key;
                    }
                }
            }

            tr.Commit();
            return bestVpId;
        }

        private static void SetWorldPlanView(Editor ed)
        {
            // Force WCS PLAN native
            ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
            using ViewTableRecord view = ed.GetCurrentView();
            view.ViewDirection = Vector3d.ZAxis;
            view.ViewTwist = 0.0;
            ed.SetCurrentView(view);
        }

        private static void ZoomToExtents(Editor ed, Extents3d ext, double padFactor)
        {
            try
            {
                Point3d min = ext.MinPoint;
                Point3d max = ext.MaxPoint;

                double width = Math.Max(max.X - min.X, 1.0);
                double height = Math.Max(max.Y - min.Y, 1.0);

                Point2d center = new Point2d((min.X + max.X) * 0.5, (min.Y + max.Y) * 0.5);

                using ViewTableRecord view = ed.GetCurrentView();
                view.CenterPoint = center;
                view.Width = width * padFactor;
                view.Height = height * padFactor;
                ed.SetCurrentView(view);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[ZoomToExtents Error] {ex.Message}");
            }
        }

        private static void ActivateViewportAndFreezeAllBut(
            Database db,
            Editor ed,
            ObjectId viewportId,
            string keepLayerName,
            Extents3d measExtents)
        {
            using Transaction tr = db.TransactionManager.StartTransaction();

            Viewport vp = (Viewport)tr.GetObject(viewportId, OpenMode.ForWrite);
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            // Build list of layers to freeze in this viewport.
            ObjectIdCollection freezeIds = new ObjectIdCollection();

            foreach (ObjectId layerId in lt)
            {
                LayerTableRecord ltr =
                    (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);

                if (string.Equals(ltr.Name, keepLayerName, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Don't try to freeze current paper viewport boundary behavior-related oddities here;
                // viewport freeze only affects the model seen through the viewport.
                freezeIds.Add(layerId);
            }

            tr.Commit();

            // Must be in active viewport/modelspace for reliable viewport freezing behavior.
            ed.SwitchToModelSpace();

            using Transaction tr2 = db.TransactionManager.StartTransaction();
            Viewport vp2 = (Viewport)tr2.GetObject(viewportId, OpenMode.ForWrite);

            // Make the chosen viewport current using system variable
            Application.SetSystemVariable("CVPORT", vp2.Number);

            // 1. Explicitly Unlock the viewport before we can alter its zoom/center view dynamically.
            vp2.Locked = false;

            vp2.FreezeLayersInViewport(freezeIds.GetEnumerator());

            tr2.Commit();

            // 2. Perform native view centering over the measpl object!
            SetWorldPlanView(ed);
            ZoomToExtents(ed, measExtents, 1.05); // Use tight padding
        }
    }
}