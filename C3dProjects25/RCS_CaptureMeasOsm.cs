using System;
using System.Collections;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Exception = System.Exception;

namespace RCS.C3D2025.Tools
{
    public class CaptureMeasOsmCommand
    {
        [CommandMethod("RCS_CAPTURE_MEAS_OSM")]
        public void CaptureMeasOsm()
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            if (mdiActiveDocument == null)
            {
                return;
            }
            Database database = mdiActiveDocument.Database;
            Editor editor = mdiActiveDocument.Editor;
            string currentLayout = LayoutManager.Current.CurrentLayout;
            string bestLayoutName = null;
            try
            {
                using (DocumentLock val = mdiActiveDocument.LockDocument())
                {
                    ObjectId layerId = EnsureLayer(database, "image-0");
                    SetCurrentLayer(database, layerId);
                    Extents3d layerExtents = GetLayerExtents(database, "measpl", "pl");
                    Point3d minPoint = layerExtents.MinPoint;
                    double num = minPoint.DistanceTo(layerExtents.MaxPoint);
                    Tolerance global = Tolerance.Global;
                    if (num < global.EqualPoint)
                    {
                        editor.WriteMessage("\nNo measurable extents found on layers \"measpl\" or \"pl\".");
                        return;
                    }
                    ObjectId largestViewportAcrossAllLayouts = GetLargestViewportAcrossAllLayouts(database, currentLayout, out bestLayoutName);
                    if (largestViewportAcrossAllLayouts.IsNull)
                    {
                        editor.WriteMessage("\nNo floating viewports found in the '8.5x11' layout.");
                        return;
                    }
                    editor.WriteMessage($"\nAutomatically locked onto floating viewport '{largestViewportAcrossAllLayouts.Handle}' located on layout '{bestLayoutName}'.");
                    LayoutManager.Current.CurrentLayout = "Model";
                    SetWorldPlanView(editor);
                    ZoomToExtents(editor, layerExtents, 1.08);
                    SetCurrentLayer(database, layerId);
                    try
                    {
                        editor.Command("_.GEOMAP", "openstreetMap");
                    }
                    catch
                    {
                        editor.Command("\u0003");
                    }
                    try
                    {
                        editor.Command("_.GEOMAPIMAGE", "_Viewport");
                    }
                    catch
                    {
                        editor.Command("\u0003");
                    }
                    LayoutManager.Current.CurrentLayout = bestLayoutName;
                    ActivateViewportAndFreezeAllBut(database, editor, largestViewportAcrossAllLayouts, "image-0", layerExtents);
                    editor.SwitchToPaperSpace();
                    database.SaveAs(mdiActiveDocument.Name, true, (DwgVersion)33, mdiActiveDocument.Database.SecurityParameters);
                    editor.WriteMessage("\nDone. Captured map to layer \"image-0\", updated viewport, and saved drawing.");
                }
            }
            catch (Exception ex)
            {
                editor.WriteMessage("\nError: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(bestLayoutName))
                    {
                        LayoutManager.Current.CurrentLayout = bestLayoutName;
                        editor.SwitchToPaperSpace();
                    }
                }
                catch
                {
                }
            }
        }

        private static ObjectId EnsureLayer(Database db, string layerName)
        {
            using (Transaction val = db.TransactionManager.StartTransaction())
            {
                LayerTable val2 = (LayerTable)val.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (val2.Has(layerName))
                {
                    ObjectId result = val2[layerName];
                    val.Commit();
                    return result;
                }
                val2.UpgradeOpen();
                LayerTableRecord val3 = new LayerTableRecord
                {
                    Name = layerName
                };
                ObjectId result2 = val2.Add(val3);
                val.AddNewlyCreatedDBObject(val3, true);
                val.Commit();
                return result2;
            }
        }

        private static void SetCurrentLayer(Database db, ObjectId layerId)
        {
            using (Transaction val = db.TransactionManager.StartTransaction())
            {
                db.Clayer = layerId;
                val.Commit();
            }
        }

        private static Extents3d GetLayerExtents(Database db, params string[] layerNames)
        {
            bool flag = false;
            Extents3d result = default(Extents3d);
            using (Transaction val = db.TransactionManager.StartTransaction())
            {
                BlockTable val2 = (BlockTable)val.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (!val2.Has(BlockTableRecord.ModelSpace))
                {
                    throw new InvalidOperationException("ModelSpace not found.");
                }
                BlockTableRecord val3 = (BlockTableRecord)val.GetObject(val2[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                
                foreach (ObjectId current in val3)
                {
                    if (!current.IsValid || current.IsErased)
                    {
                        continue;
                    }
                    DBObject obj = val.GetObject(current, OpenMode.ForRead, false);
                    Entity val4 = obj as Entity;
                    if (val4 == null) continue;
                    
                    bool match = false;
                    foreach(string ln in layerNames) 
                    {
                        if (string.Equals(val4.Layer, ln, StringComparison.OrdinalIgnoreCase)) 
                        {
                            match = true; 
                            break;
                        }
                    }
                    if (!match) continue;
                    try
                    {
                        Extents3d geometricExtents = val4.GeometricExtents;
                        if (!flag)
                        {
                            result = geometricExtents;
                            flag = true;
                        }
                        else
                        {
                            result.AddExtents(geometricExtents);
                        }
                    }
                    catch
                    {
                    }
                }
                
                val.Commit();
                if (!flag)
                {
                    throw new InvalidOperationException("No entities with extents found on layers: " + string.Join(", ", layerNames) + ".");
                }
                return result;
            }
        }

        private static ObjectId GetLargestViewportAcrossAllLayouts(Database db, string originalLayout, out string bestLayoutName)
        {
            bestLayoutName = null;
            ObjectId result = ObjectId.Null;
            double num = -1.0;
            using (Transaction val = db.TransactionManager.StartTransaction())
            {
                DBDictionary val2 = (DBDictionary)val.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                bool flag = !string.Equals(originalLayout, "Model", StringComparison.OrdinalIgnoreCase) && val2.Contains(originalLayout);
                
                foreach (DBDictionaryEntry current in val2)
                {
                    if (string.Equals(current.Key, "Model", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Enforce using only the 8.5x11 layout
                    if (!string.Equals(current.Key, "8.5x11", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    Layout val3 = (Layout)val.GetObject(current.Value, OpenMode.ForRead);
                    foreach (ObjectId viewport in val3.GetViewports())
                    {
                        ObjectId val4 = viewport;
                        if (val4.IsNull || !val4.IsValid)
                        {
                            continue;
                        }
                        DBObject obj = val.GetObject(val4, OpenMode.ForRead);
                        Viewport val5 = obj as Viewport;
                        if (val5 != null && val5.Number > 1 && val5.Width > 0.0 && val5.Height > 0.0)
                        {
                            double num2 = val5.Width * val5.Height;
                            double num3 = (flag && string.Equals(current.Key, originalLayout, StringComparison.OrdinalIgnoreCase)) ? (num2 * 1000.0) : num2;
                            if (val5.Handle.ToString().Equals("4F929", StringComparison.OrdinalIgnoreCase))
                            {
                                num3 = double.MaxValue;
                            }
                            if (num3 > num)
                            {
                                num = num3;
                                result = val4;
                                bestLayoutName = current.Key;
                            }
                        }
                    }
                }
                val.Commit();
                return result;
            }
        }

        private static void SetWorldPlanView(Editor ed)
        {
            ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
            using (ViewTableRecord currentView = ed.GetCurrentView())
            {
                currentView.ViewDirection = Vector3d.ZAxis;
                currentView.ViewTwist = 0.0;
                ed.SetCurrentView(currentView);
            }
        }

        private static void ZoomToExtents(Editor ed, Extents3d ext, double padFactor)
        {
            try
            {
                Point3d minPoint = ext.MinPoint;
                Point3d maxPoint = ext.MaxPoint;
                double num = Math.Max(maxPoint.X - minPoint.X, 1.0);
                double num2 = Math.Max(maxPoint.Y - minPoint.Y, 1.0);
                Point2d centerPoint = new Point2d((minPoint.X + maxPoint.X) * 0.5, (minPoint.Y + maxPoint.Y) * 0.5);
                using (ViewTableRecord currentView = ed.GetCurrentView())
                {
                    currentView.CenterPoint = centerPoint;
                    currentView.Width = num * padFactor;
                    currentView.Height = num2 * padFactor;
                    ed.SetCurrentView(currentView);
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage("\n[ZoomToExtents Error] " + ex.Message);
            }
        }

        private static void ActivateViewportAndFreezeAllBut(Database db, Editor ed, ObjectId viewportId, string keepLayerName, Extents3d measExtents)
        {
            using (Transaction val = db.TransactionManager.StartTransaction())
            {
                Viewport val2 = (Viewport)val.GetObject(viewportId, OpenMode.ForWrite);
                LayerTable val3 = (LayerTable)val.GetObject(db.LayerTableId, OpenMode.ForRead);
                ObjectIdCollection val4 = new ObjectIdCollection();
                foreach (ObjectId current in val3)
                {
                    LayerTableRecord val5 = (LayerTableRecord)val.GetObject(current, OpenMode.ForRead);
                    if (!string.Equals(val5.Name, keepLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        val4.Add(current);
                    }
                }
                val.Commit();
                
                ed.SwitchToModelSpace();
                using (Transaction val6 = db.TransactionManager.StartTransaction())
                {
                    Viewport val7 = (Viewport)val6.GetObject(viewportId, OpenMode.ForWrite);
                    Application.SetSystemVariable("CVPORT", val7.Number);
                    val7.Locked = false;
                    val7.FreezeLayersInViewport(val4.GetEnumerator());
                    val6.Commit();
                    SetWorldPlanView(ed);
                    ZoomToExtents(ed, measExtents, 1.05);
                }
            }
        }
    }
}
