using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using RCS.CustomLeader.Abstractions;

namespace RCS.CustomLeader.Core.Geometry
{
    public static class ArcLeaderBoxService
    {
        public static MText CreateBox(Point3d position, string text, ArcLeaderSettings settings)
        {
            if (string.IsNullOrEmpty(text))
                return null;

            var mtext = new MText();
            mtext.Location = position;
            mtext.Contents = text;
            mtext.TextHeight = settings.TextHeight;
            
            // Critical for accurate bounding box trimming later
            mtext.Attachment = AttachmentPoint.MiddleCenter;

            if (settings.UseBackgroundMask)
            {
                mtext.BackgroundFill = true;
                mtext.UseBackgroundColor = true;
                mtext.BackgroundScaleFactor = 1.25;
            }

            return mtext;
        }

        public static void TrimArcToMText(Arc arc, MText mtext, ArcLeaderSettings settings)
        {
            try
            {
                // ActualWidth is computed natively by AutoCAD even before plotting
                double hw = (mtext.ActualWidth / 2.0) + settings.BoxPadding;
                double hh = (mtext.ActualHeight / 2.0) + settings.BoxPadding;

                Point3d c = mtext.Location;

                using (var poly = new Polyline(4))
                {
                    poly.AddVertexAt(0, new Point2d(c.X - hw, c.Y - hh), 0, 0, 0);
                    poly.AddVertexAt(1, new Point2d(c.X + hw, c.Y - hh), 0, 0, 0);
                    poly.AddVertexAt(2, new Point2d(c.X + hw, c.Y + hh), 0, 0, 0);
                    poly.AddVertexAt(3, new Point2d(c.X - hw, c.Y + hh), 0, 0, 0);
                    poly.Closed = true;

                    var pts = new Point3dCollection();
                    arc.IntersectWith(poly, Intersect.OnBothOperands, pts, IntPtr.Zero, IntPtr.Zero);

                    if (pts.Count > 0)
                    {
                        // Get the intersection closest to the MText center
                        Point3d closest = pts[0];
                        double minDist = closest.DistanceTo(c);
                        for (int i = 1; i < pts.Count; i++)
                        {
                            double d = pts[i].DistanceTo(c);
                            if (d < minDist)
                            {
                                minDist = d;
                                closest = pts[i];
                            }
                        }

                        Vector3d v = closest - arc.Center;
                        double trimAngle = v.AngleOnPlane(new Plane(arc.Center, arc.Normal));

                        if (arc.StartPoint.DistanceTo(c) < arc.EndPoint.DistanceTo(c))
                            arc.StartAngle = NormalizeAngleTo(arc.StartAngle, trimAngle);
                        else
                            arc.EndAngle = NormalizeAngleTo(arc.EndAngle, trimAngle);
                    }
                }
            }
            catch { }
        }

        private static double NormalizeAngleTo(double original, double target)
        {
            while (target < original - Math.PI) target += 2 * Math.PI;
            while (target > original + Math.PI) target -= 2 * Math.PI;
            return target;
        }
    }
}
