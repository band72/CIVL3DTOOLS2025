using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using RCS.CustomLeader.Abstractions;

namespace RCS.CustomLeader.Core.Geometry
{
    public static class ArcLeaderGeometryService
    {
        public static Arc CreateArc(Point3d p1, Point3d p2, Point3d p3)
        {
            try
            {
                var arc3d = new CircularArc3d(p1, p2, p3);
                var plane = new Plane(arc3d.Center, arc3d.Normal);
                // Convert 3D mathematical circle/arc logic into a Database Arc
                return new Arc(arc3d.Center, arc3d.Normal, arc3d.Radius, arc3d.ReferenceVector.AngleOnPlane(plane), arc3d.ReferenceVector.AngleOnPlane(plane) + arc3d.EndAngle);
            }
            catch
            {
                // Fallback geometry so we never crash AutoCAD if points are identical/colinear
                return new Arc(p1, 5.0, 0, Math.PI / 2);
            }
        }

        public static double GetStartTangentAngle(Arc arc)
        {
            var plane = new Plane(arc.Center, arc.Normal);
            var firstDeriv = arc.GetFirstDerivative(arc.StartParam);
            return firstDeriv.AngleOnPlane(plane);
        }

        public static Entity CreateHeadBlock(Point3d position, double tangentAngle, ArcLeaderSettings settings)
        {
            // Dynamically scale the arrowhead based on Text Height
            double scale = settings.TextHeight;
            double length = scale * 1.5; 
            double halfWidth = scale * 0.25;

            // tangentAngle points *along* the curve leaving the tip
            Vector3d dir = new Vector3d(Math.Cos(tangentAngle), Math.Sin(tangentAngle), 0);
            Vector3d perp = new Vector3d(-Math.Sin(tangentAngle), Math.Cos(tangentAngle), 0); 

            Point3d p1 = position; // Tip of the arrow
            Point3d p2 = position + dir * length + perp * halfWidth; 
            Point3d p3 = position + dir * length - perp * halfWidth; 

            // Returns a 2D filled Solid triangle for the arrowhead
            return new Solid(p1, p2, p3);
        }
    }
}
