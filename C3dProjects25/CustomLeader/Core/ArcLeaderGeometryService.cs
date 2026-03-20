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
            return 0.0;
        }

        public static BlockReference CreateHeadBlock(Point3d position, double tangentAngle, ArcLeaderSettings settings)
        {
            // Do not simply return new BlockReference()! It lacks an ObjectId definition and will crash AutoCAD.
            return null;
        }
    }
}
