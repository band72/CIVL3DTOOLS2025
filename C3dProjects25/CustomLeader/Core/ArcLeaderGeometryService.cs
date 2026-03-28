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
                return new Arc(arc3d.Center, arc3d.Normal, arc3d.Radius,
                               arc3d.ReferenceVector.AngleOnPlane(plane),
                               arc3d.ReferenceVector.AngleOnPlane(plane) + arc3d.EndAngle);
            }
            catch
            {
                return new Arc(p1, 5.0, 0, Math.PI / 2);
            }
        }

        /// <summary>
        /// Creates a curved arc between <paramref name="from"/> and <paramref name="to"/> using
        /// the sagitta formula to match <paramref name="radius"/> exactly.
        /// When <paramref name="radius"/> is 0 or less than half the chord, falls back to a
        /// gentle 15% chord-offset heuristic.
        /// </summary>
        public static Arc CreateTailArc(Point3d from, Point3d to, double radius = 0)
        {
            try
            {
                Vector3d chord    = to - from;
                double   chordLen = chord.Length;
                if (chordLen < 1e-6) return null;

                // Midpoint of the chord
                Point3d mid = new Point3d(
                    (from.X + to.X) / 2.0,
                    (from.Y + to.Y) / 2.0,
                    (from.Z + to.Z) / 2.0);

                // Unit perpendicular to chord in the XY plane
                Vector3d perp = new Vector3d(-chord.Y, chord.X, 0).GetNormal();

                double halfChord = chordLen / 2.0;
                double offset;

                if (radius > halfChord)
                {
                    // Sagitta: the perpendicular distance from chord midpoint to arc midpoint
                    // that exactly reproduces the requested radius.
                    offset = radius - Math.Sqrt(radius * radius - halfChord * halfChord);
                }
                else
                {
                    // Fallback: radius too small for this chord — use gentle percentage heuristic
                    offset = chordLen * 0.15;
                }

                Point3d throughPoint = mid + perp * offset;
                return CreateArc(from, throughPoint, to);
            }
            catch
            {
                return null;
            }
        }

        public static double GetStartTangentAngle(Arc arc)
        {
            var plane      = new Plane(arc.Center, arc.Normal);
            var firstDeriv = arc.GetFirstDerivative(arc.StartParam);
            return firstDeriv.AngleOnPlane(plane);
        }

        public static Entity CreateHeadBlock(Point3d position, double tangentAngle, ArcLeaderSettings settings)
        {
            double scale     = settings.TextHeight;
            double length    = scale * 1.5;
            double halfWidth = scale * 0.25;

            Vector3d dir  = new Vector3d(Math.Cos(tangentAngle), Math.Sin(tangentAngle), 0);
            Vector3d perp = new Vector3d(-Math.Sin(tangentAngle), Math.Cos(tangentAngle), 0);

            Point3d p1 = position;
            Point3d p2 = position + dir * length + perp * halfWidth;
            Point3d p3 = position + dir * length - perp * halfWidth;

            return new Solid(p1, p2, p3);
        }
    }
}
