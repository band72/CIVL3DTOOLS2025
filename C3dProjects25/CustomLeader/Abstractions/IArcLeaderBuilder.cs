using Autodesk.AutoCAD.Geometry;

namespace RCS.CustomLeader.Abstractions
{
    public interface IArcLeaderBuilder
    {
        ArcLeaderDefinition Build(Point3d p1, Point3d p2, Point3d p3, string text, ArcLeaderSettings settings);
        void Rebuild(ArcLeaderDefinition definition, ArcLeaderSettings settings);
    }
}
