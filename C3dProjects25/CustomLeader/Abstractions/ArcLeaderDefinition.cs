using Autodesk.AutoCAD.DatabaseServices; // Needed for ObjectId
using Autodesk.AutoCAD.Geometry;         // Needed for Point3d

namespace RCS.CustomLeader.Abstractions
{
    public sealed class ArcLeaderDefinition
    {
        // Source geometry
        public Point3d HeadPoint { get; set; }
        public Point3d ThroughPoint { get; set; }
        public Point3d BoxPoint { get; set; }

        // Component Handles / IDs
        public ObjectId ArcId { get; set; }
        public ObjectId TailArcId { get; set; }   // V2 only: secondary tail arc
        public ObjectId HeadBlockId { get; set; }
        public ObjectId TextId { get; set; }
        public ObjectId FrameId { get; set; }
        public ObjectId GroupId { get; set; }

        // Metadata
        public string TextValue { get; set; } = string.Empty;
        public string StyleName { get; set; } = "Default";
        public int SchemaVersion { get; set; } = 1;
    }
}
