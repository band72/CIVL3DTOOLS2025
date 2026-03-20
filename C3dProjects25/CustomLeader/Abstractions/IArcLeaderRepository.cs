using Autodesk.AutoCAD.DatabaseServices;

namespace RCS.CustomLeader.Abstractions
{
    public interface IArcLeaderRepository
    {
        void Save(ArcLeaderDefinition definition);
        ArcLeaderDefinition Load(ObjectId anyComponentId);
        bool IsCustomLeader(ObjectId entityId);
    }
}
