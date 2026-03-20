using System;
using Autodesk.AutoCAD.DatabaseServices;
using RCS.CustomLeader.Abstractions;

namespace RCS.CustomLeader.Core.Persistence
{
    public class XDataArcLeaderRepository : IArcLeaderRepository
    {
        private const string AppName = "RCS_ARC_LEADER";

        public void Save(ArcLeaderDefinition definition)
        {
            try
            {
                // TODO: Serialize `definition` points and IDs into an XData ResultBuffer
                // Attach the ResultBuffer to the Group or the main Arc entity
            }
            catch (Exception ex)
            {
                // TODO: log robustly or throw domain specific exception
                throw new Exception($"Failed to save ArcLeader XData: {ex.Message}", ex);
            }
        }

        public ArcLeaderDefinition Load(ObjectId anyComponentId)
        {
            try
            {
                // TODO: Traverse from the selected component up to the Group
                // Extract the XData and populate an ArcLeaderDefinition object
                return new ArcLeaderDefinition(); 
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to load ArcLeader definition: {ex.Message}", ex);
            }
        }

        public bool IsCustomLeader(ObjectId entityId)
        {
            try
            {
                // TODO: Check if entity has the "RCS_ARC_LEADER" XData footprint attached
                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}
