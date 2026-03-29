using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.Civil.ApplicationServices;
using System;

namespace RCS.C3D2025.Tools
{
    public class CoordinateSystemCommands
    {
        [CommandMethod("RCS_SET_FL83EF")]
        public void SetFlCoordinateSystem()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                var civDoc = CivilApplication.ActiveDocument;
                if (civDoc != null)
                {
                    civDoc.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode = "FL83-EF";
                    doc.Editor.WriteMessage("\n[RCS] Coordinate System strictly bound to FL83-EF (NAD83 Florida State Planes, East Zone, US Foot).");
                    
                    // Hook aerial map via command line execution stream
                    doc.SendStringToExecute("_GEOMAP\n_A\n", true, false, true);
                }
                else
                {
                    doc.Editor.WriteMessage("\n[RCS ERROR] CivilApplication.ActiveDocument is null.");
                }
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n[RCS ERROR] Failed to assign mapping settings: {ex.Message}");
            }
        }
    }
}
