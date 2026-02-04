using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025.Tools
{
    public class Commands : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    "\nRCS_C3D2025_Tools loaded. Commands: RCS_TEST, RCS_HELP_RCS, RCS_FIX_DESC_KEY_SCALE");
            }
            catch { }
        }

        public void Terminate() { }

        [CommandMethod("RCS_TEST")]
        public void RCS_TEST()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Editor.WriteMessage("\nRCS_C3D2025_Tools is working.");
        }

        [CommandMethod("RCS_HELP_RCS")]
        public void RCS_HELP_RCS()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            ed.WriteMessage(
                "\nRCS Commands:" +
                "\n - RCS_TEST" +
                "\n - RCS_FIX_DESC_KEY_SCALE (FixedScaleFactor=0.02, UseDrawingScale=true)");
        }
    }
}
