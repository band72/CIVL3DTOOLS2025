using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025.Tools
{
    public class HelpCommand
    {
        [CommandMethod("RCS_HELP")]
        public void ShowHelp()
        {
            try
            {
                var win = new HelpWindow();
                Application.ShowModalWindow(win);
            }
            catch (System.Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage($"\nFailed to launch Help window: {ex.Message}");
                }
            }
        }
    }
}
