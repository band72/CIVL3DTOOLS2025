using System;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using RCS.CustomLeader.Abstractions;

namespace RCS.CustomLeader.AutoCAD.Commands
{
    public class SetArcLeaderTextSizeCommand
    {
        [CommandMethod("RCS_ARCLEADER_TEXTSIZE")]
        public void SetTextSizeCommand()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            var opts = new PromptDoubleOptions($"\nEnter new Arc Leader Text Height <{ArcLeaderSettings.Current.TextHeight}>: ")
            {
                AllowNegative = false,
                AllowZero = false,
                DefaultValue = ArcLeaderSettings.Current.TextHeight,
                UseDefaultValue = true
            };

            var res = ed.GetDouble(opts);
            if (res.Status == PromptStatus.OK)
            {
                ArcLeaderSettings.Current.TextHeight = res.Value;
                
                // Automatically scale the boxing and arrowhead proportions to match the new text size
                ArcLeaderSettings.Current.BoxOffset = res.Value * 2.0;
                ArcLeaderSettings.Current.BoxPadding = res.Value * 0.8;
                
                ed.WriteMessage($"\nArc Leader Text Size updated to {res.Value}. Arrowhead dimensions and box padding will automatically scale to match.");
            }
        }
    }
}
