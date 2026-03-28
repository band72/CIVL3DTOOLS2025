using System;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using RCS.CustomLeader.Abstractions;
using RCS.CustomLeader.AutoCAD.Jigs;
using RCS.CustomLeader.Core.Builders;
using RCS.CustomLeader.Core.Persistence;

namespace RCS.CustomLeader.AutoCAD.Commands
{
    /// <summary>
    /// RCS_ARCLEADER_V2
    /// Same 3-point interaction as V1 (head → through → box) but the final annotation
    /// contains only:  arrowhead tip  +  gently-curved tail arc  +  text box.
    /// The primary arc is used in-memory to derive geometry, then discarded.
    /// </summary>
    public class CreateArcLeaderV2Command
    {
        [CommandMethod("RCS_ARCLEADER_V2")]
        public void CreateV2Command()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed  = doc.Editor;
            var db  = doc.Database;

            try
            {
                // 1. Head point (arrow tip)
                var p1Res = ed.GetPoint("\nSpecify leader head point: ");
                if (p1Res.Status != PromptStatus.OK) return;

                // 2. Through point (defines main arc curvature / arrowhead base)
                var p2Res = ed.GetPoint(new PromptPointOptions("\nSpecify arc through point: ")
                {
                    BasePoint    = p1Res.Value,
                    UseBasePoint = true
                });
                if (p2Res.Status != PromptStatus.OK) return;

                // 3. Drag jig for box point — previews tail arc + arrowhead + text box
                var settings = ArcLeaderSettings.Current;
                var jig      = new ArcLeaderV2PlacementJig(p1Res.Value, p2Res.Value, settings);

                var p3Res = ed.Drag(jig);
                if (p3Res.Status != PromptStatus.OK) return;

                // 4. Text content
                var strRes = ed.GetString(new PromptStringOptions("\nEnter text: "));
                if (strRes.Status != PromptStatus.OK) return;

                // 5. Build V2 and persist
                var builder = new ArcLeaderBuilder(db);
                var repo    = new XDataArcLeaderRepository();

                var definition = builder.BuildV2(
                    p1Res.Value,
                    p2Res.Value,
                    jig.BoxPoint,
                    strRes.StringResult,
                    settings);

                repo.Save(definition);

                ed.WriteMessage("\nArc Leader V2 created successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating Arc Leader V2: {ex.Message}");
            }
        }
    }
}
