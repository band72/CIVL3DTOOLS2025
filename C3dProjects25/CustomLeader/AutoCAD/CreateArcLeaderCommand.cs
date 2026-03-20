using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using RCS.CustomLeader.Abstractions;
using RCS.CustomLeader.AutoCAD.Jigs;
using RCS.CustomLeader.Core.Builders;
using RCS.CustomLeader.Core.Persistence;

namespace RCS.CustomLeader.AutoCAD.Commands
{
    public class CreateArcLeaderCommand
    {
        [CommandMethod("RCS_ARCLEADER")]
        public void CreateCommand()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // 1. Get Head Point
                var p1Res = ed.GetPoint("\nSpecify leader head point: ");
                if (p1Res.Status != PromptStatus.OK) return;

                // 2. Get Through Point
                var p2Res = ed.GetPoint(new PromptPointOptions("\nSpecify arc through point: ") { BasePoint = p1Res.Value, UseBasePoint = true });
                if (p2Res.Status != PromptStatus.OK) return;

                // 3. User Jig for Box Point
                var settings = new ArcLeaderSettings(); // Later loaded from JSON/Dictionary
                var jig = new ArcLeaderPlacementJig(p1Res.Value, p2Res.Value, settings);
                
                var p3Res = ed.Drag(jig);
                if (p3Res.Status != PromptStatus.OK) return;

                // 4. Ask for Text (Default to INTEX)
                var strRes = ed.GetString(new PromptStringOptions("\nEnter text: ") { DefaultValue = settings.DefaultText });
                if (strRes.Status != PromptStatus.OK) return;

                // 5. Build and Save
                var builder = new ArcLeaderBuilder(db);
                var repo = new XDataArcLeaderRepository();

                var definition = builder.Build(p1Res.Value, p2Res.Value, jig.BoxPoint, strRes.StringResult, settings);
                repo.Save(definition);

                ed.WriteMessage("\nArc Leader created successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating Arc Leader: {ex.Message}");
                // TODO: robust logging mechanism (e.g. NLog, Serilog)
            }
        }
    }
}
