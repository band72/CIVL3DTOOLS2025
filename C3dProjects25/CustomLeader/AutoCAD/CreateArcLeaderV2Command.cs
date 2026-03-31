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

            // Save and suppress OSNAP before entering the try block so catch can always restore it
            int savedOsnap = Convert.ToInt32(
                Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("OSMODE"));

            try
            {
                // 1. Head point (arrow tip) — OSNAP OFF so snap cannot bend the arrow direction
                var p1Opts = new PromptPointOptions("\nSpecify leader head point (OSNAP suspended): ")
                {
                    AllowNone = false
                };
                Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", 0);

                var p1Res = ed.GetPoint(p1Opts);
                if (p1Res.Status != PromptStatus.OK)
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", savedOsnap);
                    return;
                }

                // 2. Through point — also snap-free to prevent arc distortion mid-pick
                var p2Res = ed.GetPoint(new PromptPointOptions("\nSpecify arc through point: ")
                {
                    BasePoint    = p1Res.Value,
                    UseBasePoint = true,
                    AllowNone    = false
                });
                if (p2Res.Status != PromptStatus.OK)
                {
                    Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", savedOsnap);
                    return;
                }

                // 3. Drag jig for box point — previews tail arc + arrowhead + text box
                //    Restore snap before the jig so the text box can optionally snap
                Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", savedOsnap);

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
                // Always restore OSNAP even on error
                try { Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("OSMODE", savedOsnap); }
                catch { }
                ed.WriteMessage($"\nError creating Arc Leader V2: {ex.Message}");
            }
        }
    }
}
