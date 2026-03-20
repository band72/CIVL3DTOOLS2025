using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using RCS.CustomLeader.Abstractions;
using RCS.CustomLeader.Core.Geometry;

namespace RCS.CustomLeader.AutoCAD.Jigs
{
    public class ArcLeaderPlacementJig : DrawJig
    {
        private Point3d _p1;
        private Point3d _p2;
        private Point3d _p3; // The current mouse location
        private ArcLeaderSettings _settings;

        public Point3d BoxPoint => _p3;

        public ArcLeaderPlacementJig(Point3d p1, Point3d p2, ArcLeaderSettings settings)
        {
            _p1 = p1;
            _p2 = p2;
            _settings = settings;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            try
            {
                var options = new JigPromptPointOptions("\nSpecify text box location: ");
                options.BasePoint = _p2;
                options.UseBasePoint = true;

                var result = prompts.AcquirePoint(options);
                
                if (result.Value.IsEqualTo(_p3))
                    return SamplerStatus.NoChange;

                _p3 = result.Value;
                return SamplerStatus.OK;
            }
            catch
            {
                return SamplerStatus.Cancel;
            }
        }

        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            try
            {
                // Draw Provisional Arc
                using (var arc = ArcLeaderGeometryService.CreateArc(_p1, _p2, _p3))
                {
                    if (arc != null) 
                    {
                        draw.Geometry.Draw(arc);
                        
                        // Because actual block inserts require Database access which isn't safe inside a Jig WorldDraw
                        // We will instead draw a temporary cross or vector line at p1 showing the tangent angle.
                        var tangentAngle = ArcLeaderGeometryService.GetStartTangentAngle(arc);
                        // A short graphical 1-unit line pointing tangentially where the arrow block will point
                        var dummyArrowPath = new Autodesk.AutoCAD.Geometry.Vector3d(Math.Cos(tangentAngle), Math.Sin(tangentAngle), 0).GetNormal() * (_settings.TextHeight * 5);
                        draw.Geometry.WorldLine(_p1, _p1 + dummyArrowPath);
                    }
                }

                // Draw Provisional Box using _p3
                using (var mtext = ArcLeaderBoxService.CreateBox(_p3, _settings.DefaultText, _settings))
                {
                    if (mtext != null) draw.Geometry.Draw(mtext);
                }

                return true;
            }
            catch
            {
                // Returning false disables drawing the frame, but doesn't crash the jig
                return false;
            }
        }
    }
}
