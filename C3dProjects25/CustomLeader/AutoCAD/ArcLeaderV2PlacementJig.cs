using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using RCS.CustomLeader.Abstractions;
using RCS.CustomLeader.Core.Geometry;

namespace RCS.CustomLeader.AutoCAD.Jigs
{
    /// <summary>
    /// V2 placement jig. Previews the tail arc (from p2 to the computed main-arc endpoint)
    /// plus the arrowhead indicator and text box. The main arc is NOT drawn — it exists only
    /// to derive geometry during the drag loop.
    /// </summary>
    public class ArcLeaderV2PlacementJig : DrawJig
    {
        private readonly Point3d         _p1;
        private readonly Point3d         _p2;
        private readonly ArcLeaderSettings _settings;
        private Point3d                  _p3; // current mouse location

        public Point3d BoxPoint => _p3;

        public ArcLeaderV2PlacementJig(Point3d p1, Point3d p2, ArcLeaderSettings settings)
        {
            _p1       = p1;
            _p2       = p2;
            _settings = settings;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            try
            {
                var options = new JigPromptPointOptions("\nSpecify text box location: ")
                {
                    BasePoint    = _p2,
                    UseBasePoint = true
                };

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
                // ── 1. Main arc (same as V1) ───────────────────────────────────────────────
                double  tangentAngle = 0.0;
                Point3d firstArcEnd  = _p3;

                using (var mainArc = ArcLeaderGeometryService.CreateArc(_p1, _p2, _p3))
                {
                    if (mainArc != null)
                    {
                        tangentAngle = ArcLeaderGeometryService.GetChordAngleAt(mainArc, _p1, _settings.TextHeight * 1.5);

                        // Trim against text box to find the far endpoint
                        using (var tempMtext = ArcLeaderBoxService.CreateBox(_p3, _settings.DefaultText, _settings))
                        {
                            if (tempMtext != null)
                                ArcLeaderBoxService.TrimArcToMText(mainArc, tempMtext, _settings);
                        }

                        firstArcEnd = mainArc.StartPoint.DistanceTo(_p3) < mainArc.EndPoint.DistanceTo(_p3)
                                      ? mainArc.StartPoint
                                      : mainArc.EndPoint;

                        // ── 2. Arrowhead indicator (same as V1) ───────────────────────────
                        double  arrowLength = _settings.TextHeight * 1.5;
                        var     tangentDir  = new Vector3d(Math.Cos(tangentAngle), Math.Sin(tangentAngle), 0);
                        Point3d arrowBase   = _p1 + tangentDir * arrowLength;
                        draw.Geometry.WorldLine(_p1, arrowBase);

                        // ── 3. Second arc: exactly traces main arc ────────────────────────
                        using (var tailArc = (Arc)mainArc.Clone())
                        {
                            Point3d trueBaseOnArc = tailArc.GetClosestPointTo(arrowBase, false);
                            double paramAtBase = tailArc.GetParameterAtPoint(trueBaseOnArc);

                            double arcLength = tailArc.GetDistanceAtParameter(tailArc.EndParam);
                            if (arcLength > arrowLength)
                            {
                                try
                                {
                                    if (tailArc.StartPoint.DistanceTo(_p1) < tailArc.EndPoint.DistanceTo(_p1))
                                    {
                                        tailArc.StartAngle = paramAtBase;
                                    }
                                    else
                                    {
                                        tailArc.EndAngle = paramAtBase;
                                    }
                                    draw.Geometry.Draw(tailArc);
                                }
                                catch { }
                            }
                        }
                    }
                }

                // ── 4. Text box (same as V1) ──────────────────────────────────────────────
                using (var mtext = ArcLeaderBoxService.CreateBox(_p3, _settings.DefaultText, _settings))
                {
                    if (mtext != null)
                        draw.Geometry.Draw(mtext);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
