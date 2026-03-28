using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using RCS.CustomLeader.Abstractions;
using RCS.CustomLeader.Core.Geometry;
using RCS.CustomLeader.Core.Persistence;

namespace RCS.CustomLeader.Core.Builders
{
    public class ArcLeaderBuilder : IArcLeaderBuilder
    {
        private readonly Database _db;

        public ArcLeaderBuilder(Database db)
        {
            _db = db ?? throw new ArgumentNullException(nameof(db));
        }

        public ArcLeaderDefinition Build(Point3d p1, Point3d p2, Point3d p3, string text, ArcLeaderSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            var def = new ArcLeaderDefinition 
            { 
                HeadPoint = p1, 
                ThroughPoint = p2, 
                BoxPoint = p3, 
                TextValue = text ?? string.Empty 
            };

            using (var tr = _db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(_db.CurrentSpaceId, OpenMode.ForWrite);

                // Ensure layers exist safely
                EnsureLayerExists(tr, settings.ArcLayer);
                EnsureLayerExists(tr, settings.HeadLayer);
                EnsureLayerExists(tr, settings.TextLayer);

                // 1. Generate Arc
                var arc = ArcLeaderGeometryService.CreateArc(p1, p2, p3);
                double tangentAngle = 0.0;

                if (arc != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.ArcLayer))
                        arc.Layer = settings.ArcLayer;

                    def.ArcId = btr.AppendEntity(arc);
                    tr.AddNewlyCreatedDBObject(arc, true);

                    tangentAngle = ArcLeaderGeometryService.GetChordAngleAt(arc, p1, settings.TextHeight * 1.5);
                }

                // 2. Generate Head Block (Safely)
                var headInfo = ArcLeaderGeometryService.CreateHeadBlock(p1, tangentAngle, settings);
                if (headInfo != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.HeadLayer))
                        headInfo.Layer = settings.HeadLayer;

                    def.HeadBlockId = btr.AppendEntity(headInfo);
                    tr.AddNewlyCreatedDBObject(headInfo, true);
                }

                // 3. Generate MText Box
                var mtext = ArcLeaderBoxService.CreateBox(p3, text, settings);
                if (mtext != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.TextLayer))
                        mtext.Layer = settings.TextLayer;

                    def.TextId = btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);

                    // Mathematically trim the Arc so it snaps to the edge of the text box
                    if (arc != null)
                    {
                        ArcLeaderBoxService.TrimArcToMText(arc, mtext, settings);
                    }
                }

                // 4. Group them safely
                var ids = new List<ObjectId>();
                if (def.ArcId != ObjectId.Null) ids.Add(def.ArcId);
                if (def.HeadBlockId != ObjectId.Null) ids.Add(def.HeadBlockId);
                if (def.TextId != ObjectId.Null) ids.Add(def.TextId);
                
                if (ids.Count > 0)
                {
                    def.GroupId = GroupService.CreateGroup(_db, tr, ids);
                }

                tr.Commit();
            }

            return def;
        }

        /// <summary>
        /// V2 Arc Leader: identical to V1 (main arc + arrowhead + text box) PLUS a second arc.
        /// The second arc runs from the arrowhead's back end (arrowBase) through the midpoint
        /// of that chord to the trimmed endpoint of the first arc (the text-box end).
        /// This creates a "double-arc" leader with a small curved tail behind the arrowhead.
        /// </summary>
        public ArcLeaderDefinition BuildV2(Point3d p1, Point3d p2, Point3d p3, string text, ArcLeaderSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            var def = new ArcLeaderDefinition
            {
                HeadPoint    = p1,
                ThroughPoint = p2,
                BoxPoint     = p3,
                TextValue    = text ?? string.Empty,
                StyleName    = "V2"
            };

            using (var tr = _db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(_db.CurrentSpaceId, OpenMode.ForWrite);

                EnsureLayerExists(tr, settings.ArcLayer);
                EnsureLayerExists(tr, settings.HeadLayer);
                EnsureLayerExists(tr, settings.TextLayer);

                // ── Step 1: Main arc (computed and added to DB) ──────────────────────────
                var arc = ArcLeaderGeometryService.CreateArc(p1, p2, p3);
                double tangentAngle = 0.0;

                if (arc != null)
                {
                    tangentAngle = ArcLeaderGeometryService.GetChordAngleAt(arc, p1, settings.TextHeight * 1.5);
                }

                // ── Step 2: Arrowhead (same as V1) ───────────────────────────────────────
                var headInfo = ArcLeaderGeometryService.CreateHeadBlock(p1, tangentAngle, settings);
                if (headInfo != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.HeadLayer))
                        headInfo.Layer = settings.HeadLayer;

                    def.HeadBlockId = btr.AppendEntity(headInfo);
                    tr.AddNewlyCreatedDBObject(headInfo, true);
                }

                // ── Step 3: Text box (same as V1) + trim main arc ────────────────────────
                var mtext = ArcLeaderBoxService.CreateBox(p3, text, settings);
                if (mtext != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.TextLayer))
                        mtext.Layer = settings.TextLayer;

                    def.TextId = btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);

                    if (arc != null)
                        ArcLeaderBoxService.TrimArcToMText(arc, mtext, settings);
                }

                // ── Step 4: Second arc (V2 addition) ─────────────────────────────────────
                // arrowBase = back end of the arrowhead solid
                // firstArcEnd = the trimmed endpoint of the main arc closest to p3 (text box)
                // secondArc = arrowBase → chord-midpoint (slight perpendicular offset) → firstArcEnd
                if (arc != null)
                {
                    double  arrowLength = settings.TextHeight * 1.5;
                    var     tangentDir  = new Vector3d(Math.Cos(tangentAngle), Math.Sin(tangentAngle), 0);
                    Point3d arrowBase   = p1 + tangentDir * arrowLength;

                    // Build second arc: perfectly traces the first arc (same circle)
                    Arc tailArc = (Arc)arc.Clone();
                    
                    // Find the exact point on the circle closest to the arrowhead base
                    Point3d trueBaseOnArc = tailArc.GetClosestPointTo(arrowBase, false);
                    double paramAtBase = tailArc.GetParameterAtPoint(trueBaseOnArc);

                    double arcLength = tailArc.GetDistanceAtParameter(tailArc.EndParam);
                    if (arcLength <= arrowLength)
                    {
                        // The entire arc is physically covered by the solid arrowhead.
                        tailArc.Dispose();
                        tailArc = null;
                    }
                    else
                    {
                        try
                        {
                            if (tailArc.StartPoint.DistanceTo(p1) < tailArc.EndPoint.DistanceTo(p1))
                            {
                                tailArc.StartAngle = paramAtBase;
                            }
                            else
                            {
                                tailArc.EndAngle = paramAtBase;
                            }
                        }
                        catch
                        {
                            tailArc.Dispose();
                            tailArc = null;
                        }
                    }
                    if (tailArc != null)
                    {
                        if (!string.IsNullOrWhiteSpace(settings.ArcLayer))
                            tailArc.Layer = settings.ArcLayer;

                        def.TailArcId = btr.AppendEntity(tailArc);
                        tr.AddNewlyCreatedDBObject(tailArc, true);
                    }
                }

                if (arc != null)
                {
                    arc.Dispose();
                }

                // ── Step 5: Group all entities ────────────────────────────────────────────
                var ids = new List<ObjectId>();
                if (def.ArcId       != ObjectId.Null) ids.Add(def.ArcId);
                if (def.HeadBlockId != ObjectId.Null) ids.Add(def.HeadBlockId);
                if (def.TextId      != ObjectId.Null) ids.Add(def.TextId);
                if (def.TailArcId   != ObjectId.Null) ids.Add(def.TailArcId);

                if (ids.Count > 0)
                    def.GroupId = GroupService.CreateGroup(_db, tr, ids);

                tr.Commit();
            }

            return def;
        }

        public void Rebuild(ArcLeaderDefinition definition, ArcLeaderSettings settings)
        {
            if (definition == null) throw new ArgumentNullException(nameof(definition));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            // Erase the old component entities, then reconstruct via Build() with the original geometry.
            using (var tr = _db.TransactionManager.StartTransaction())
            {
                EraseIfValid(tr, definition.ArcId);
                EraseIfValid(tr, definition.HeadBlockId);
                EraseIfValid(tr, definition.TextId);
                EraseIfValid(tr, definition.GroupId);
                tr.Commit();
            }

            // Clear stale IDs before rebuilding
            definition.ArcId       = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null;
            definition.HeadBlockId = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null;
            definition.TextId      = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null;
            definition.GroupId     = Autodesk.AutoCAD.DatabaseServices.ObjectId.Null;

            // Reconstruct using stored geometry
            var rebuilt = Build(definition.HeadPoint, definition.ThroughPoint, definition.BoxPoint,
                                definition.TextValue, settings);

            // Copy new IDs back into the passed-in definition so callers stay in sync
            definition.ArcId       = rebuilt.ArcId;
            definition.HeadBlockId = rebuilt.HeadBlockId;
            definition.TextId      = rebuilt.TextId;
            definition.GroupId     = rebuilt.GroupId;
        }

        private static void EraseIfValid(Transaction tr, Autodesk.AutoCAD.DatabaseServices.ObjectId id)
        {
            if (id == Autodesk.AutoCAD.DatabaseServices.ObjectId.Null || id.IsErased) return;
            try
            {
                var obj = tr.GetObject(id, OpenMode.ForWrite, false);
                obj?.Erase();
            }
            catch { /* entity may already be gone */ }
        }

        private void EnsureLayerExists(Transaction tr, string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return;

            var lt = (LayerTable)tr.GetObject(_db.LayerTableId, OpenMode.ForRead);
            if (!lt.Has(layerName))
            {
                lt.UpgradeOpen();
                var ltr = new LayerTableRecord { Name = layerName };
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }
    }
}
