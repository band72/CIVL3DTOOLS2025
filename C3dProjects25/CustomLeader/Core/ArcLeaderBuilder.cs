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

                    tangentAngle = ArcLeaderGeometryService.GetStartTangentAngle(arc);
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
        /// V2 Arc Leader: same interaction as V1 (head → through → box) but the final drawing
        /// contains only the arrowhead + a gently-curved tail arc + text box.
        /// The primary arc (p1→p2→p3) is computed in memory solely to derive the arrowhead
        /// tangent angle and the tail arc endpoint, then discarded.
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

                // ── Step 1: derive geometry from main arc (in-memory, NOT added to DB) ──────
                double  tangentAngle = 0.0;
                Point3d tailArcEnd   = p3; // safe fallback
                double  mainArcRadius = 0.0;

                var mainArc = ArcLeaderGeometryService.CreateArc(p1, p2, p3);
                if (mainArc != null)
                {
                    tangentAngle  = ArcLeaderGeometryService.GetStartTangentAngle(mainArc);
                    mainArcRadius = mainArc.Radius;

                    // Build a temporary MText to find where the main arc would be trimmed
                    using (var tempMtext = ArcLeaderBoxService.CreateBox(p3, text, settings))
                    {
                        ArcLeaderBoxService.TrimArcToMText(mainArc, tempMtext, settings);
                    }

                    // After trimming, the endpoint closest to p3 is where the tail arc terminates
                    tailArcEnd = mainArc.StartPoint.DistanceTo(p3) < mainArc.EndPoint.DistanceTo(p3)
                                 ? mainArc.StartPoint
                                 : mainArc.EndPoint;

                    mainArc.Dispose(); // Never added to DB — dispose manually
                }

                // ── Step 2: arrowhead at p1 (identical to V1) ─────────────────────────────
                var headInfo = ArcLeaderGeometryService.CreateHeadBlock(p1, tangentAngle, settings);
                if (headInfo != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.HeadLayer))
                        headInfo.Layer = settings.HeadLayer;

                    def.HeadBlockId = btr.AppendEntity(headInfo);
                    tr.AddNewlyCreatedDBObject(headInfo, true);
                }

                // ── Step 3: text box at p3 (identical to V1) ──────────────────────────────
                var mtext = ArcLeaderBoxService.CreateBox(p3, text, settings);
                if (mtext != null)
                {
                    if (!string.IsNullOrWhiteSpace(settings.TextLayer))
                        mtext.Layer = settings.TextLayer;

                    def.TextId = btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);

                    // ── Step 4: tail arc from p2 to tailArcEnd, gently curved ───────────
                    var tailArc = ArcLeaderGeometryService.CreateTailArc(p2, tailArcEnd, mainArcRadius);
                    if (tailArc != null)
                    {
                        if (!string.IsNullOrWhiteSpace(settings.ArcLayer))
                            tailArc.Layer = settings.ArcLayer;

                        // Trim tail arc so it doesn't overlap the text box
                        ArcLeaderBoxService.TrimArcToMText(tailArc, mtext, settings);

                        def.TailArcId = btr.AppendEntity(tailArc);
                        tr.AddNewlyCreatedDBObject(tailArc, true);
                    }
                }

                // ── Step 5: group all committed entities ───────────────────────────────────
                var ids = new List<ObjectId>();
                if (def.TailArcId   != ObjectId.Null) ids.Add(def.TailArcId);
                if (def.HeadBlockId != ObjectId.Null) ids.Add(def.HeadBlockId);
                if (def.TextId      != ObjectId.Null) ids.Add(def.TextId);

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
