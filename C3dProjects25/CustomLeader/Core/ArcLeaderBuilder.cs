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
            _db = db;
        }

        public ArcLeaderDefinition Build(Point3d p1, Point3d p2, Point3d p3, string text, ArcLeaderSettings settings)
        {
            var def = new ArcLeaderDefinition { HeadPoint = p1, ThroughPoint = p2, BoxPoint = p3, TextValue = text };

            using (var tr = _db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(_db.CurrentSpaceId, OpenMode.ForWrite);

                // Ensure layers exist
                EnsureLayerExists(tr, settings.ArcLayer);
                EnsureLayerExists(tr, settings.HeadLayer);
                EnsureLayerExists(tr, settings.TextLayer);

                // 1. Generate Arc
                var arc = ArcLeaderGeometryService.CreateArc(p1, p2, p3);
                arc.Layer = settings.ArcLayer;
                def.ArcId = btr.AppendEntity(arc);
                tr.AddNewlyCreatedDBObject(arc, true);

                // 2. Generate Head Block (Safely)
                var tangentAngle = ArcLeaderGeometryService.GetStartTangentAngle(arc);
                var headInfo = ArcLeaderGeometryService.CreateHeadBlock(p1, tangentAngle, settings);
                if (headInfo != null)
                {
                    def.HeadBlockId = btr.AppendEntity(headInfo);
                    tr.AddNewlyCreatedDBObject(headInfo, true);
                }

                // 3. Generate MText Box
                var mtext = ArcLeaderBoxService.CreateBox(p3, text, settings);
                if (mtext != null)
                {
                    def.TextId = btr.AppendEntity(mtext);
                    tr.AddNewlyCreatedDBObject(mtext, true);
                }

                // 4. Group them
                var ids = new List<ObjectId> { def.ArcId };
                if (def.HeadBlockId != ObjectId.Null) ids.Add(def.HeadBlockId);
                if (def.TextId != ObjectId.Null) ids.Add(def.TextId);
                
                def.GroupId = GroupService.CreateGroup(_db, tr, ids);

                tr.Commit();
            }
            return def;
        }

        public void Rebuild(ArcLeaderDefinition definition, ArcLeaderSettings settings)
        {
            // TODO: Open existing entities by ID, recalculate geometry, and update them in place
        }

        private void EnsureLayerExists(Transaction tr, string layerName)
        {
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
