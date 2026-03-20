using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;

namespace RCS.CustomLeader.Core.Persistence
{
    public static class GroupService
    {
        public static ObjectId CreateGroup(Database db, Transaction tr, IEnumerable<ObjectId> entities)
        {
            var dict = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);
            var group = new Group("Custom Leader Group", true);
            
            foreach (var id in entities)
                group.Append(id);

            var groupId = dict.SetAt("*", group);
            tr.AddNewlyCreatedDBObject(group, true);
            
            return groupId;
        }
    }
}
