using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using RCS.CustomLeader.Abstractions;

namespace RCS.CustomLeader.Core.Geometry
{
    public static class ArcLeaderBoxService
    {
        public static MText CreateBox(Point3d position, string text, ArcLeaderSettings settings)
        {
            var mtext = new MText();
            mtext.Location = position;
            mtext.Contents = string.IsNullOrEmpty(text) ? "INTEX" : text;
            mtext.TextHeight = settings.TextHeight;
            return mtext;
        }
    }
}
