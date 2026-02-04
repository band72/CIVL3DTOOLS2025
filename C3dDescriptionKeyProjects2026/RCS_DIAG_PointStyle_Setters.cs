// RCS_DIAG_POINTSTYLE_SETTERS.cs
// Diagnostic command to discover PointStyle display/layer setters
// Civil 3D 2025 / .NET Framework 4.8 safe

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.Reflection;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025
{
    public static class RcsDiagPointStyleSetters
    {
        [CommandMethod("RCS_DIAG_POINTSTYLE_SETTERS")]
        public static void RCS_DIAG_POINTSTYLE_SETTERS()
        {
            RcsError.RunCommandSafe("RCS_DIAG_POINTSTYLE_SETTERS", () =>
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                Editor ed = doc.Editor;
                Database db = doc.Database;

                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    CivilDocument civDoc = CivilApplication.ActiveDocument;
                    if (civDoc == null)
                        throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                    const int MAX_STYLES = 50;
                    int dumped = 0;
                    int scanned = 0;

                    foreach (ObjectId id in civDoc.Styles.PointStyles)
                    {
                        scanned++;
                        if (dumped >= MAX_STYLES) break;
                        if (id.IsNull || id.IsErased) continue;

                        Autodesk.Civil.DatabaseServices.Styles.PointStyle ps = null;
                        try
                        {
                            ps = tr.GetObject(id, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Styles.PointStyle;
                        }
                        catch { continue; }

                        if (ps == null) continue;
                        dumped++;

                        string psName = SafeName(ps);
                        RcsError.Info(ed, "====================================================");
                        RcsError.Info(ed, "POINTSTYLE SETTER DIAG: '" + psName + "'  Id=" + id);

                        MethodInfo[] methods = ps.GetType().GetMethods(
                            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

                        foreach (MethodInfo mi in methods)
                        {
                            if (mi == null) continue;
                            string name = mi.Name ?? "";

                            if (name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0 &&
                                name.IndexOf("Display", StringComparison.OrdinalIgnoreCase) < 0 &&
                                name.IndexOf("Style", StringComparison.OrdinalIgnoreCase) < 0)
                                continue;

                            if (!name.StartsWith("Set", StringComparison.OrdinalIgnoreCase) &&
                                name.IndexOf("Layer", StringComparison.OrdinalIgnoreCase) < 0)
                                continue;

                            ParameterInfo[] pars = mi.GetParameters();
                            string sig = BuildSignature(mi);
                            RcsError.Info(ed, "  " + sig);

                            for (int i = 0; i < pars.Length; i++)
                            {
                                ParameterInfo p = pars[i];
                                Type pt = p.ParameterType;
                                string hint =
                                    pt == typeof(ObjectId) ? "ObjectId" :
                                    pt == typeof(string) ? "string" :
                                    (pt.IsEnum ? "enum " + pt.FullName : pt.FullName);

                                RcsError.Info(ed,
                                    "     arg" + i + ": " + p.Name + " : " + hint);
                            }
                        }
                    }

                    RcsError.Info(ed,
                        "RCS_DIAG_POINTSTYLE_SETTERS complete. scanned=" +
                        scanned + " dumped=" + dumped +
                        "  log=" + RcsError.LogPath);

                    tr.Commit();
                }
            });
        }

        private static string BuildSignature(MethodInfo mi)
        {
            string ret = mi.ReturnType != null ? mi.ReturnType.Name : "void";
            var ps = mi.GetParameters();
            List<string> args = new List<string>();
            foreach (var p in ps)
            {
                args.Add(
                    (p.ParameterType != null ? p.ParameterType.Name : "<?>") +
                    " " + p.Name);
            }
            return ret + " " + mi.Name + "(" + string.Join(", ", args) + ")";
        }

        private static string SafeName(object o)
        {
            try
            {
                PropertyInfo pi = o.GetType().GetProperty(
                    "Name",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi != null)
                    return pi.GetValue(o, null) as string ?? "(unnamed)";
            }
            catch { }
            return o.GetType().Name;
        }
    }
}
