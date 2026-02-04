using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices.Styles;
using System;
using System.Collections.Generic;
using System.IO;

namespace RCS.C3D2025
{
    /// <summary>
    /// RCS Delete All Point Styles (Civil 3D 2025 / .NET 8)
    ///
    /// Command:
    ///   RCS_DELETE_ALL_POINTSTYLES
    ///
    /// Notes:
    /// - Collects PointStyle ObjectIds first, then attempts Erase() in a write transaction.
    /// - Skips "Standard" (add more protected names if needed).
    /// - Logs to C:\temp\c3doutput.txt
    /// </summary>
    public class RcsDeleteAllPointStylesNet8
    {
        private const string LogPath = @"C:\temp\c3doutput.txt";

        [CommandMethod("RCS_DELETE_ALL_POINTSTYLES")]
        public static void RCS_DELETE_ALL_POINTSTYLES()
        {
            RunCommandSafe("RCS_DELETE_ALL_POINTSTYLES", () =>
            {
                var doc = Application.DocumentManager.MdiActiveDocument ?? throw new InvalidOperationException("No active drawing.");
                var ed = doc.Editor;
                var db = doc.Database;

                var civDoc = CivilApplication.ActiveDocument ?? throw new InvalidOperationException("CivilApplication.ActiveDocument is null.");

                if (!PromptYesNo(ed, "\nThis will attempt to ERASE all Point Styles in this drawing. Continue?", defaultYes: false))
                {
                    ed.WriteMessage("\nCanceled.\n");
                    return;
                }

                using (doc.LockDocument())
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ids = new List<ObjectId>();
                    foreach (ObjectId id in civDoc.Styles.PointStyles)
                    {
                        if (id.IsNull || id.IsErased) continue;
                        ids.Add(id);
                    }

                    int scanned = ids.Count;
                    int deleted = 0, protectedSkip = 0, inUseSkip = 0, failed = 0;

                    var protectedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    {
                        "Standard"
                    };

                    foreach (var id in ids)
                    {
                        try
                        {
                            var ps = tr.GetObject(id, OpenMode.ForWrite, false) as PointStyle;
                            if (ps == null) { failed++; continue; }

                            string name = SafeName(ps);

                            if (protectedNames.Contains(name))
                            {
                                protectedSkip++;
                                continue;
                            }

                            if (!ps.IsErased)
                            {
                                ps.Erase(true);
                                deleted++;
                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception aex)
                        {
                            Warn(ed, $"Delete failed for PointStyle id={id.Handle}: {aex.ErrorStatus} {aex.Message}");

                            if (aex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.ObjectIsReferenced ||
                                aex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotAllowedForThisProxy)
                            {
                                inUseSkip++;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            failed++;
                            Warn(ed, $"Delete failed for PointStyle id={id.Handle}: {ex.Message}");
                        }
                    }

                    tr.Commit();

                    Info(ed, $"DELETE_ALL_POINTSTYLES complete. scanned={scanned} deleted={deleted} protectedSkip={protectedSkip} inUseSkip={inUseSkip} failed={failed}");
                    ed.WriteMessage($"\nRCS_DELETE_ALL_POINTSTYLES: scanned={scanned} deleted={deleted} protectedSkip={protectedSkip} inUseSkip={inUseSkip} failed={failed}\nLog: {LogPath}\n");
                }
            });
        }

        private static bool PromptYesNo(Editor ed, string message, bool defaultYes)
        {
            try
            {
                var pko = new PromptKeywordOptions(message);
                pko.AllowNone = true;
                pko.Keywords.Add("No");
                pko.Keywords.Add("Yes");
                pko.Keywords.Default = defaultYes ? "Yes" : "No";
                var pr = ed.GetKeywords(pko);
                if (pr.Status == PromptStatus.OK)
                    return string.Equals(pr.StringResult, "Yes", StringComparison.OrdinalIgnoreCase);
                return defaultYes;
            }
            catch { return defaultYes; }
        }

        private static string SafeName(object o)
        {
            try
            {
                var pi = o.GetType().GetProperty("Name", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (pi != null)
                {
                    var v = pi.GetValue(o, null) as string;
                    if (!string.IsNullOrWhiteSpace(v)) return v;
                }
            }
            catch { }
            return o?.GetType().Name ?? "(null)";
        }

        private static void RunCommandSafe(string name, Action act)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            try
            {
                Info(ed, $"--- {name} START --- Drawing={(doc?.Name ?? "(null)")}");
                act();
                Info(ed, $"--- {name} END OK ---");
            }
            catch (System.Exception ex)
            {
                Error(ed, $"--- {name} FAILED --- {ex.Message}\n{ex}");
                try { ed?.WriteMessage($"\nERROR: {ex.Message}\n"); } catch { }
            }
        }

        private static void Info(Editor ed, string msg) => Write(ed, "INFO", msg);
        private static void Warn(Editor ed, string msg) => Write(ed, "WARN", msg);
        private static void Error(Editor ed, string msg) => Write(ed, "ERROR", msg);

        private static void Write(Editor ed, string lvl, string msg)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{lvl}] [DWG={(Application.DocumentManager.MdiActiveDocument?.Name ?? "(null)")}] {msg}\n");
            }
            catch { }
            try { ed?.WriteMessage($"\n[{lvl}] {msg}"); } catch { }
        }
    }
}
