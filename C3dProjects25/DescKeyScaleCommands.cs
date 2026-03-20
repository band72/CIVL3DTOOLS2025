using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace RCS.C3D2025.Tools
{
    public class DescKeyScaleCommands
    {
        [CommandMethod("RCS_FIX_DESC_KEY_SCALE")]
        public void RCS_FIX_DESC_KEY_SCALE()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            int setsTouched = 0, keysTouched = 0, keysFailed = 0, setsFailed = 0;
            string logPath = @"C:\Temp\c3doutput.txt";

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
            }
            catch { }

            using (doc.LockDocument())
            using (var tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    var setIds = PointDescriptionKeySetCollection.GetPointDescriptionKeySets(doc.Database);

                    if (setIds == null || setIds.Count == 0)
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now}] No description key sets found.\n");
                        ed.WriteMessage("\nNo description key sets found.");
                        return;
                    }

                    foreach (ObjectId setId in setIds)
                    {
                        try
                        {
                            if (tr.GetObject(setId, OpenMode.ForWrite) is PointDescriptionKeySet set)
                            {
                                setsTouched++;

                                var keyIds = set.GetPointDescriptionKeyIds();
                                foreach (ObjectId keyId in keyIds)
                                {
                                    try
                                    {
                                        if (tr.GetObject(keyId, OpenMode.ForWrite) is PointDescriptionKey key)
                                        {
                                            key.FixedScaleFactor = 0.02;
                                            key.ApplyFixedScaleFactor = true;
                                            key.ApplyDrawingScale = true;
                                            key.ApplyScaleParameter = false;
                                            key.ApplyScaleXY = true;
                                            
                                            keysTouched++;
                                        }
                                        else
                                        {
                                            keysFailed++;
                                            File.AppendAllText(logPath, $"[{DateTime.Now}] Null PointDescriptionKey in set {set.Name}.\n");
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        keysFailed++;
                                        File.AppendAllText(logPath, $"[{DateTime.Now}] Failed to access key in set {set.Name}: {ex.Message}\n");
                                    }
                                }
                            }
                            else
                            {
                                setsFailed++;
                                File.AppendAllText(logPath, $"[{DateTime.Now}] Failed to open key set with ID {setId}.\n");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            setsFailed++;
                            File.AppendAllText(logPath, $"[{DateTime.Now}] Exception accessing key set {setId}: {ex.Message}\n");
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($@"\nDone. Sets updated: {setsTouched}, Sets failed: {setsFailed}, Keys updated: {keysTouched}, Keys failed: {keysFailed}");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Completed. Sets updated: {setsTouched}, failed: {setsFailed}. Keys updated: {keysTouched}, failed: {keysFailed}\n");
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\nERROR: {ex.Message}");
                    File.AppendAllText(logPath, $"[{DateTime.Now}] Fatal error: {ex.Message}\n");
                }
            }
        }
    }
}
