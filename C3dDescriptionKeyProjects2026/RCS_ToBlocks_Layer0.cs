using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace RCS.C3D2026
{
    public class RCS_Blocks_ToLayer0
    {
        private const string LogPath = @"C:\temp\c3d_blocks_layer0.log";

        // ------------------------------------------------------------
        // COMMAND
        // ------------------------------------------------------------
        [CommandMethod("RCS_SET_ALL_BLOCKS_TO_LAYER0")]
        public void SetAllBlocksToLayer0()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            Directory.CreateDirectory(@"C:\temp");
            Log("=== START RCS_SET_ALL_BLOCKS_TO_LAYER0 ===");

            int brChanged = 0;
            int defEntChanged = 0;
            int failed = 0;

            try
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    ObjectId layer0Id = db.LayerZero;
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    // ------------------------------------------------
                    // 1) ALL BlockReferences in Model + Paper space
                    // ------------------------------------------------
                    foreach (ObjectId btrId in bt)
                    {
                        var space = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        if (!space.IsLayout && btrId != bt[BlockTableRecord.ModelSpace])
                            continue;

                        foreach (ObjectId entId in space)
                        {
                            try
                            {
                                var br = tr.GetObject(entId, OpenMode.ForWrite, false) as BlockReference;
                                if (br == null) continue;

                                if (br.LayerId != layer0Id)
                                {
                                    br.LayerId = layer0Id;
                                    brChanged++;
                                }
                            }
                            catch
                            {
                                failed++;
                            }
                        }
                    }

                    // ------------------------------------------------
                    // 2) ALL entities inside ALL block definitions
                    // ------------------------------------------------
                    foreach (ObjectId btrId in bt)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                        // Skip layouts (ModelSpace / PaperSpace)
                        if (btr.IsLayout) continue;

                        if (!btr.IsWriteEnabled)
                            btr.UpgradeOpen();

                        foreach (ObjectId entId in btr)
                        {
                            try
                            {
                                var ent = tr.GetObject(entId, OpenMode.ForWrite, false) as Entity;
                                if (ent == null) continue;

                                if (ent.LayerId != layer0Id)
                                {
                                    ent.LayerId = layer0Id;
                                    defEntChanged++;
                                }
                            }
                            catch
                            {
                                failed++;
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage(
                    $"\nRCS_SET_ALL_BLOCKS_TO_LAYER0 completed." +
                    $"\nBlockReferences changed: {brChanged}" +
                    $"\nBlock definition entities changed: {defEntChanged}" +
                    $"\nFailures: {failed}" +
                    $"\nLog: {LogPath}\n"
                );

                Log($"SUMMARY brChanged={brChanged}, defEntChanged={defEntChanged}, failed={failed}");
            }
            catch (Exception ex)
            {
                Log("FATAL ERROR", ex);
                ed.WriteMessage($"\nERROR: {ex.Message}\nSee {LogPath}\n");
            }
        }

        private static void Log(string msg, Exception ex = null)
        {
            try
            {
                using (var sw = new StreamWriter(LogPath, true))
                {
                    sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {msg}");
                    if (ex != null)
                        sw.WriteLine($"{ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
                }
            }
            catch { }
        }
    }
}
