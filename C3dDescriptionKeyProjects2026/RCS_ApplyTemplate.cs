using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace RCS.C3D2025
{
    public class RcsApplyTemplate
    {
        // Isolation safeguards (prevent "hang" from excessive recursion)
        private static int _failCount = 0;
        private const int MaxFails = 50;
        private const int MaxDepth = 12;

        /// <summary>
        /// Command:
        ///   RCS_APPLY_TEMPLATE
        /// Purpose:
        ///   Force-merge a selected DWT/DWG template into the current drawing by cloning named objects
        ///   (layers, linetypes, styles, blocks, etc.) with DuplicateRecordCloning.Replace (template wins).
        /// Logging:
        ///   C:\temp\c3doutput.txt
        /// </summary>
        [CommandMethod("RCS_APPLY_TEMPLATE")]
        public static void RCS_APPLY_TEMPLATE()
        {
            RcsError.RunCommandSafe("RCS_APPLY_TEMPLATE", () =>
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                if (doc == null) throw new InvalidOperationException("No active document.");

                var ed = doc.Editor;

                var pfo = new PromptOpenFileOptions("\nSelect template drawing (.dwt/.dwg) to merge into current drawing:")
                {
                    Filter = "AutoCAD Drawing (*.dwg;*.dwt)|*.dwg;*.dwt|All Files (*.*)|*.*",
                    DialogCaption = "Select DWT/DWG Template"
                };

                var pr = ed.GetFileNameForOpen(pfo);
                if (pr.Status != PromptStatus.OK || string.IsNullOrWhiteSpace(pr.StringResult))
                {
                    RcsError.Warn(ed, "User canceled file selection.");
                    return;
                }

                var path = pr.StringResult.Trim();
                ApplyTemplateIntoCurrentDrawing(doc, path);

                ed.WriteMessage($"\nTemplate merge finished. (See {RcsError.LogPath})");
            });
        }

        public static void ApplyTemplateIntoCurrentDrawing(Document doc, string templatePath)
        {
            var ed = doc.Editor;

            // reset isolation counters each run
            _failCount = 0;

            if (string.IsNullOrWhiteSpace(templatePath))
                throw new ArgumentException("Template path is empty.");

            if (!File.Exists(templatePath))
                throw new FileNotFoundException("Template file not found.", templatePath);

            var ext = Path.GetExtension(templatePath).ToLowerInvariant();
            if (ext != ".dwg" && ext != ".dwt")
                throw new ArgumentException("Template must be a .dwg or .dwt file.");

            RcsError.Info(ed, $"Source template: {templatePath}");

            using (var sourceDb = new Database(false, true))
            {
                // 1) Read source database
                SafeStep(ed, "ReadDwgFile", () =>
                {
                    sourceDb.ReadDwgFile(templatePath, FileShare.Read, allowCPConversion: true, password: "");
                    sourceDb.CloseInput(true);
                });

                // 2) Lock and clone into target
                using (doc.LockDocument())
                {
                    var targetDb = doc.Database;

                    using (var trTarget = targetDb.TransactionManager.StartTransaction())
                    using (var trSource = sourceDb.TransactionManager.StartTransaction())
                    {
                        // Collect per-category so we can clone to correct owners (prevents WblockCloneObjects failures/hangs)
                        var layerIds = new ObjectIdCollection();
                        var ltypeIds = new ObjectIdCollection();
                        var textStyleIds = new ObjectIdCollection();
                        var dimStyleIds = new ObjectIdCollection();
                        var ucsIds = new ObjectIdCollection();
                        var blockBtrIds = new ObjectIdCollection();
                        var mleaderStyleIds = new ObjectIdCollection();
                        var tableStyleIds = new ObjectIdCollection();

                        AddSymbolTableSafe(trSource, sourceDb.LayerTableId, layerIds, "LayerTable");
                        AddSymbolTableSafe(trSource, sourceDb.LinetypeTableId, ltypeIds, "LinetypeTable");
                        AddSymbolTableSafe(trSource, sourceDb.TextStyleTableId, textStyleIds, "TextStyleTable");
                        AddSymbolTableSafe(trSource, sourceDb.DimStyleTableId, dimStyleIds, "DimStyleTable");
                        AddSymbolTableSafe(trSource, sourceDb.UcsTableId, ucsIds, "UcsTable");

                        AddDictionaryEntriesSafe(trSource, sourceDb.MLeaderStyleDictionaryId, mleaderStyleIds, "MLeaderStyleDictionary");
                        AddDictionaryEntriesSafe(trSource, sourceDb.TableStyleDictionaryId, tableStyleIds, "TableStyleDictionary");

                        // Block defs
                        SafeStep(ed, "Collect BlockTableRecords", () =>
                        {
                            var bt = (BlockTable)trSource.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);
                            int added = 0;

                            foreach (ObjectId btrId in bt)
                            {
                                if (btrId.IsNull || btrId.IsErased) continue;

                                var btr = (BlockTableRecord)trSource.GetObject(btrId, OpenMode.ForRead);

                                // avoid anonymous/layout blocks (common offenders + not useful as "template content")
                                if (btr.IsAnonymous) continue;
                                if (btr.IsLayout) continue;


                                // skip spaces explicitly
                                if (string.Equals(btr.Name, BlockTableRecord.ModelSpace, StringComparison.OrdinalIgnoreCase)) continue;
                                if (string.Equals(btr.Name, BlockTableRecord.PaperSpace, StringComparison.OrdinalIgnoreCase)) continue;

                                // skip xref/dependent blocks (can cause demand-loading / long post-merge rebuilds)
                                if (btr.IsFromExternalReference) continue;
                                if (btr.IsDependent) continue;

                                // skip xref-names like "XREFNAME|BLOCKNAME"
                                if (!string.IsNullOrWhiteSpace(btr.Name) && btr.Name.Contains("|")) continue;
                                blockBtrIds.Add(btrId);
                                added++;
                            }

                            RcsError.Info(ed, $"BlockTableRecords added: {added}");
                        });

                        // 3) Clone to correct owner with bounded isolation
                        SafeStep(ed, "Clone Template Content", () =>
                        {
                            // Progress meter so Civil doesn't look "hung"
                            var pm = new ProgressMeter();
                            pm.SetLimit(8);
                            pm.Start("Merging template...");
                            try
                            {
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, layerIds, targetDb.LayerTableId, "LAYERS", DuplicateRecordCloning.Replace, batchSize: 150);
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, ltypeIds, targetDb.LinetypeTableId, "LINETYPES", DuplicateRecordCloning.Replace, batchSize: 150);
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, textStyleIds, targetDb.TextStyleTableId, "TEXTSTYLES", DuplicateRecordCloning.Replace, batchSize: 150);
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, dimStyleIds, targetDb.DimStyleTableId, "DIMSTYLES", DuplicateRecordCloning.Replace, batchSize: 150);
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, ucsIds, targetDb.UcsTableId, "UCS", DuplicateRecordCloning.Replace, batchSize: 150);

                                // Blocks are heavy; use smaller batches
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, blockBtrIds, targetDb.BlockTableId, "BLOCK_DEFS", DuplicateRecordCloning.Ignore, batchSize: 25);

                                // Dict entries
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, mleaderStyleIds, targetDb.MLeaderStyleDictionaryId, "MLEADER_STYLES", DuplicateRecordCloning.Replace, batchSize: 50);
                                pm.MeterProgress(); CloneNamedObjectsRobust(targetDb, trSource, tableStyleIds, targetDb.TableStyleDictionaryId, "TABLE_STYLES", DuplicateRecordCloning.Replace, batchSize: 50);
                            }
                            finally
                            {
                                pm.Stop();
                            }

                            RcsError.Info(ed, $"Clone summary: FailCount={_failCount} (MaxFails={MaxFails}, MaxDepth={MaxDepth})");
                        });

                        trSource.Commit();
                        trTarget.Commit();
                    }

                    // REGEN after massive cloning can appear like a hang.
                    // Leave it manual; user can run REGEN if needed.
                    RcsError.Info(ed, "Skipping automatic REGEN. Run REGEN manually if needed.");
                }
            }
        }

        // ================== Robust Clone Engine ==================

        private static void CloneNamedObjectsRobust(
            Database targetDb,
            Transaction trSource,
            ObjectIdCollection ids,
            ObjectId ownerId,
            string label,
            DuplicateRecordCloning drc,
            int batchSize = 150)
        {
            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;

            if (ids == null || ids.Count == 0)
            {
                RcsError.Info(ed, $"{label}: nothing to clone.");
                return;
            }

            RcsError.Info(ed, $"{label}: cloning {ids.Count} record(s) (batch={batchSize}).");

            for (int i = 0; i < ids.Count; i += batchSize)
            {
                if (_failCount >= MaxFails)
                {
                    RcsError.Warn(ed, $"{label}: aborting further work (MaxFails={MaxFails}).");
                    return;
                }

                var batch = new ObjectIdCollection();
                for (int j = i; j < Math.Min(i + batchSize, ids.Count); j++)
                    batch.Add(ids[j]);

                TryCloneBounded(targetDb, trSource, batch, ownerId, label, drc, depth: 0);
            }
        }

        private static void TryCloneBounded(
            Database targetDb,
            Transaction trSource,
            ObjectIdCollection ids,
            ObjectId ownerId,
            string label,
            DuplicateRecordCloning drc,
            int depth)
        {
            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;

            if (_failCount >= MaxFails)
                return;

            if (depth > MaxDepth)
            {
                _failCount++;
                RcsError.Warn(ed, $"{label}: depth limit reached (MaxDepth={MaxDepth}). Skipping {ids.Count} item(s).");
                return;
            }

            try
            {
                var mapping = new IdMapping();
                targetDb.WblockCloneObjects(
                    ids,
                    ownerId,
                    mapping,
                    drc,
                    deferTranslation: false
                );

                return;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception aex)
            {
                _failCount++;
                RcsError.Warn(ed, $"{label}: clone failed ({aex.ErrorStatus}) depth={depth} count={ids.Count} :: {aex.Message}");
            }
            catch (System.Exception ex)
            {
                _failCount++;
                RcsError.Fail(ed, $"{label}: STEP FAIL | {ex.Message}", ex as Autodesk.AutoCAD.Runtime.Exception ?? null);
                throw;
            }

            // isolate only if useful
            if (ids.Count <= 1)
            {
                LogBadSingle(trSource, ids, label);
                return;
            }

            int mid = ids.Count / 2;
            TryCloneBounded(targetDb, trSource, Sub(ids, 0, mid), ownerId, label, drc, depth + 1);
            TryCloneBounded(targetDb, trSource, Sub(ids, mid, ids.Count), ownerId, label, drc, depth + 1);
        }

        private static void LogBadSingle(Transaction trSource, ObjectIdCollection ids, string label)
        {
            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;

            try
            {
                if (ids == null || ids.Count == 0)
                {
                    RcsError.Fail(ed, $"{label}: SKIP single bad record (empty id collection).");
                    return;
                }

                var dbo = trSource.GetObject(ids[0], OpenMode.ForRead, false);
                string name = "";
                if (dbo is SymbolTableRecord str) name = str.Name;

                RcsError.Fail(ed, $"{label}: SKIP single bad record. Type={dbo.GetType().Name} Name='{name}' ObjectId={ids[0]}");
            }
            catch
            {
                RcsError.Fail(ed, $"{label}: SKIP single bad record. (unreadable)");
            }
        }

        private static ObjectIdCollection Sub(ObjectIdCollection src, int startInclusive, int endExclusive)
        {
            var col = new ObjectIdCollection();
            for (int i = startInclusive; i < endExclusive; i++)
                col.Add(src[i]);
            return col;
        }

        // ================== Original Helpers (kept/expanded) ==================

        private static void SafeStep(Editor ed, string stepName, Action action)
        {
            try
            {
                RcsError.Info(ed, $"STEP START: {stepName}");
                action();
                RcsError.Info(ed, $"STEP OK: {stepName}");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception aex)
            {
                RcsError.Fail(ed, $"STEP FAIL: {stepName} | AutoCAD={aex.ErrorStatus} | {aex.Message}", aex);
                throw;
            }
            catch (System.Exception ex)
            {
                RcsError.Fail(ed, $"STEP FAIL: {stepName} | {ex.Message}", ex as Autodesk.AutoCAD.Runtime.Exception);
                throw;
            }
        }

        private static void AddSymbolTableSafe(Transaction trSource, ObjectId tableId, ObjectIdCollection idsToClone, string label)
        {
            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;

            SafeStep(ed!, $"Collect {label}", () =>
            {
                var table = (SymbolTable)trSource.GetObject(tableId, OpenMode.ForRead);

                int added = 0;
                foreach (ObjectId id in table)
                {
                    if (id.IsNull || id.IsErased) continue;
                    idsToClone.Add(id);
                    added++;
                }

                RcsError.Info(ed, $"{label} added: {added}");
            });
        }

        private static void AddDictionaryEntriesSafe(Transaction trSource, ObjectId dictId, ObjectIdCollection idsToClone, string label)
        {
            var ed = Application.DocumentManager.MdiActiveDocument?.Editor;

            SafeStep(ed!, $"Collect {label}", () =>
            {
                if (dictId.IsNull)
                {
                    RcsError.Warn(ed, $"{label} Id is null.");
                    return;
                }

                var dict = trSource.GetObject(dictId, OpenMode.ForRead) as DBDictionary;
                if (dict == null)
                {
                    RcsError.Warn(ed, $"{label} is not a DBDictionary.");
                    return;
                }

                int added = 0;
                foreach (DBDictionaryEntry entry in dict)
                {
                    var id = entry.Value;
                    if (id.IsNull || id.IsErased) continue;
                    idsToClone.Add(id);
                    added++;
                }

                RcsError.Info(ed, $"{label} added: {added}");
            });
        }
    }
}
