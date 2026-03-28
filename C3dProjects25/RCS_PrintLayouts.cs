using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;

namespace RCS.C3D2025.Tools
{
    public class PrintLayoutsCommand
    {
        [CommandMethod("RCS_PRINT_MULTI_SHEETS")]
        public void PrintMultipleSheetsFromLayout()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 1. Ensure we are in Paper Space (Layout)
            if (db.TileMode || db.PaperSpaceVportId == ObjectId.Null)
            {
                ed.WriteMessage("\nError: This command must be run from the Paper Space Layout containing your sheets.");
                return;
            }

            try
            {
                // 2. Select the Sheets (Blocks or Polylines that border them)
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nSelect the Sheet boundaries (Titleblocks or Polylines) by windowing: ";
                
                PromptSelectionResult psr = ed.GetSelection(pso);
                if (psr.Status != PromptStatus.OK) return;

                List<Extents3d> sheetExtents = new List<Extents3d>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in psr.Value)
                    {
                        Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent != null)
                        {
                            sheetExtents.Add(ent.GeometricExtents);
                        }
                    }

                    if (sheetExtents.Count == 0)
                    {
                        ed.WriteMessage("\nNo valid boundaries selected.");
                        return;
                    }

                    // 3. Sort them Left-to-Right naturally so they print in chronological order
                    sheetExtents = sheetExtents.OrderBy(e => e.MinPoint.X).ToList();

                    // 4. Get Current Layout and its PlotSettings
                    LayoutManager layoutMgr = LayoutManager.Current;
                    Layout currentLayout = tr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

                    ed.WriteMessage($"\nFound {sheetExtents.Count} sheets. Initializing PlotEngine with current Page Setup...");

                    // 5. Execute Plotting
                    if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                    {
                        // Set background plotting to false temporarily to prevent threading collisions
                        short bgPlotStr = (short)Application.GetSystemVariable("BACKGROUNDPLOT");
                        Application.SetSystemVariable("BACKGROUNDPLOT", 0);
                        int plottedCount = 0;
                        PlotEngine engine = PlotFactory.CreatePublishEngine();
                        using (engine)
                        {
                            PlotProgressDialog progressDlg = new PlotProgressDialog(false, sheetExtents.Count, true);
                            using (progressDlg)
                            {
                                progressDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle, "RCS Multi-Sheet Plotter");
                                progressDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Batch");
                                progressDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                                progressDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Overall Batch Progress");
                                progressDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Current Sheet Progress");

                                progressDlg.OnBeginPlot();
                                progressDlg.IsVisible = true;
                                engine.BeginPlot(progressDlg, null);

                                // Construct the physical file path specifically for a single Multi-Page PDF
                                string pathName = doc.Name;
                                string dwgDir = string.IsNullOrEmpty(pathName) || !System.IO.Path.IsPathRooted(pathName) 
                                    ? System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) 
                                    : System.IO.Path.GetDirectoryName(pathName);
                                string dwgName = string.IsNullOrEmpty(pathName) ? "Drawing" : System.IO.Path.GetFileNameWithoutExtension(pathName);
                                string outputFile = System.IO.Path.Combine(dwgDir, $"{dwgName}_MultiSheet.pdf");

                                // Create a generic master settings configuration to instantiate the overall plotting document spooler
                                PlotSettingsValidator validator = PlotSettingsValidator.Current;
                                PlotSettings masterPs = new PlotSettings(currentLayout.ModelType);
                                masterPs.CopyFrom(currentLayout);
                                try { validator.SetPlotConfigurationName(masterPs, "DWG To PDF.pc3", masterPs.CanonicalMediaName); }
                                catch { validator.SetPlotConfigurationName(masterPs, "DWG To PDF.pc3", null); }
                                
                                PlotInfo masterInfo = new PlotInfo();
                                masterInfo.Layout = currentLayout.ObjectId;
                                masterInfo.OverrideSettings = masterPs;
                                PlotInfoValidator masterValidator = new PlotInfoValidator();
                                masterValidator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                masterValidator.Validate(masterInfo);

                                bool isPlotToFile = masterInfo.ValidatedConfig.IsPlotToFile;

                                // Instantiate the SINGLE root document wrapper (This binds all subsequent pages natively together)
                                engine.BeginDocument(masterInfo, doc.Name, null, 1, isPlotToFile, isPlotToFile ? outputFile : null);

                                for (int i = 0; i < sheetExtents.Count; i++)
                                {
                                    Extents3d ext = sheetExtents[i];

                                    // Build specifically targeted override parameters for this current page
                                    PlotInfo pageInfoRef = new PlotInfo();
                                    pageInfoRef.Layout = currentLayout.ObjectId;

                                    PlotSettings ps = new PlotSettings(currentLayout.ModelType);
                                    ps.CopyFrom(currentLayout);

                                    try { validator.SetPlotConfigurationName(ps, "DWG To PDF.pc3", ps.CanonicalMediaName); }
                                    catch { validator.SetPlotConfigurationName(ps, "DWG To PDF.pc3", null); }

                                    validator.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                                    
                                    // Apply mathematically mirrored margins exactly buffered symmetrically at 0.25 offset uniformly
                                    double margin = 0.25;
                                    Extents2d window = new Extents2d(ext.MinPoint.X - margin, ext.MinPoint.Y - margin, ext.MaxPoint.X + margin, ext.MaxPoint.Y + margin);
                                    validator.SetPlotWindowArea(ps, window);
                                    
                                    validator.SetUseStandardScale(ps, true);
                                    validator.SetStdScaleType(ps, StdScaleType.ScaleToFit); 
                                    validator.SetPlotCentered(ps, true);

                                    pageInfoRef.OverrideSettings = ps;
                                    PlotInfoValidator piValidator = new PlotInfoValidator();
                                    piValidator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                    piValidator.Validate(pageInfoRef);

                                    // Render this specific iteration block inside the master document container
                                    progressDlg.OnBeginSheet();
                                    progressDlg.LowerSheetProgressRange = 0;
                                    progressDlg.UpperSheetProgressRange = 100;
                                    progressDlg.SheetProgressPos = 0;

                                    PlotPageInfo pageInfoData = new PlotPageInfo();
                                    
                                    // CRITICAL FIX: Only pass true for the bLastPage argument on the absolute final sheet!
                                    bool isLastPage = (i == sheetExtents.Count - 1);
                                    engine.BeginPage(pageInfoData, pageInfoRef, isLastPage, null);
                                    
                                    engine.BeginGenerateGraphics(null);
                                    engine.EndGenerateGraphics(null);
                                    engine.EndPage(null);

                                    progressDlg.SheetProgressPos = 100;
                                    progressDlg.OnEndSheet();
                                    progressDlg.PlotProgressPos = i + 1;
                                    plottedCount++;
                                }

                                // Finalize the multi-page aggregation document stream definitively
                                engine.EndDocument(null);
                                engine.EndPlot(null);
                                progressDlg.OnEndPlot();
                            }
                        }
                        
                        // Restore bg plot setting explicitly
                        Application.SetSystemVariable("BACKGROUNDPLOT", bgPlotStr);

                        ed.WriteMessage($"\nSuccess: {plottedCount} sheets plotted.");
                    }
                    else
                    {
                        ed.WriteMessage("\nError: A plot process is already active. Please wait or cancel it.");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during plotting: {ex.Message}");
            }
        }
    }
}
