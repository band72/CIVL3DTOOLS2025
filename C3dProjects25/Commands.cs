using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.Tools.Commands))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.Tools.LocalizeCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.Tools.DescKeyScaleCommands))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RcsTools.ExportDescriptionKeysCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(C3dProjects25.RcsImportPointStylesNet8))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RcsApplyTemplate))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.Tools.CaptureMeasOsmCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.Tools.ConvertCogoCodesCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RcsDeleteAllPointStylesNet8))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RCS_DescKeySet_ImpExp))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RcsDiagPointStyleSetters))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RcsPointStyleExportDeleteImport))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RCS_PointStyle_BlockMarkerMatch))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RCS_PointStyleLayerPolicy))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RCS_PointStyle_ByLayerComponents))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.RCS_PointStyle_DescKeyLayerApply))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2025.Tools.PrintLayoutsCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.QA.RcsQaMaster))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(Civil3D_SymbolTable.Commands))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(Civil3D_SymbolTable.RCS_SymbolTableRobust))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(C3dProjects25.Tables.RcsTableCommands))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.C3D2026.RCS_Blocks_ToLayer0))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.CustomLeader.AutoCAD.Commands.CreateArcLeaderCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.CustomLeader.AutoCAD.Commands.CreateArcLeaderV2Command))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(RCS.CustomLeader.AutoCAD.Commands.SetArcLeaderTextSizeCommand))]
[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(C3dProjects25.RcsDrawChecklistCommand))]

namespace RCS.C3D2025.Tools
{
    public class Commands : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                    "\nRCS_C3D2025_Tools loaded. Type RCS_HELP_RCS for a full command list.");
                
                // Initialize Ribbon
                RibbonUI.InitializeRibbon();
            }
            catch { }
        }

        public void Terminate() { }

        [CommandMethod("RCS_TEST")]
        public void RCS_TEST()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Editor.WriteMessage("\nRCS_C3D2025_Tools is working.");
        }

        [CommandMethod("RCS_HELP_RCS")]
        public void RCS_HELP_RCS()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            ed.WriteMessage(
                "\nRCS Tools — Command Reference" +
                "\n\n--- QA Tools ---" +
                "\n  RCS_QA_RUN              Run QA validation pass (tags failures red)" +
                "\n  RCS_QA_TAGGER           Manually tag objects with a QA type" +
                "\n  RCS_QA_AUTOTAG          Auto-tag all Text/MText/Images in current space" +
                "\n  RCS_QA_AUTOTYPE         Auto-infer QA types from text content" +
                "\n  RCS_QA_FIX_DUPLICATES   Reassign duplicate XData IDs" +
                "\n\n--- Point Styles ---" +
                "\n  RCS_EXPORT_POINTSTYLES_CSV               Export point styles to CSV" +
                "\n  RCS_IMPORT_POINTSTYLES_V4                Import point styles from CSV" +
                "\n  RCS_DELETE_POINTSTYLES_FROM_CSV          Delete styles listed in CSV" +
                "\n  RCS_DELETE_ALL_POINTSTYLES               Delete ALL point styles" +
                "\n  RCS_FORCE_POINTSTYLE_ALL_VIEWS_BYLAYER   Set all views to ByLayer" +
                "\n  RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES  Apply DescKey layers to styles" +
                "\n\n--- Description Keys ---" +
                "\n  RCS_EXPORT_DESCKEY_CODE_BLOCKS  Export DescKey codes + block names" +
                "\n  RCS_IMPORT_DESC_KEYSETSV2       Import DescKey sets" +
                "\n  RCS_FIX_DESC_KEY_SCALE          Fix DescKey scale factors" +
                "\n\n--- Tables & Symbols ---" +
                "\n  RCS_CreateSymbolTableRobust           Build a symbol table" +
                "\n  RCS_TABLES_FROM_WINDOW                Build Line/Curve tables from window" +
                "\n  RCS_BUILD_CURVE_TABLE                 Build a Curve table from MText labels" +
                "\n  RCS_MATCH_POINTSTYLE_BLOCK_MARKERS_NET Match point style block markers" +
                "\n\n--- Drafting Utilities ---" +
                "\n  RCS_SET_ALL_BLOCKS_TO_LAYER0  Move all blocks to Layer 0" +
                "\n  RCS_CONVERT_COGO_CODES        Batch convert COGO raw descriptions to master codes" +
                "\n  RCS_CAPTURE_MEAS_OSM          Export COGO point OSM field capture report (tab-delimited .txt)" +
                "\n  RCS_APPLY_TEMPLATE            Apply a drawing template" +
                "\n  RCS_ARCLEADER                 Create an arc leader annotation" +
                "\n  RCS_ARCLEADER_V2              Create an arc leader V2 (tail arc + trimmed to text box)" +
                "\n  RCS_ARCLEADER_TEXTSIZE        Set the arc leader text size" +
                "\n  RCS_PRINT_MULTI_SHEETS        Batch print layouts to a single multi-page PDF");
        }
        [CommandMethod("RCS_SET_FL83EF")]
        public void SetFlCoordinateSystem()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            try
            {
                // Step 1: Set via Civil 3D Settings API
                try
                {
                    var civDoc = CivilApplication.ActiveDocument;
                    if (civDoc != null)
                    {
                        var uzs = civDoc.Settings.DrawingSettings.UnitZoneSettings;
                        ed.WriteMessage($"\n[RCS] Current CRS: '{uzs.CoordinateSystemCode}'");
                        uzs.CoordinateSystemCode = "FL83-EF";
                        ed.WriteMessage($"\n[RCS] CRS set to: '{uzs.CoordinateSystemCode}'");
                    }
                    else { ed.WriteMessage("\n[RCS] WARN: CivilDocument null."); }
                }
                catch (System.Exception ex) { ed.WriteMessage($"\n[RCS] WARN Civil API: {ex.Message}"); }

                // Activate aerial imagery — send keyword on its own line, no trailing space
                doc.SendStringToExecute("GEOMAP\nAerial\n", true, false, true);

                ed.WriteMessage("\n[RCS] FL83-EF applied. Aerial imagery queued.");
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[RCS ERROR] {ex.Message}"); }
        }

        [CommandMethod("RCS_DIAG_CRS")]
        public void DiagnoseCrs()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            try
            {
                var civDoc = CivilApplication.ActiveDocument;
                if (civDoc == null) { ed.WriteMessage("\n[DIAG] CivilDocument = NULL"); return; }
                var uzs = civDoc.Settings.DrawingSettings.UnitZoneSettings;
                ed.WriteMessage($"\n[DIAG] CoordinateSystemCode = '{uzs.CoordinateSystemCode}'");
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n[DIAG] Exception: {ex.Message}"); }
        }
        [CommandMethod("RCS_HELP")]
        public void ShowHelp()
        {
            try
            {
                var win = new HelpWindow();
                Application.ShowModalWindow(win);
            }
            catch (System.Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage($"\nFailed to launch Help window: {ex.Message}");
            }
        }
    }
}
