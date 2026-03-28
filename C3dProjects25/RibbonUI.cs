using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.Windows;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Globalization;
using System.Windows;

namespace RCS.C3D2025.Tools
{
    public class RibbonUI 
    {
        private static bool _isLoaded = false;

        public static void InitializeRibbon()
        {
            if (Autodesk.Windows.ComponentManager.Ribbon == null)
            {
                Autodesk.Windows.ComponentManager.ItemInitialized += ComponentManager_ItemInitialized;
            }
            else
            {
                CreateRibbon();
            }
        }

        private static void ComponentManager_ItemInitialized(object sender, RibbonItemEventArgs e)
        {
            if (Autodesk.Windows.ComponentManager.Ribbon != null)
            {
                CreateRibbon();
                Autodesk.Windows.ComponentManager.ItemInitialized -= ComponentManager_ItemInitialized;
            }
        }

        private static void CreateRibbon()
        {
            try
            {
                if (_isLoaded) return;

                RibbonControl ribbon = ComponentManager.Ribbon;
                if (ribbon != null)
                {
                    RibbonTab rcsTab = ribbon.FindTab("RCS_TOOLS");
                    if (rcsTab == null)
                    {
                        rcsTab = new RibbonTab();
                        rcsTab.Title = "RCS Tools";
                        rcsTab.Id = "RCS_TOOLS";
                        ribbon.Tabs.Add(rcsTab);
                    }

                    int btnIdx = 1;

                    // --- PANEL: QA Tools ---
                    RibbonPanel qaPanel = CreatePanel(rcsTab, "QA Tools");
                    AddRibbonButton(qaPanel.Source, "Run QA", "RCS_QA_RUN ", btnIdx++);
                    AddRibbonButton(qaPanel.Source, "Tagger", "RCS_QA_TAGGER ", btnIdx++);
                    AddRibbonButton(qaPanel.Source, "Auto Tag", "RCS_QA_AUTOTAG ", btnIdx++);
                    AddRibbonButton(qaPanel.Source, "Auto Type", "RCS_QA_AUTOTYPE ", btnIdx++);
                    AddRibbonButton(qaPanel.Source, "Fix Duplicates", "RCS_QA_FIX_DUPLICATES ", btnIdx++);

                    // --- PANEL: Point Styles ---
                    RibbonPanel stylesPanel = CreatePanel(rcsTab, "Point Styles");
                    AddRibbonButton(stylesPanel.Source, "Export Styles", "RCS_EXPORT_POINTSTYLES_CSV ", btnIdx++);
                    AddRibbonButton(stylesPanel.Source, "Import Styles", "RCS_IMPORT_POINTSTYLES_V4 ", btnIdx++);
                    AddRibbonButton(stylesPanel.Source, "Delete Styles", "RCS_DELETE_POINTSTYLES_FROM_CSV ", btnIdx++);
                    AddRibbonButton(stylesPanel.Source, "Delete All Styles", "RCS_DELETE_ALL_POINTSTYLES ", btnIdx++);
                    AddRibbonButton(stylesPanel.Source, "Force ByLayer", "RCS_FORCE_POINTSTYLE_ALL_VIEWS_BYLAYER ", btnIdx++);
                    AddRibbonButton(stylesPanel.Source, "Apply Desc Layers", "RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES ", btnIdx++);

                    // --- PANEL: Desc Keys ---
                    RibbonPanel descPanel = CreatePanel(rcsTab, "Desc Keys");
                    AddRibbonButton(descPanel.Source, "Export DescKeys", "RCS_EXPORT_DESCKEY_CODE_BLOCKS ", btnIdx++);
                    AddRibbonButton(descPanel.Source, "Import DescKeys", "RCS_IMPORT_DESC_KEYSETSV2 ", btnIdx++);
                    AddRibbonButton(descPanel.Source, "Fix DescKey Scale", "RCS_FIX_DESC_KEY_SCALE ", btnIdx++);

                    // --- PANEL: Tables & Symbols ---
                    RibbonPanel tablePanel = CreatePanel(rcsTab, "Tables & Symbols");
                    AddRibbonButton(tablePanel.Source, "Symbol Table", "RCS_CreateSymbolTableRobust ", btnIdx++);
                    AddRibbonButton(tablePanel.Source, "Tables Window", "RCS_TABLES_FROM_WINDOW ", btnIdx++);
                    AddRibbonButton(tablePanel.Source, "Curve Table", "RCS_BUILD_CURVE_TABLE ", btnIdx++);
                    AddRibbonButton(tablePanel.Source, "Match Markers", "RCS_MATCH_POINTSTYLE_BLOCK_MARKERS_NET ", btnIdx++);

                    // --- PANEL: Drafting ---
                    RibbonPanel draftPanel = CreatePanel(rcsTab, "Drafting Utilities");
                    AddRibbonButton(draftPanel.Source, "Blocks To Layer0", "RCS_SET_ALL_BLOCKS_TO_LAYER0 ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "Convert Cogo", "RCS_CONVERT_COGO_CODES ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "Apply Template", "RCS_APPLY_TEMPLATE ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "Capture OSM", "RCS_CAPTURE_MEAS_OSM ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "Create ArcLeader", "RCS_ARCLEADER ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "ArcLeader V2", "RCS_ARCLEADER_V2 ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "Set TextSize", "RCS_ARCLEADER_TEXTSIZE ", btnIdx++);
                    AddRibbonButton(draftPanel.Source, "Print Sheets", "RCS_PRINT_MULTI_SHEETS ", btnIdx++);

                    _isLoaded = true;
                }
            }
            catch (System.Exception ex)
            {
                var ed = Application.DocumentManager.MdiActiveDocument?.Editor;
                if (ed != null) ed.WriteMessage($"\nError loading Ribbon: {ex.Message}");
            }
        }

        private static RibbonPanel CreatePanel(RibbonTab tab, string title)
        {
            RibbonPanelSource panelSrc = new RibbonPanelSource();
            panelSrc.Title = title;
            RibbonPanel panel = new RibbonPanel();
            panel.Source = panelSrc;
            tab.Panels.Add(panel);
            return panel;
        }

        private static ImageSource CreateTextBitmap(string text, int size)
        {
            var visual = new DrawingVisual();
            using (var drawingContext = visual.RenderOpen())
            {
                var brush = new LinearGradientBrush(
                    Colors.DeepSkyBlue,
                    Colors.Navy,
                    new System.Windows.Point(0, 0),
                    new System.Windows.Point(1, 1));
                    
                drawingContext.DrawRoundedRectangle(
                    brush,
                    new Pen(Brushes.White, 1),
                    new Rect(0, 0, size, size),
                    4, 4);

                var formattedText = new FormattedText(
                    text,
                    CultureInfo.InvariantCulture,
                    FlowDirection.LeftToRight,
                    new Typeface(new FontFamily("Arial"), FontStyles.Normal, FontWeights.Bold, FontStretches.Normal),
                    size * 0.45,
                    Brushes.White,
                    VisualTreeHelper.GetDpi(visual).PixelsPerDip);

                double x = (size - formattedText.Width) / 2;
                double y = (size - formattedText.Height) / 2;
                drawingContext.DrawText(formattedText, new System.Windows.Point(x, y));
            }

            var bitmap = new RenderTargetBitmap(
                size, size, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(visual);
            return bitmap;
        }

        private static void AddRibbonButton(RibbonPanelSource panelSrc, string text, string commandName, int number)
        {
            RibbonButton btn = new RibbonButton();
            // Automatically break at first space for Ribbon stacking
            int breakIndex = text.IndexOf(" ");
            if (breakIndex > 0)
                btn.Text = text.Substring(0, breakIndex) + "\n" + text.Substring(breakIndex + 1);
            else
                btn.Text = text;
            
            btn.ShowText = true;
            btn.ShowImage = true; 
            
            // Assign generated bitmaps for the button numbers
            btn.Image = CreateTextBitmap(number.ToString(), 16);
            btn.LargeImage = CreateTextBitmap(number.ToString(), 32);

            btn.CommandParameter = commandName;
            btn.CommandHandler = new RibbonCommandHandler();
            btn.Size = RibbonItemSize.Large;
            btn.Orientation = System.Windows.Controls.Orientation.Vertical;

            panelSrc.Items.Add(btn);
        }
    }

    public class RibbonCommandHandler : System.Windows.Input.ICommand
    {
        public event EventHandler CanExecuteChanged
        {
            add { }
            remove { }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            // parameter is the value of CommandParameter — the command string directly.
            // (It is NOT the RibbonButton itself — that was the previous bug.)
            string cmd = parameter as string;
            if (!string.IsNullOrEmpty(cmd))
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                doc?.SendStringToExecute(cmd, true, false, false);
            }
        }
    }
}
