using System.Collections.Generic;
using System.Windows;

namespace RCS.C3D2025.Tools
{
    public partial class HelpWindow : Window
    {
        public HelpWindow()
        {
            InitializeComponent();
            LoadCommands();
        }

        private void LoadCommands()
        {
            var list = new List<CommandDetails>
            {
                new CommandDetails(1, "Run QA", "RCS_QA_RUN", "Executes the master Quality Assurance engine. Enumerates all survey figures and CogoPoints, analyzing their intersecting structures and geometries to strictly validate topological rules like self-intersecting boundaries and loose connections."),
                new CommandDetails(2, "Tagger", "RCS_QA_TAGGER", "Spawns a selection interop session allowing you to hand-pick drawing geometries. It then attaches specific compliance XData payload schemas (metadata) natively inside the selected objects' AppData block."),
                new CommandDetails(3, "Auto Tag", "RCS_QA_AUTOTAG", "A bulk operation that iterates the BlockTableRecord across targeted standard linework layers and automatically injects systematic QC compliance XData tags to thousands of lines at once."),
                new CommandDetails(4, "Auto Type", "RCS_QA_AUTOTYPE", "Analyzes the description strings of survey figures (e.g., BLDG, EP, FC). Compares properties against an internal regex matrix and rigidly assigns categorized survey figure line-types natively within the AeccSurvey system."),
                new CommandDetails(5, "Fix Duplicates", "RCS_QA_FIX_DUPLICATES", "Scans the entire CivilDocument for CogoPoints containing identical coordinates (within a mathematically tiny 0.001 positional tolerance) and logically cleans up and physically deletes overlapping duplicate definitions."),
                new CommandDetails(6, "Export Styles", "RCS_EXPORT_POINTSTYLES_CSV", "Taps into `CivilDocument.Styles.PointStyles`. Iterates every single style, extracts the Name, AutoCADLayer, Maker Block Name, and Label constraints, and systematically dumps the output stream into an external auditing CSV utilizing CsvHelper."),
                new CommandDetails(7, "Import Styles", "RCS_IMPORT_POINTSTYLES_V4", "Reads the standard CSV mapping sequence. Iterates the active drawing's Point Styles, clones missing styles, and rigidly assigns the DBObject definitions (Layer bindings, Block linkages, scaling parameters) to match the CSV architecture exactly."),
                new CommandDetails(8, "Delete Styles", "RCS_DELETE_POINTSTYLES_FROM_CSV", "Consumes a targeted CSV file map, explicitly identifies listed Point Styles via internal ObjectIDs, and aggressively erases them from the drawing database whilst safely intercepting 'eWasErased' failures."),
                new CommandDetails(9, "Delete All", "RCS_DELETE_ALL_POINTSTYLES", "Performs an unconditional master loop over the `PointStyles` root collection (strategically bypassing the protected 'Standard' style) and natively purges every custom visual style from geometric memory."),
                new CommandDetails(10, "Force ByLayer", "RCS_FORCE_POINTSTYLE_ALL_VIEWS_BYLAYER", "Dives forcefully deep into the hierarchical display properties of every active Point Style. For every active `DisplayStyle` viewstate (Plan, Model, Profile, Section), it overwrites Maker/Label component color codes dynamically to `Color.FromByLayer()`."),
                new CommandDetails(11, "Apply Desc Layers", "RCS_APPLY_DESCKEY_LAYERS_TO_POINTSTYLES", "Scans every single `PointDescriptionKey` inside the current active DescKeySet. It reads the specific target insertion layer from the Key, accesses the assigned Point Style, and forcibly molds the Point Style's base entity layer sequence to mirror the Key's insertion layer flawlessly."),
                new CommandDetails(12, "Export DescKeys", "RCS_EXPORT_DESCKEY_CODE_BLOCKS", "Sequences through the active Description Key Sets memory banks, formatting every single key's raw Code string, Format, associative PointStyle, and Target Layer schema into a cleanly exportable Excel/CSV data matrix file."),
                new CommandDetails(13, "Import DescKeys", "RCS_IMPORT_DESC_KEYSETSV2", "Consumes flat CSV files directly into code arrays and dynamically mints new native standard Description Key Sets, securely injecting them with raw keys, spawning any missing referenced Layers, and gracefully bridging missing styles."),
                new CommandDetails(14, "Fix Scale", "RCS_FIX_DESC_KEY_SCALE", "Transacts bulk modifications directly to physical Block Point Scale geometries inside all dynamic styles, explicitly forcing their deployment behavior to scale consistently with the active viewport Annotation scale."),
                new CommandDetails(15, "Symbol Table", "RCS_CreateSymbolTableRobust", "Actively queries the database `BlockTable` linked to referenced Point Styles. Iterates the drawing array to programmatically drop an elegant, visual AutoCAD `Table` database object with physical blocks previewed next to descriptive schemas."),
                new CommandDetails(16, "Line/Curve Tbl", "RCS_TABLES_FROM_WINDOW", "Prompts Editor.GetSelection window inputs. The code recursively sweeps through the captured selection set of Polyline / Line entities, parses geometry components, generates serialized tag numbers, and autonomously drops exhaustive Civil 3D Schedule block tables."),
                new CommandDetails(17, "Curve Table", "RCS_BUILD_CURVE_TABLE", "Specialized isolation engine. Code actively extracts pure CircularArc3d geometry out of chaotic survey strings, extracts Radius, Chord, Delta, and Arc Length parameters natively, and synthesizes standard geometric CAD curve tables automatically."),
                new CommandDetails(18, "Match Markers", "RCS_MATCH_POINTSTYLE_BLOCK_MARKERS_NET", "Loops all primitive CogoPoints in the database. Calculates the actively assigned marker block from the Style layout, calculates marker rotation, and natively spawns a detached basic AutoCAD `BlockReference` physically over the point for true 3D visual export pipelines."),
                new CommandDetails(19, "Blks To Layer0", "RCS_SET_ALL_BLOCKS_TO_LAYER0", "Code forcibly unlocks the internal `BlockTableRecord` structure. Safely enumerates into the nested primitive entities stored within your specific blocks and aggressively forces all their raw layer IDs to precisely `Layer 0` to secure strict ByLayer inheritance definitions."),
                new CommandDetails(20, "Convert Cogo", "RCS_CONVERT_COGO_CODES", "Iterates rapidly via transaction processing over all Civil Cogo Points. Performs Regex code parsing sanitization against all Raw Descriptions, physically dropping rogue quotes, hyphens, and whitespace barriers so points map beautifully across Description Keys."),
                new CommandDetails(21, "Apply Template", "RCS_APPLY_TEMPLATE", "Connects via advanced DBX cloning methodologies to stealthily rip styles and templates out of an external `.dwt` master file. It strictly clones layers, standard line styles, Block definitions, and deep Civil 3D style structures natively into your live production drawing."),
                new CommandDetails(22, "Capture OSM", "RCS_CAPTURE_MEAS_OSM", "Internal diagnostic utility snippet. Actively snaps into native CAD environment callbacks, locking onto Object Snap Measure endpoints to aggregate vertex geometry payload arrays for complex boundary checking."),
                new CommandDetails(23, "ArcLeader V1", "RCS_ARCLEADER", "Interactive multi-step placement jig command. Math engine builds an exact CircularArc3d bounding the 3 click nodes, spawns a geometric length Solid Arrowhead, drops an MText box configured with visual mask vectors, and links/groups the database artifacts directly onto the specific active CAD layer."),
                new CommandDetails(24, "ArcLeader V2", "RCS_ARCLEADER_V2", "The V2 variant that clones the main trajectory mathematical arc invisibly in memory. Traces the curve through the arrowhead using safe GetChordAngleAt trigonometry logic computationally intersecting the arc length, and appends a single truncated CAD Arc trace exactly bridging the arrowhead solid and Text box logic block."),
                new CommandDetails(25, "Text Size", "RCS_ARCLEADER_TEXTSIZE", "Invokes a bare command-line Double-precision parameter prompt via the Editor. Captures the numeric size value and persists the scaled assignment centrally against the active `ArcLeaderSettings.Current.TextHeight` class scope."),
                new CommandDetails(26, "Print Sheets", "RCS_PRINT_MULTI_SHEETS", "A heavy batch deployment processing sequence on PaperSpace layouts. Scans bounding limits on standard Viewports, overrides plot definitions to center output uniformly on 8.5x11 PDF schema, and rigidly calibrates 0.25-inch paper margins while synchronously cycling and deploying background print pipelines."),
                new CommandDetails(27, "Set FL CRS", "RCS_SET_FL83EF", "Programmatically targets the `CivilApplication.ActiveDocument.Settings` module, aggressively setting `CoordinateSystemCode = \"FL83-EF\"` to force the database to NAD83 Florida State Planes, East Zone, US Foot. Instantly fires a macro to query Autodesk cloud proxies and natively map the high-resolution Bing Aerial hybrid overlay into model space."),
                new CommandDetails(28, "Help Guide", "RCS_HELP", "You are here. Instantiates a native C# WPF `Window` directly over the main window application frame. Uses a dynamic `ObservableCollection` loop mapping to systematically populate customized code diagnostic profiles securely displayed across an ItemsControl component grid.")
            };

            CmdList.ItemsSource = list;
        }
    }

    public class CommandDetails
    {
        public int Number { get; set; }
        public string Title { get; set; }
        public string Command { get; set; }
        public string Description { get; set; }

        public CommandDetails(int number, string title, string command, string desc)
        {
            Number = number;
            Title = title;
            Command = command;
            Description = desc;
        }
    }
}
