using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025.Tools
{
    public class ConvertCogoCodesCommand
    {
        // Embedded lookup table: ORIGINAL -> (MASTER_CODE, MASTER_DESCRIPTION)
        private static readonly Dictionary<string, (string MasterCode, string MasterDesc)> _codeMap =
            new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase)
            {
                { "CMF", ("CMF 4X4 (NO-ID)", "CMF 4X4 (NO-ID)") },
                { "15RCP", ("RCP", "Reinforced Concrete Pipe") },
                { "18CMP", ("CMP", "Corrugated Metal Pipe") },
                { "18RCP", ("RCP", "Reinforced Concrete Pipe") },
                { "24CMP", ("CMP", "Corrugated Metal Pipe") },
                { "24RCP", ("RCP", "Reinforced Concrete Pipe") },
                { "2ND", ("BC Building Corner", "") },
                { "36CMP", ("CMP", "Corrugated Metal Pipe") },
                { "36RCP", ("RCP", "Reinforced Concrete Pipe") },
                { "4'CLF", ("F0 Fence line", "") },
                { "4'MF", ("F0 Fence line", "") },
                { "4'VF", ("F0 Fence line", "") },
                { "4'WF", ("F0 Fence line", "") },
                { "48CMP", ("CMP", "Corrugated Metal Pipe") },
                { "48RCP", ("RCP", "Reinforced Concrete Pipe") },
                { "6'CLF", ("F0 Fence line", "") },
                { "6'MF", ("F0 Fence line", "") },
                { "6'VF", ("F0 Fence line", "") },
                { "6'WF", ("F0 Fence line", "") },
                { "8'CLF", ("F0 Fence line", "") },
                { "8'MF", ("F0 Fence line", "") },
                { "8'VF", ("F0 Fence line", "") },
                { "8'WF", ("F0 Fence line", "") },
                { "ADJ", ("BC Building Corner", "") },
                { "BFP", ("BFP", "Back Flow Preventer") },
                { "BLD", ("BC Building Corner", "") },
                { "BOC", ("BOC1", "Back of Curb end") },
                { "BRICK", ("BHW", "Brick Headwall") },
                { "CALC", ("COGO", "Calculated Point") },
                { "CB", ("CCB", "Corner Catch Basin") },
                { "CLDITCH", ("FLD", "FlowLine Ditch") },
                { "CLDR", ("CLD", "CenterLine Dirt") },
                { "CLGRAVEL", ("TCL", "Topo. Clay") },
                { "CLRD", ("ES Edge Stone", "") },
                { "CO", ("CO Clean Out", "") },
                { "COLL", ("CBC", "Center Brick Column") },
                { "CONC", ("TC Topo. Concrete", "") },
                { "CONC BH", ("CHW", "Concrete Headwall") },
                { "CONC S/W", ("EC Edge Concrete", "") },
                { "COV", ("BE Building Edge", "") },
                { "CP", ("DT Disk Traverse", "") },
                { "CPP", ("CPP", "Concrete Power Pole") },
                { "DF", ("DRFLD", "Drainfield") },
                { "DMH", ("DMH", "Drainage ManHole") },
                { "EOW", ("EOW", "Edge Of Water") },
                { "EP", ("CA Corner Asphalt", "") },
                { "ET", ("ET Electric Transformer", "") },
                { "EUB", ("ET Electric Transformer", "") },
                { "FCM 4/4", ("WTP", "Wood Telephone Pole") },
                { "FH", ("FH Fire Hydrant", "") },
                { "FIP 1/2 CAP/ILL", ("IPF", "Iron Pipe Found") },
                { "FIP 1/2 LB", ("IPF", "Iron Pipe Found") },
                { "FIP 1/2 NO/ID", ("IPF", "Iron Pipe Found") },
                { "FIP 3/4 NO/ID", ("IPF", "Iron Pipe Found") },
                { "FIR 1 NO/ID", ("IRF", "Iron Rod Found") },
                { "FIR 1/2 CAP/ILL", ("IRF", "Iron Rod Found") },
                { "FIR 1/2 LB", ("IRF", "Iron Rod Found") },
                { "FIR 1/2 NO/ID", ("IRF", "Iron Rod Found") },
                { "FIR 5/8 CAP/ILL", ("IRF", "Iron Rod Found") },
                { "FIR 5/8 NO/ID", ("IRF", "Iron Rod Found") },
                { "FND", ("IRF", "Iron Rod Found") },
                { "FND LB", ("IRF", "Iron Rod Found") },
                { "FOC", ("FOC0", "FACE OF CURB") },
                { "G", ("FLG", "FlowLine Gutter") },
                { "GAR", ("GRG02", "Garage Line") },
                { "GMV", ("GVV", "Gas Valve Vault") },
                { "GRV", ("TG Topo. Gravel", "") },
                { "GVV", ("GV Gas Valve", "") },
                { "HC", ("BOC0", "Back of Curb line") },
                { "HCS", ("HCP", "Handicap Parking") },
                { "LP", ("CLP", "Concrete Light Pole") },
                { "MES", ("MES1", "Mitered End Section") },
                { "MH", ("DMH", "Drainage ManHole") },
                { "MHD", ("DMH", "Drainage ManHole") },
                { "MHW", ("MHW", "Masonry Headwall") },
                { "NG", ("NG", "Natural Ground") },
                { "OHW", ("OH", "Overhead") },
                { "PAV", ("PAV", "Edge of paver") },
                { "PED", ("PED", "Pedestrian Crossing") },
                { "POOL", ("SPE", "Swimming Pool Edge") },
                { "PS", ("PS", "PAINT STRIPE") },
                { "SBS", ("SBS", "SPEED BUMP SIGN") },
                { "SCB", ("SCB", "Sprinkler Control Box") },
                { "SCRN", ("SCRN", "Edge of screen") },
                { "SEP", ("SPTB", "Septic structure bottom") },
                { "SEP LID", ("SPTT", "Septic structure top") },
                { "SHD", ("SHD", "Edge of Shed") },
                { "SIR 1/2 RCS LB8484", ("SIRC 1/2 \\P RCS LB8484", "SIRC 1/2 RCS LB8484") },
                { "SIRC 1/2 RCS LB8484", ("SIRC 1/2 \\P RCS LB8484", "SIRC 1/2 RCS LB8484") },
                { "SIRC", ("SIR 1/2 \\P  LB8484", "SIR 1/2 RCS LB8484") },
                { "SIR", ("SIR 1/2 \\P LB8484", "SIR 1/2 RCS LB8484") },
                { "SITE BM", ("BM Bench Mark", "") },
                { "SLS", ("SLS", "Speed Limit Sign") },
                { "SMH", ("SMH", "Sanitary ManHole") },
                { "SND RCS LB8484", ("DS Disk Set", "") },
                { "SS", ("STP", "Stop Sign") },
                { "STEPS", ("STEPS", "Steps") },
                { "TC", ("MPSP", "Metal Pedestrian Signal Pole") },
                { "TIDAL", ("EOW", "Edge Of Water") },
                { "TILE", ("TILE", "Edge of Tile") },
                { "TOB", ("TOB", "Top Of Bank") },
                { "TP", ("DT Disk Traverse", "") },
                { "TSB", ("MTSP", "Metal Traffic Signal Pole") },
                { "TSCB", ("TSCB", "Traffic Signal Control Box") },
                { "UR", ("UR", "Utility Riser") },
                { "WELL", ("WELL", "Water Well") },
                { "WM", ("WM Water Meter", "") },
                { "WOOD BH", ("EW", "Edge of Wood") },
                { "WOOD DECK", ("WDCK", "Wood Deck") },
                { "WOOD DOCK", ("WDOK", "Wood Dock") },
                { "WPP", ("WPP", "Wood Power Pole") },
                { "WV", ("WV Water Valve", "") },
                { "X", ("XCF", "X Cut Found") },
                { "XCUT", ("XCF", "X Cut Found") },
                { "IPF-1/2-NO-ID", ("FIP 1/2\\P (NO-ID)", "FIP 1/2\\P (NO-ID)") },
                { "IPF-5/8-NO-ID", ("FIP 5/8\\P (NO-ID)", "FIP 5/8\\P (NO-ID)") }
            };

        [CommandMethod("RCS_CONVERT_COGO_CODES")]
        public void RunConvertCogoCodes()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            int cntAll = 0;
            int cntHit = 0;
            int cntSkip = 0;

            try
            {
                // 1. Select COGO points (or press Enter to select ALL)
                PromptSelectionOptions pso = new PromptSelectionOptions
                {
                    MessageForAdding = "\nSelect COGO points to convert (Enter = ALL): ",
                    AllowDuplicates = false
                };
                
                TypedValue[] filterValues = { new TypedValue((int)DxfCode.Start, "AECC_COGO_POINT") };
                SelectionFilter filter = new SelectionFilter(filterValues);

                PromptSelectionResult psr = ed.GetSelection(pso, filter);

                SelectionSet ss;
                if (psr.Status == PromptStatus.OK)
                {
                    ss = psr.Value;
                }
                else if (psr.Status == PromptStatus.None || psr.Status == PromptStatus.Error)
                {
                    // User pressed Enter without selection -> Process ALL COGO points
                    PromptSelectionResult allPsr = ed.SelectAll(filter);
                    if (allPsr.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nNo COGO points found.");
                        return;
                    }
                    ss = allPsr.Value;
                }
                else
                {
                    // Cancelled or other
                    return;
                }

                if (ss == null || ss.Count == 0)
                {
                    ed.WriteMessage("\nNo COGO points selected.");
                    return;
                }

                // 2. Iterate and update each point
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in ss)
                    {
                        if (selObj != null)
                        {
                            var point = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as CogoPoint;
                            if (point == null) continue;

                            cntAll++;
                            string raw = point.RawDescription;
                            if (raw != null)
                            {
                                string key = raw.Trim();
                                if (_codeMap.TryGetValue(key, out var mapping))
                                {
                                    // Upgrade to Write mode
                                    point.UpgradeOpen();

                                    string newCode = mapping.MasterCode;
                                    string newDesc = mapping.MasterDesc;

                                    if (!string.IsNullOrWhiteSpace(newCode))
                                    {
                                        point.RawDescription = newCode;
                                    }

                                    if (!string.IsNullOrWhiteSpace(newDesc))
                                    {
                                        // Update DescriptionFormat for the FullDescription
                                        point.DescriptionFormat = newDesc;
                                    }

                                    cntHit++;
                                }
                                else
                                {
                                    cntSkip++;
                                }
                            }
                            else
                            {
                                cntSkip++;
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nDone. Processed: {cntAll} | Converted: {cntHit} | No Match: {cntSkip}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError converting COGO codes: {ex.Message}");
            }
        }
    }
}
