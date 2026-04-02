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
                { "FIP", ("FIP 5/8 \\P (NO-ID)","FIP 5/8 \\P (NO-ID)") },
                { "FIR", ("FIR 5/8 \\P (NO-ID)","FIR 5/8 \\P (NO-ID)") },
                { "SIRC", ( "SIR 1/2 \\P (LB8484) ", "SIR 1/2 \\P (LB8484) ") },
                { "NDF", ("NAIL&DISK\\P FOUND","NAIL&DISK \\P FND") },
                { "SIP", ( "SIP 1/2 \\P (LB8484) ", "SIP 1/2 \\P (LB8484) ") },
                { "IPS", ( "SIP 1/2 \\P (LB8484) " , "SIP 1/2 \\P IP (LB8484) ") },
                { "CMF", ("CMF 4X4 (NO-ID)", "CMF 4X4 \\P (NO-ID)") },
                { "IRS*", ( "SIR 1/2 \\P (LB8484) ", "SIR 1/2 \\P IR (LB8484) ") },
                { "IRF*", ( "FIR 1/2 \\P (NO-ID) ", "FIR 1/2 \\P IR (NO-ID) ") },
                { "IPS*", ( "SIP 1/2 \\P (LB8484) ", "SIP 1/2 \\P IP (LB8484) ") },
                { "IPF*", ( "FIP 1/2 \\P (NO-ID) ", "FIP 1/2 \\P IP (NO-ID) ") },
                { "FIP5/8IRNOID", ("FIR 5/8 \\P (NO-ID)","FIR 5/8 \\P (NO-ID)") },
                { "FIR1/2IRNOID", ("FIP 1/2 \\P (NO-ID)","FIR 1/2 \\P (NO-ID)") },
                { "FIR5/8IPNOID", ("FIR 5/8 \\P (NO-ID)","FIR 5/8 \\P (NO-ID)") },
                { "FIP1/2IPNOID", ("FIP 1/2 \\P (NO-ID)","FIP 1/2 \\P (NO-ID)") },
                { "FOUND NAIL & DISK*", ("NAIL&DISK\\P FOUND","NAIL&DISK \\P FND") },
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

        // Pre-cleaned version of _codeMap keys (hyphens, quotes stripped) built once at class load.
        // Avoids redundant string allocations inside the nested per-point loops.
        private static readonly Dictionary<string, (string MasterCode, string MasterDesc)> _cleanedCodeMap =
            BuildCleanedCodeMap();

        private static Dictionary<string, (string MasterCode, string MasterDesc)> BuildCleanedCodeMap()
        {
            var result = new Dictionary<string, (string MasterCode, string MasterDesc)>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in _codeMap)
            {
                string cleanKey = kvp.Key.Replace("-", "").Replace("'", "").Replace("\"", "");
                // Last writer wins for any accidental duplicates after cleaning
                result[cleanKey] = kvp.Value;
            }
            return result;
        }

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
                else if (psr.Status == PromptStatus.Error)
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
                                // First step in cleaning: remove hyphens, single quotes, and double quotes
                                string cleanedRaw = raw.Replace("-", "").Replace("'", "").Replace("\"", "");

                                // Track whether the point has already been upgraded to ForWrite
                                // to avoid a double UpgradeOpen() when cleaning AND a map match both fire.
                                bool isUpgraded = false;

                                // Unconditionally save the cleaned string back to the point immediately
                                if (raw != cleanedRaw)
                                {
                                    point.UpgradeOpen();
                                    isUpgraded = true;
                                    point.RawDescription = cleanedRaw;
                                }

                                string searchKey = cleanedRaw.Trim();

                                bool matched = false;
                                string remainder = "";
                                (string MasterCode, string MasterDesc) mapping = (string.Empty, string.Empty);

                                // 1. Try Exact match first using pre-cleaned key map (no per-iteration allocations)
                                foreach (var kvp in _cleanedCodeMap)
                                {
                                    if (!kvp.Key.EndsWith("*") && string.Equals(searchKey, kvp.Key, StringComparison.OrdinalIgnoreCase))
                                    {
                                        matched = true;
                                        mapping = kvp.Value;
                                        remainder = ""; // Exact matches leave no remainder
                                        break;
                                    }
                                }

                                // 2. Try Wildcard match fallback (key ends with *)
                                if (!matched)
                                {
                                    foreach (var kvp in _cleanedCodeMap)
                                    {
                                        if (kvp.Key.EndsWith("*"))
                                        {
                                            string prefix = kvp.Key.Substring(0, kvp.Key.Length - 1); // remove the *
                                            if (searchKey.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                                            {
                                                matched = true;
                                                mapping = kvp.Value;
                                                remainder = searchKey.Substring(prefix.Length);
                                                break;
                                            }
                                        }
                                    }
                                }

                                if (matched)
                                {
                                    // Only upgrade to Write mode if not already done during the cleaning step
                                    if (!isUpgraded)
                                    {
                                        point.UpgradeOpen();
                                        isUpgraded = true;
                                    }

                                    string newCode = mapping.MasterCode;
                                    string newDesc = mapping.MasterDesc;

                                    // Dynamic inline regex formatter to correctly space fractions and inject \P line breaks for IDs
                                    Func<string, string> formatSuffixes = (text) => 
                                    {
                                        if (string.IsNullOrWhiteSpace(text)) return text;
                                        
                                        // Ensure space between Letters and Fractions (e.g. FIP1/2 -> FIP 1/2)
                                        text = System.Text.RegularExpressions.Regex.Replace(text, @"([A-Za-z])(\d+/\d+)", "$1 $2");
                                        
                                        // Sub-format NOID suffix (e.g. 1/2NOID -> 1/2 \P NO-ID)
                                        text = System.Text.RegularExpressions.Regex.Replace(text, @"(\d+/\d+)\s*(NOID)", "$1 \\P NO-ID", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                        
                                        // Sub-format LB suffix (e.g. 1/2LB1234 -> 1/2 \P LB-1234)
                                        text = System.Text.RegularExpressions.Regex.Replace(text, @"(\d+/\d+)\s*(LB)\s*(\d+)", "$1 \\P $2-$3", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                        
                                        return text.Replace("  ", " ").Trim();
                                    };

                                    if (!string.IsNullOrWhiteSpace(newCode))
                                    {
                                        point.RawDescription = formatSuffixes(newCode + remainder);
                                    }

                                    if (!string.IsNullOrWhiteSpace(newDesc))
                                    {
                                        // Update DescriptionFormat for the FullDescription
                                        point.DescriptionFormat = formatSuffixes(newDesc + remainder);
                                    }

                                    cntHit++;
                                }
                                else
                                {
                                    // Diagnostic: show exact bytes so invisible chars are visible
                                    byte[] _diagBytes = System.Text.Encoding.UTF8.GetBytes(searchKey);
                                    string hexDump = string.Join(" ", Array.ConvertAll(_diagBytes, b => b.ToString("X2")));
                                    ed.WriteMessage($"\n  [NO-MATCH] pt#{point.PointNumber} raw=\"{searchKey}\" len={searchKey.Length} hex=[{hexDump}]");
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
