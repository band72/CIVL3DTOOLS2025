using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace C3dProjects25.Tables
{
    public class RcsTableCommands
    {
        [CommandMethod("RCS_TABLES_FROM_WINDOW")]
        public void TablesFromWindow()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            var pr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\nSelect TEXT/MTEXT with L# and C# info: " }, 
                new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT") }));

            if (pr.Status != PromptStatus.OK) return;

            var items = new List<TextItem>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject so in pr.Value)
                {
                    var ent = tr.GetObject(so.ObjectId, OpenMode.ForRead);
                    if (ent is MText mt) items.Add(new TextItem { Pos = mt.Location, Text = mt.Text, Style = mt.TextStyleName });
                    else if (ent is DBText dt) items.Add(new TextItem { Pos = dt.Position, Text = dt.TextString, Style = dt.TextStyleName });
                }

                // Top-to-Bottom, Left-to-Right Sort
                items = items.OrderByDescending(i => i.Pos.Y).ThenBy(i => i.Pos.X).ToList();

                var parser = new TableParser();
                parser.ParseMixed(items.Select(i => i.Text).ToList());

                if (parser.Lines.Count > 0)
                {
                    var ptRes = ed.GetPoint("\nPick insertion point for LINE TABLE: ");
                    if (ptRes.Status == PromptStatus.OK)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        var tbl = BuildLineTable(ptRes.Value, parser.Lines);
                        btr.AppendEntity(tbl);
                        tr.AddNewlyCreatedDBObject(tbl, true);
                    }
                }

                if (parser.Curves.Count > 0)
                {
                    var ptRes = ed.GetPoint("\nPick insertion point for CURVE TABLE: ");
                    if (ptRes.Status == PromptStatus.OK)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        var tbl = BuildCurveTable(ptRes.Value, parser.Curves);
                        btr.AppendEntity(tbl);
                        tr.AddNewlyCreatedDBObject(tbl, true);
                    }
                }

                tr.Commit();
                ed.WriteMessage("\nDone.");
            }
        }

        [CommandMethod("RCS_BUILD_CURVE_TABLE")]
        public void BuildCurveTableCmd()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            var db = doc.Database;

            var pr = ed.GetSelection(new PromptSelectionOptions { MessageForAdding = "\nSelect MText Curve Labels: " }, 
                new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "MTEXT") }));

            if (pr.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var curves = new List<CurveRecV24>();

                foreach (SelectedObject so in pr.Value)
                {
                    if (tr.GetObject(so.ObjectId, OpenMode.ForRead) is MText mt)
                    {
                        var rec = ParseCurveV24(mt.Text);
                        if (rec != null) curves.Add(rec);
                    }
                }

                if (curves.Count > 0)
                {
                    curves = curves.OrderBy(c => ParseIdNum(c.ID)).ToList();
                    var ptRes = ed.GetPoint("\nPick Table Insertion Point: ");
                    if (ptRes.Status == PromptStatus.OK)
                    {
                        var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                        var tbl = BuildCurveTableV24(ptRes.Value, curves);
                        btr.AppendEntity(tbl);
                        tr.AddNewlyCreatedDBObject(tbl, true);
                    }
                }
                tr.Commit();
            }
        }

        private int ParseIdNum(string id)
        {
            var match = Regex.Match(id ?? "", @"\d+");
            return match.Success ? int.Parse(match.Value) : 0;
        }

        // =========================================
        // DOMAIN MODELS & PARSING
        // =========================================

        class TextItem
        {
            public Point3d Pos;
            public string Text;
            public string Style;
        }

        class LineRec
        {
            public string ID;
            public string P_Bear, P_Dist, D_Bear, D_Dist, M_Bear, M_Dist;
            public string ExplicitRefTag;

            public string RefType 
            {
                get 
                {
                    if (!string.IsNullOrEmpty(ExplicitRefTag)) return ExplicitRefTag;
                    if (!string.IsNullOrEmpty(P_Bear) || !string.IsNullOrEmpty(P_Dist)) return "P";
                    if (!string.IsNullOrEmpty(D_Bear) || !string.IsNullOrEmpty(D_Dist)) return "D";
                    return "P";
                }
            }
            
            public void ApplyRefFallbacks()
            {
                bool isRefD = (RefType == "D" || RefType == "R");
                string rB = isRefD ? D_Bear : P_Bear;
                string rD = isRefD ? D_Dist : P_Dist;
                
                // Copy missing from M
                if (string.IsNullOrEmpty(rB)) rB = M_Bear;
                if (string.IsNullOrEmpty(rD)) rD = M_Dist;

                // Copy missing M from ref
                if (string.IsNullOrEmpty(M_Bear)) M_Bear = rB;
                if (string.IsNullOrEmpty(M_Dist)) M_Dist = rD;

                if (isRefD) { D_Bear = rB; D_Dist = rD; }
                else { P_Bear = rB; P_Dist = rD; }
            }
        }

        class CurveRec
        {
            public string ID;
            public string P_Rad, M_Rad, P_Arc, M_Arc, P_Delta, M_Delta, P_ChBrg, P_ChDst, M_ChBrg, M_ChDst;
        }

        class CurveRecV24
        {
            public string ID, R, L, DELTA, CHB, CHD, T;
        }

        class TableParser
        {
            public List<LineRec> Lines = new List<LineRec>();
            public List<CurveRec> Curves = new List<CurveRec>();

            public void ParseMixed(List<string> texts)
            {
                LineRec curLine = null;
                CurveRec curCurve = null;
                string mode = "";

                foreach (var block in texts)
                {
                    var sentences = block.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim());
                    foreach (var s in sentences)
                    {
                        var u = s.ToUpper();
                        if (Regex.IsMatch(u, @"^L-?\d+")) { mode = "L"; curLine = new LineRec { ID = s }; Lines.Add(curLine); continue; }
                        if (Regex.IsMatch(u, @"^C-?\d+")) { mode = "C"; curCurve = new CurveRec { ID = s }; Curves.Add(curCurve); continue; }

                        string tag = GetTag(u);
                        string strip = StripTags(s);

                        if (mode == "L" && curLine != null)
                        {
                            if (Regex.IsMatch(u, @"[NS].*[EW]"))
                            {
                                var parts = ParseBearingDist(strip);
                                ApplyLineData(curLine, tag, parts.bear, parts.dist);
                            }
                            else if (Regex.IsMatch(strip, @"^[\d\.]+"))
                            {
                                ApplyLineData(curLine, tag, null, strip);
                            }
                            else if (!string.IsNullOrEmpty(tag))
                            {
                                ApplyLineData(curLine, tag, null, strip);
                            }
                        }
                        else if (mode == "C" && curCurve != null)
                        {
                            if (u.Contains("R=") || u.Contains("R =")) ApplyCurveData(curCurve, tag, "R", ExtractVal(strip, "R"));
                            else if (u.Contains("L=") || u.Contains("L =")) ApplyCurveData(curCurve, tag, "L", ExtractVal(strip, "L"));
                            else if (u.Contains("=") || u.Contains("Δ") || u.Contains("D=") || u.Contains("DELTA")) ApplyCurveData(curCurve, tag, "D", ExtractValOffset(strip));
                            else if (Regex.IsMatch(u, @"[NS].*[EW]"))
                            {
                                var parts = ParseBearingDist(strip);
                                ApplyCurveData(curCurve, tag, "CHB", parts.bear);
                                ApplyCurveData(curCurve, tag, "CHD", parts.dist);
                            }
                        }
                    }
                }
            }

            string GetTag(string u)
            {
                if (u.Contains("(P&M)") || u.Contains("(P&A)") || u.Contains("(P & M)")) return "PM";
                if (u.Contains("(D&M)") || u.Contains("(D&A)")) return "DM";
                if (u.Contains("(R&M)") || u.Contains("(R&A)") || u.Contains("(R & M)")) return "RM";
                if (u.Contains("(P)")) return "P";
                if (u.Contains("(D)")) return "D";
                if (u.Contains("(R)")) return "R";
                if (u.Contains("(M)") || u.Contains("(A)")) return "M";
                return "";
            }

            string StripTags(string s) => Regex.Replace(s, @"\([PDMAR& ]+\)", "", RegexOptions.IgnoreCase).Trim();

            (string bear, string dist) ParseBearingDist(string s)
            {
                var match = Regex.Match(s, @"([NS][^0-9]*\d+.*?['""]?[EW])\s+(.*)", RegexOptions.IgnoreCase);
                if (match.Success) return (match.Groups[1].Value.Trim(), match.Groups[2].Value.Trim());
                // Fallback splitting on last space
                int ls = s.LastIndexOf(' ');
                if (ls > 0) return (s.Substring(0, ls).Trim(), s.Substring(ls + 1).Trim());
                return (s, "");
            }

            string ExtractVal(string s, string key)
            {
                var match = Regex.Match(s, $@"{key}\s*=\s*(.*)", RegexOptions.IgnoreCase);
                return match.Success ? match.Groups[1].Value.Trim() : s;
            }

            string ExtractValOffset(string s)
            {
                int idx = s.IndexOf('=');
                if (idx < 0) idx = s.IndexOf('Δ');
                if (idx >= 0) return s.Substring(idx + 1).Trim();
                return s;
            }

            void ApplyLineData(LineRec r, string tag, string bear, string dist)
            {
                if (tag == "PM") { 
                    if (string.IsNullOrEmpty(r.ExplicitRefTag)) r.ExplicitRefTag = "P";
                    if (!string.IsNullOrEmpty(bear)) { r.P_Bear = bear; r.M_Bear = bear; }
                    if (!string.IsNullOrEmpty(dist)) { r.P_Dist = dist; r.M_Dist = dist; }
                }
                else if (tag == "DM" || tag == "RM") { 
                    if (string.IsNullOrEmpty(r.ExplicitRefTag)) r.ExplicitRefTag = tag == "DM" ? "D" : "R";
                    if (!string.IsNullOrEmpty(bear)) { r.D_Bear = bear; r.M_Bear = bear; }
                    if (!string.IsNullOrEmpty(dist)) { r.D_Dist = dist; r.M_Dist = dist; }
                }
                else if (tag == "P") { 
                    if (string.IsNullOrEmpty(r.ExplicitRefTag)) r.ExplicitRefTag = "P";
                    if (!string.IsNullOrEmpty(bear)) r.P_Bear = bear; 
                    if (!string.IsNullOrEmpty(dist)) r.P_Dist = dist; 
                }
                else if (tag == "D" || tag == "R") { 
                    if (string.IsNullOrEmpty(r.ExplicitRefTag)) r.ExplicitRefTag = tag;
                    if (!string.IsNullOrEmpty(bear)) r.D_Bear = bear; 
                    if (!string.IsNullOrEmpty(dist)) r.D_Dist = dist; 
                }
                else if (tag == "M") { 
                    if (!string.IsNullOrEmpty(bear)) r.M_Bear = bear; 
                    if (!string.IsNullOrEmpty(dist)) r.M_Dist = dist; 
                }
                else if (tag == "") {
                    if (r.P_Bear == null && r.P_Dist == null && r.D_Bear == null && r.D_Dist == null) {
                        if (!string.IsNullOrEmpty(bear)) r.P_Bear = bear;
                        if (!string.IsNullOrEmpty(dist)) r.P_Dist = dist;
                    } else if (r.M_Bear == null && r.M_Dist == null) {
                        if (!string.IsNullOrEmpty(bear)) r.M_Bear = bear;
                        if (!string.IsNullOrEmpty(dist)) r.M_Dist = dist;
                    }
                }
            }

            void ApplyCurveData(CurveRec r, string tag, string type, string val)
            {
                bool isP = (tag == "PM" || tag == "P" || tag == "D" || tag == "R" || tag == "DM" || tag == "RM");
                bool isM = (tag == "PM" || tag == "M" || tag == "DM" || tag == "RM" || tag == "A");
                if (type == "R") { if (isP) r.P_Rad = val; if (isM) r.M_Rad = val; }
                if (type == "L") { if (isP) r.P_Arc = val; if (isM) r.M_Arc = val; }
                if (type == "D") { if (isP) r.P_Delta = val; if (isM) r.M_Delta = val; }
                if (type == "CHB") { if (isP) r.P_ChBrg = val; if (isM) r.M_ChBrg = val; }
                if (type == "CHD") { if (isP) r.P_ChDst = val; if (isM) r.M_ChDst = val; }
            }
        }

        private CurveRecV24 ParseCurveV24(string text)
        {
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();
            if (lines.Count == 0) return null;

            string id = lines.FirstOrDefault(l => Regex.IsMatch(l, @"^C-?\d+")) ?? "";
            string R_pm = "", L_p = "", L_m = "", D_p = "", D_m = "", CB_p = "", CD_p = "", CB_m = "", CD_m = "";

            foreach (var l in lines)
            {
                string u = l.ToUpper();
                string clean = Regex.Replace(l, @"\([PDMAR& ]+\)", "", RegexOptions.IgnoreCase).Trim();

                if (u.Contains("R=") || u.Contains("R ="))
                {
                    R_pm = ExtractVal(clean, "R");
                }
                else if (u.Contains("L=") || u.Contains("L ="))
                {
                    string v = ExtractVal(l, "L");
                    ParsePM(v, out L_p, out L_m);
                }
                else if ((u.Contains("=") || u.Contains("Δ") || u.Contains("D=")) && !u.Contains("R=") && !u.Contains("L="))
                {
                    string v = ExtractValOffset(l);
                    ParsePM(v, out D_p, out D_m);
                }
                else if (Regex.IsMatch(u, @"[NS].*[EW]"))
                {
                    bool isP = u.Contains("(P)") || u.Contains("(D)") || u.Contains("(R)"); 
                    bool isM = u.Contains("(M)") || u.Contains("(A)");
                    var parts = ParseBearingDist(clean);
                    if (isP) { CB_p = parts.bear; CD_p = parts.dist; }
                    if (isM) { CB_m = parts.bear; CD_m = parts.dist; }
                }
            }

            double? rad = ParseNum(R_pm);
            double? dP = ParseDms(D_p);
            double? dM = ParseDms(D_m);
            string Tp = rad.HasValue && dP.HasValue ? (rad.Value * Math.Tan(dP.Value * Math.PI / 360.0)).ToString("F2") : "";
            string Tm = rad.HasValue && dM.HasValue ? (rad.Value * Math.Tan(dM.Value * Math.PI / 360.0)).ToString("F2") : "";

            string rTag = text.ToUpper().Contains("(D)") ? "D" : text.ToUpper().Contains("(R)") ? "R" : "P";
            string mTag = text.ToUpper().Contains("(A)") ? "A" : "M";

            return new CurveRecV24
            {
                ID = id,
                R = $"{R_pm}\n({rTag}&{mTag})",
                L = $"{L_p}({rTag})\n{L_m}({mTag})",
                DELTA = $"{D_p}({rTag})\n{D_m}({mTag})",
                CHB = $"{CB_p}({rTag})\n{CB_m}({mTag})",
                CHD = $"{CD_p}({rTag})\n{CD_m}({mTag})",
                T = $"{(Tp != "" ? Tp + "'(" + rTag + ")\n" : "")}{(Tm != "" ? Tm + "'(" + mTag + ")" : "")}"
            };
        }

        private void ParsePM(string val, out string pVal, out string mVal)
        {
            int pPos = val.IndexOf("(P)", StringComparison.OrdinalIgnoreCase);
            if (pPos < 0) pPos = val.IndexOf("(D)", StringComparison.OrdinalIgnoreCase);
            if (pPos < 0) pPos = val.IndexOf("(R)", StringComparison.OrdinalIgnoreCase);
            
            int mPos = val.IndexOf("(M)", StringComparison.OrdinalIgnoreCase);
            if (mPos < 0) mPos = val.IndexOf("(A)", StringComparison.OrdinalIgnoreCase);

            if (pPos >= 0 && mPos >= 0)
            {
                if (pPos < mPos) {
                    pVal = val.Substring(0, pPos).Trim();
                    mVal = val.Substring(pPos + 3, mPos - (pPos + 3)).Trim();
                } else {
                    mVal = val.Substring(0, mPos).Trim();
                    pVal = val.Substring(mPos + 3, pPos - (mPos + 3)).Trim();
                }
            }
            else
            {
                pVal = Regex.Replace(val, @"\([PDMAR]\)", "", RegexOptions.IgnoreCase).Trim();
                mVal = "";
            }
        }

        private double? ParseNum(string s)
        {
            var m = Regex.Match(s ?? "", @"[\d\.]+");
            return m.Success && double.TryParse(m.Value, out double d) ? d : (double?)null;
        }

        private double? ParseDms(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            string c = s.Replace("°", " ").Replace("d", " ").Replace("D", " ").Replace("'", " ").Replace("\"", " ").Replace("-", " ");
            var parts = c.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(x => double.TryParse(x, out var d) ? d : 0).ToList();
            if (parts.Count == 0) return null;
            double deg = parts[0];
            if (parts.Count > 1) deg += parts[1] / 60.0;
            if (parts.Count > 2) deg += parts[2] / 3600.0;
            return deg;
        }

        // These private helpers serve ParseCurveV24 (which lives outside the nested TableParser class).
        // Note: TableParser has its own identical copies for use within that class.
        private string ExtractVal(string s, string key)
        {
            var match = Regex.Match(s, $@"{key}\s*=\s*(.*)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : s;
        }

        private string ExtractValOffset(string s)
        {
            int idx = s.IndexOf('=');
            if (idx < 0) idx = s.IndexOf('Δ');
            if (idx >= 0) return s.Substring(idx + 1).Trim();
            return s;
        }

        private (string bear, string dist) ParseBearingDist(string s)
        {
            var match = Regex.Match(s, @"([NS][^0-9]*\d+.*?['""]?[EW])\s+(.*)", RegexOptions.IgnoreCase);
            if (match.Success) return (match.Groups[1].Value.Trim(), match.Groups[2].Value.Trim());
            int ls = s.LastIndexOf(' ');
            if (ls > 0) return (s.Substring(0, ls).Trim(), s.Substring(ls + 1).Trim());
            return (s, "");
        }


        // =========================================
        // TABLE GENERATION
        // =========================================
        private Table BuildLineTable(Point3d pos, List<LineRec> recs)
        {
            var tbl = new Table();
            tbl.Position = pos;
            tbl.InsertRows(0, 1.0, 2 + recs.Count * 2);
            tbl.InsertColumns(0, 15.0, 4);

            tbl.SetRowHeight(0.12);
            tbl.Columns[0].Width = 13.0; tbl.Columns[1].Width = 13.0; tbl.Columns[2].Width = 30.0; tbl.Columns[3].Width = 15.0;

            tbl.Cells[0, 0].TextString = "LINE TABLE";
            tbl.MergeCells(CellRange.Create(tbl, 0, 0, 0, 3));

            tbl.Cells[1, 0].TextString = "LINE"; tbl.Cells[1, 1].TextString = "TYPE";
            tbl.Cells[1, 2].TextString = "BEARING"; tbl.Cells[1, 3].TextString = "DIST";

            int r = 2;
            foreach (var l in recs)
            {
                l.ApplyRefFallbacks();
                bool isRefD = (l.RefType == "D" || l.RefType == "R");
                tbl.Cells[r, 0].TextString = l.ID ?? ""; tbl.Cells[r, 1].TextString = l.RefType ?? "";
                tbl.Cells[r, 2].TextString = (isRefD ? l.D_Bear : l.P_Bear) ?? "";
                tbl.Cells[r, 3].TextString = (isRefD ? l.D_Dist : l.P_Dist) ?? "";
                r++;
                tbl.Cells[r, 0].TextString = l.ID ?? ""; tbl.Cells[r, 1].TextString = "A";
                tbl.Cells[r, 2].TextString = l.M_Bear ?? ""; tbl.Cells[r, 3].TextString = l.M_Dist ?? "";
                r++;
            }

            FormatTable(tbl);
            return tbl;
        }

        private Table BuildCurveTable(Point3d pos, List<CurveRec> recs)
        {
            var tbl = new Table();
            tbl.Position = pos;
            tbl.InsertRows(0, 1.0, 2 + recs.Count);
            tbl.InsertColumns(0, 15.0, 11);

            tbl.Cells[0, 0].TextString = "CURVE TABLE";
            tbl.MergeCells(CellRange.Create(tbl, 0, 0, 0, 10));

            tbl.Cells[1, 0].TextString = "CURVE"; tbl.Cells[1, 1].TextString = "RADIUS (P)"; tbl.Cells[1, 2].TextString = "RADIUS (M)";
            tbl.Cells[1, 3].TextString = "ARC (P)"; tbl.Cells[1, 4].TextString = "ARC (M)"; tbl.Cells[1, 5].TextString = "DELTA (P)";
            tbl.Cells[1, 6].TextString = "DELTA (M)"; tbl.Cells[1, 7].TextString = "CH BRG (P)"; tbl.Cells[1, 8].TextString = "CH DST (P)";
            tbl.Cells[1, 9].TextString = "CH BRG (M)"; tbl.Cells[1, 10].TextString = "CH DST (M)";

            int r = 2;
            foreach (var c in recs)
            {
                tbl.Cells[r, 0].TextString = c.ID ?? ""; tbl.Cells[r, 1].TextString = c.P_Rad ?? ""; tbl.Cells[r, 2].TextString = c.M_Rad ?? "";
                tbl.Cells[r, 3].TextString = c.P_Arc ?? ""; tbl.Cells[r, 4].TextString = c.M_Arc ?? ""; tbl.Cells[r, 5].TextString = c.P_Delta ?? "";
                tbl.Cells[r, 6].TextString = c.M_Delta ?? ""; tbl.Cells[r, 7].TextString = c.P_ChBrg ?? ""; tbl.Cells[r, 8].TextString = c.P_ChDst ?? "";
                tbl.Cells[r, 9].TextString = c.M_ChBrg ?? ""; tbl.Cells[r, 10].TextString = c.M_ChDst ?? "";
                r++;
            }

            FormatTable(tbl);
            return tbl;
        }

        private Table BuildCurveTableV24(Point3d pos, List<CurveRecV24> recs)
        {
            var tbl = new Table();
            tbl.Position = pos;
            tbl.InsertRows(0, 10.0, 2 + recs.Count);
            tbl.InsertColumns(0, 20.0, 7);

            tbl.Columns[0].Width = 12.0; tbl.Columns[1].Width = 38.0; tbl.Columns[2].Width = 20.0; tbl.Columns[3].Width = 20.0;
            tbl.Columns[4].Width = 20.0; tbl.Columns[5].Width = 20.0; tbl.Columns[6].Width = 25.0;

            tbl.Cells[0, 0].TextString = "CURVE TABLE";
            tbl.MergeCells(CellRange.Create(tbl, 0, 0, 0, 6));

            tbl.Cells[1, 0].TextString = "CURVE"; tbl.Cells[1, 1].TextString = "CH BRG"; tbl.Cells[1, 2].TextString = "CH DST";
            tbl.Cells[1, 3].TextString = "ARC"; tbl.Cells[1, 4].TextString = "R"; tbl.Cells[1, 5].TextString = "T";
            tbl.Cells[1, 6].TextString = "DELTA";

            int r = 2;
            foreach (var c in recs)
            {
                tbl.Cells[r, 0].TextString = c.ID ?? ""; tbl.Cells[r, 1].TextString = c.CHB ?? ""; tbl.Cells[r, 2].TextString = c.CHD ?? "";
                tbl.Cells[r, 3].TextString = c.L ?? ""; tbl.Cells[r, 4].TextString = c.R ?? ""; tbl.Cells[r, 5].TextString = c.T ?? "";
                tbl.Cells[r, 6].TextString = c.DELTA ?? "";
                r++;
            }

            FormatTable(tbl, 1.8, 0.50, 0.25);
            return tbl;
        }

        private void FormatTable(Table tbl, double textHeight = 1.8, double marginH = 0.50, double marginV = 0.25)
        {
            for (int i = 0; i < tbl.Rows.Count; i++)
            {
                for (int j = 0; j < tbl.Columns.Count; j++)
                {
                    tbl.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
                    tbl.Cells[i, j].TextHeight = textHeight;
                    // TopEdge/BottomEdge is obsolete in ObjectARX .NET for Table Cell.
                    // AutoCAD tables apply default standard padding automatically which looks best natively.
                }
            }
            tbl.Cells[0, 0].TextHeight = textHeight * 1.35;
            for (int i = 0; i < tbl.Columns.Count; i++) tbl.Cells[1, i].TextHeight = textHeight * 1.15;
            tbl.GenerateLayout();
        }
    }
}
