using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Table = Autodesk.AutoCAD.DatabaseServices.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;


namespace Civil3D_SymbolTable
{
    public class Commands
    {
        private class PointSymbolData
        {
            public string Code { get; set; }
            public string Description { get; set; }
            public ObjectId BlockId { get; set; }
        }

        [CommandMethod("RCS_CreateSymbolTable")]
        public void CreateSymbolTable()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // --- STEP 1: Gather Point Data ---
                    var cogoPointIds = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.CogoPoints;

                    if (cogoPointIds.Count == 0)
                    {
                        ed.WriteMessage("\nError: No Cogo Points found.");
                        return;
                    }

                    List<PointSymbolData> tableData = new List<PointSymbolData>();
                    HashSet<string> processedCodes = new HashSet<string>();

                    // Open BlockTable once for lookups
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                    foreach (ObjectId pointId in cogoPointIds)
                    {
                        CogoPoint point = tr.GetObject(pointId, OpenMode.ForRead) as CogoPoint;
                        string code = point.RawDescription;

                        if (string.IsNullOrWhiteSpace(code)) continue;

                        string codeUpper = code.ToUpper();

                        if (!processedCodes.Contains(codeUpper))
                        {
                            // --- CRASH FIX START ---
                            // If the point style is set to <default> (Point Group controlled), StyleId is Null.
                            ObjectId styleId = point.StyleId;
                            if (styleId.IsNull)
                            {
                                // Skip points that don't have an explicit style assigned
                                continue;
                            }
                            // --- CRASH FIX END ---

                            PointStyle style = tr.GetObject(styleId, OpenMode.ForRead) as PointStyle;
                            ObjectId symbolBlockId = ObjectId.Null;

                            // Check if style uses a Block (Symbol)
                            if (style.MarkerType == PointMarkerDisplayType.UseSymbolForMarker)
                            {
                                string blockName = style.MarkerSymbolName;
                                if (!string.IsNullOrEmpty(blockName) && bt.Has(blockName))
                                {
                                    symbolBlockId = bt[blockName];
                                }
                            }

                            // RULE: Only add if a valid Block Symbol was found
                            if (symbolBlockId != ObjectId.Null)
                            {
                                processedCodes.Add(codeUpper);

                                tableData.Add(new PointSymbolData
                                {
                                    Code = codeUpper,
                                    Description = point.FullDescription,
                                    BlockId = symbolBlockId
                                });
                            }
                        }
                    }

                    if (tableData.Count == 0)
                    {
                        ed.WriteMessage("\nError: No points with valid assigned Block Styles found.");
                        return;
                    }

                    // --- STEP 2: Ask User for Location ---
                    PromptPointOptions ppo = new PromptPointOptions("\nSelect top-left corner for Symbol Table: ");
                    PromptPointResult ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK) return;

                    // --- STEP 3: Generate the Table ---
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    Table tb = new Table();
                    tb.TableStyle = db.Tablestyle;
                    tb.Position = ppr.Value;
                    tb.SetSize(tableData.Count + 2, 3);

                    // Columns
                    tb.Columns[0].Width = 15.0;
                    tb.Columns[1].Width = 40.0;
                    tb.Columns[2].Width = 15.0;

                    // Row Heights
                    double rowHeight = 8.0;
                    tb.Rows[0].Height = rowHeight;
                    tb.Rows[1].Height = rowHeight;

                    // Title
                    tb.Cells[0, 0].TextString = "SYMBOL LEGEND";
                    tb.Cells[0, 0].TextHeight = 3.5;
                    tb.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;
                    tb.MergeCells(CellRange.Create(tb, 0, 0, 0, 2));

                    // Headers
                    string[] headers = { "CODE", "DESCRIPTION", "SYMBOL" };
                    for (int i = 0; i < 3; i++)
                    {
                        tb.Cells[1, i].TextString = headers[i];
                        tb.Cells[1, i].TextHeight = 3.0;
                        tb.Cells[1, i].Alignment = CellAlignment.MiddleCenter;
                    }

                    // --- STEP 4: Fill Data ---
                    int row = 2;
                    foreach (var data in tableData.OrderBy(x => x.Code))
                    {
                        tb.Rows[row].Height = rowHeight;

                        // Col 0: Code
                        tb.Cells[row, 0].TextString = data.Code.ToUpper();
                        tb.Cells[row, 0].TextHeight = 2.5;
                        tb.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;

                        // Col 1: Description
                        tb.Cells[row, 1].TextString = (data.Description ?? "").ToUpper();
                        tb.Cells[row, 1].TextHeight = 2.5;
                        tb.Cells[row, 1].Alignment = CellAlignment.MiddleLeft;

                        // Col 2: Symbol
                        // We already filtered out null blocks
                        tb.Cells[row, 2].BlockTableRecordId = data.BlockId;

                        // Scaling Logic: 50% Reduction via Padding
                        tb.SetAutoScale(row, 2, true);

                        double margin = rowHeight * 0.25; // 25% top + 25% bottom = 50% Used space

                        tb.SetMargin(row, 2, CellMargins.Top, margin);
                        tb.SetMargin(row, 2, CellMargins.Bottom, margin);
                        tb.SetMargin(row, 2, CellMargins.Left, 1.0);
                        tb.SetMargin(row, 2, CellMargins.Right, 1.0);

                        tb.Cells[row, 2].Alignment = CellAlignment.MiddleCenter;

                        row++;
                    }

                    btr.AppendEntity(tb);
                    tr.AddNewlyCreatedDBObject(tb, true);
                    tr.Commit();

                    ed.WriteMessage($"\nSuccess: Generated table with {tableData.Count} symbols.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }
}