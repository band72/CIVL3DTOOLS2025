using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace C3dProjects25
{
    public class RcsDrawChecklistCommand
    {
        [CommandMethod("RCS_DRAW_CHECKLIST")]
        public void DrawChecklist()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            var window = new DrawChecklistWindow(doc.Name);
            bool? dialogResult = Autodesk.AutoCAD.ApplicationServices.Core.Application.ShowModalWindow(window);
            
            if (dialogResult != true) 
                return;

            var jobNo = window.txtJobNo.Text ?? "";
            var dateStr = window.txtDate.Text ?? "";
            var surveyType = window.cmbSurveyType.Text ?? "";
            var pSurvey = window.chkPriorSurvey.IsChecked == true ? "Yes" : "No";
            var rev3rd = window.chk3rdParty.IsChecked == true ? "Yes" : "No";
            var attention = window.txtAttention.Text ?? "";
            
            var items = window.Items.ToList();

            var pr = ed.GetPoint("\nSelect insertion point for Checklist Table: ");
            if (pr.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                // Ensure layer exists
                var lyTbl = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                string chkLayer = "0-RCS-Checklist";
                if (!lyTbl.Has(chkLayer))
                {
                    lyTbl.UpgradeOpen();
                    var ltr = new LayerTableRecord { Name = chkLayer };
                    lyTbl.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }

                var tbl = new Table();
                tbl.Position = pr.Value;
                tbl.Layer = chkLayer;
                
                int metaRows = 5; 
                int headerRows = 2; // Title + Column Headers
                int itemRows = items.Count;
                int totalRows = metaRows + headerRows + itemRows;

                tbl.InsertRows(0, 1.5, totalRows);
                tbl.InsertColumns(0, 20.0, 2);

                tbl.Columns[0].Width = 12.0;  // Approved Y/N
                tbl.Columns[1].Width = 100.0; // Description

                // Title config
                tbl.Cells[0, 0].TextString = "CAD DRAWING CHECKLIST";
                tbl.MergeCells(CellRange.Create(tbl, 0, 0, 0, 1));

                // Metadata Rows
                tbl.Cells[1, 0].TextString = "DWG Name:";
                tbl.Cells[1, 1].TextString = jobNo;
                
                tbl.Cells[2, 0].TextString = "Date:";
                tbl.Cells[2, 1].TextString = dateStr;
                
                tbl.Cells[3, 0].TextString = "Type:";
                tbl.Cells[3, 1].TextString = surveyType;
                
                tbl.Cells[4, 0].TextString = "Reviewed/Prior:";
                tbl.Cells[4, 1].TextString = $"3rd Party: {rev3rd} | Prior: {pSurvey}";

                // Attention row
                tbl.Cells[5, 0].TextString = "Needs Attention:";
                tbl.Cells[5, 1].TextString = attention;

                // Headers
                tbl.Cells[6, 0].TextString = "APPROVED";
                tbl.Cells[6, 1].TextString = "ITEM DESCRIPTION";

                int r = 7;
                foreach (var itm in items)
                {
                    tbl.Cells[r, 0].TextString = itm.IsApproved ? "[ X ]" : "[   ]";
                    tbl.Cells[r, 1].TextString = itm.Description;
                    r++;
                }

                // Format Table
                for (int i = 0; i < tbl.Rows.Count; i++)
                {
                    for (int j = 0; j < tbl.Columns.Count; j++)
                    {
                        tbl.Cells[i, j].Alignment = CellAlignment.MiddleLeft;
                        tbl.Cells[i, j].TextHeight = 1.0;
                    }
                }
                
                // Specific alignments
                tbl.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;
                tbl.Cells[0, 0].TextHeight = 1.35;
                tbl.Cells[6, 0].Alignment = CellAlignment.MiddleCenter;
                tbl.Cells[6, 1].Alignment = CellAlignment.MiddleCenter;
                tbl.Cells[6, 0].TextHeight = 1.15;
                tbl.Cells[6, 1].TextHeight = 1.15;
                
                for(int i = 7; i < totalRows; i++)
                {
                   tbl.Cells[i, 0].Alignment = CellAlignment.MiddleCenter; 
                }

                tbl.GenerateLayout();
                btr.AppendEntity(tbl);
                tr.AddNewlyCreatedDBObject(tbl, true);

                tr.Commit();
            }

            ed.WriteMessage("\nChecklist Table Created Successfully.");
        }
    }
}
