using System;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace RCS.C3D2025.Tools
{
    public class CleanupDrawingCommand
    {
        [CommandMethod("RCS_CLEANUP_DRAWING")]
        public void RunCleanupDrawing()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                // 1. Run ConvertCogoCodes automatically
                ed.WriteMessage("\n--- Running Cleanup: Convert COGO Codes ---");
                var convertCmd = new ConvertCogoCodesCommand();
                convertCmd.ExecuteConvertCogoCodes(true); // true = autoSelectAll

                // 2. Search drawing for SQ-FT and ACRES to enforce +/- symbol
                ed.WriteMessage("\n--- Running Cleanup: Verifying SQ-FT and ACRES symbols ---");

                int updatedTextCount = 0;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    
                    // Iterate through all layouts and model space
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                        if (btr.IsLayout)
                        {
                            foreach (ObjectId objId in btr)
                            {
                                // Fast check by object class before opening
                                if (objId.ObjectClass.Name == "AcDbText" || objId.ObjectClass.Name == "AcDbMText")
                                {
                                    var obj = tr.GetObject(objId, OpenMode.ForRead);

                                    if (obj is DBText dbText)
                                    {
                                        if (ProcessText(dbText.TextString, out string newText))
                                        {
                                            dbText.UpgradeOpen();
                                            dbText.TextString = newText;
                                            updatedTextCount++;
                                        }
                                    }
                                    else if (obj is MText mText)
                                    {
                                        if (ProcessText(mText.Contents, out string newContent))
                                        {
                                            mText.UpgradeOpen();
                                            mText.Contents = newContent;
                                            updatedTextCount++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nCleanup complete. Drawing texts updated: {updatedTextCount}\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError during cleanup drawing: {ex.Message}");
            }
        }

        private bool ProcessText(string text, out string newText)
        {
            newText = text;
            if (string.IsNullOrWhiteSpace(text)) return false;

            // Fast check if it contains the targets at all
            if (text.IndexOf("SQ", StringComparison.OrdinalIgnoreCase) < 0 && 
                text.IndexOf("ACRES", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            // Check if it already has a plus-minus symbol somewhere
            bool hasSymbol = text.Contains("±") || 
                             text.IndexOf("%%p", StringComparison.OrdinalIgnoreCase) >= 0 || 
                             text.IndexOf("\\U+00B1", StringComparison.OrdinalIgnoreCase) >= 0;

            bool changed = false;

            // Replace instances of ACRES or SQ-FT with acronyms, and inject ± before them if missing
            string replaced = Regex.Replace(text, @"\b(SQ\.?[- \s]*FT\.?|ACRES)(?!\w)", match =>
            {
                changed = true;
                string newValue;
                string originalPattern = match.Value.ToUpper();
                
                if (originalPattern.Contains("ACRE"))
                {
                    newValue = "Ac";
                }
                else
                {
                    newValue = "SF";
                }

                // If no symbol exists anywhere in the text, we prepend \U+00B1 
                if (!hasSymbol)
                {
                    newValue = "\\U+00B1 " + newValue;
                    hasSymbol = true; // prevent adding it twice if there are multiple matches
                }

                return newValue;
            }, RegexOptions.IgnoreCase);

            if (changed && replaced != text)
            {
                newText = replaced;
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage($"\n  [CleanupText] '{text.Replace("\n", "\\n")}' --> '{newText.Replace("\n", "\\n")}'");
                }
                return true;
            }

            return false;
        }
    }
}
