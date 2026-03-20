// ============================================================
// RCS_QA_MASTER.cs
// Civil 3D 2025 – Unified QA Tagging + Validation System
// Updated: Supports Text, MText, and Images in Any Space
// ============================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

//[assembly: CommandClass(typeof(RCS.QA.RcsQaMaster))]

namespace RCS.QA
{
    public class RcsQaMaster : IExtensionApplication
    {
        const string AppName = "RCS_QA";
        const string LogPath = @"C:\temp\c3doutput.txt";

        #region INIT
        public void Initialize() => Log("INIT");
        public void Terminate() => Log("TERM");
        #endregion

        // =====================================================
        // COMMANDS
        // =====================================================

        /// <summary>
        /// Manual Tagger (CLI Version)
        /// </summary>
        [CommandMethod("RCS_QA_TAGGER")]
        public void QaTaggerCli()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            string currentType = "GEN";

            while (true)
            {
                var opt = new PromptEntityOptions($"\nSelect object (Text/Image) to tag as '{currentType}' [Type/Exit] <Exit>: ");
                opt.Keywords.Add("Type");
                opt.Keywords.Add("Exit");
                opt.AllowNone = true;
                opt.SetMessageAndKeywords($"\nSelect object (Text/Image) to tag as '{currentType}' [Type/Exit] <Exit>: ", "Type Exit");

                var res = ed.GetEntity(opt);

                if (res.Status == PromptStatus.Keyword)
                {
                    if (res.StringResult == "Type")
                    {
                        var typeOpt = new PromptStringOptions($"\nEnter new QA Type <{currentType}>: ") { AllowSpaces = false, DefaultValue = currentType, UseDefaultValue = true };
                        var typeRes = ed.GetString(typeOpt);
                        if (typeRes.Status == PromptStatus.OK && !string.IsNullOrWhiteSpace(typeRes.StringResult))
                        {
                            currentType = typeRes.StringResult.ToUpper();
                            ed.WriteMessage($"\nCurrent QA Type set to: {currentType}");
                        }
                    }
                    else break;
                }
                else if (res.Status == PromptStatus.OK)
                {
                    RunTransaction((tr, db) => {
                        var ent = tr.GetObject(res.ObjectId, OpenMode.ForWrite) as Entity;
                        if (ent != null)
                        {
                            var d = ReadXData(ent);
                            // Preserve existing ID, or create new
                            string id = string.IsNullOrEmpty(d.id) ? Guid.NewGuid().ToString("D") : d.id;
                            WriteXData(ent, id, currentType);
                            ed.WriteMessage($"\nTagged {ent.GetType().Name} handle={ent.Handle} as {currentType}.");
                        }
                    });
                }
                else break;
            }
            ed.WriteMessage("\nRCS_QA_TAGGER Completed.");
        }

        /// <summary>
        /// Auto-Tags ALL Text, MText, and Images in the current space.
        /// If an object is already tagged, it is skipped.
        /// </summary>
        [CommandMethod("RCS_QA_AUTOTAG")]
        public void AutoTag()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            int count = 0;

            RunOnLayoutEntities(ent =>
            {
                if (!HasXData(ent))
                {
                    string type = "GEN"; // Default type for auto-tag

                    // Optional: Auto-detect type based on object type?
                    if (ent is RasterImage) type = "IMAGE";

                    WriteXData(ent, Guid.NewGuid().ToString("D"), type);
                    count++;
                }
            });

            ed.WriteMessage($"\nRCS_QA_AUTOTAG: Tagged {count} new objects (Text/MText/Image).");
        }

        [CommandMethod("RCS_QA_AUTOTYPE")]
        public void AutoType()
        {
            RunOnLayoutEntities(ent =>
            {
                // Only run logic on Text/MText for content inference
                if (ent is DBText || ent is MText)
                {
                    var data = ReadXData(ent);
                    string type = InferType(ent);
                    WriteXData(ent,
                        string.IsNullOrWhiteSpace(data.id) ? Guid.NewGuid().ToString("D") : data.id,
                        type);
                }
            });
        }

        [CommandMethod("RCS_QA_FIX_DUPLICATES")]
        public void FixDuplicates()
        {
            RunTransaction((tr, db) =>
            {
                var map = new Dictionary<string, List<Entity>>();

                foreach (var ent in GetTargetEntities(tr, db))
                {
                    var d = ReadXData(ent);
                    if (string.IsNullOrWhiteSpace(d.id)) continue;
                    if (!map.ContainsKey(d.id)) map[d.id] = new();
                    map[d.id].Add(ent);
                }

                int fixedCt = 0;
                foreach (var g in map.Where(m => m.Value.Count > 1))
                {
                    foreach (var e in g.Value.Skip(1))
                    {
                        e.UpgradeOpen();
                        WriteXData(e, Guid.NewGuid().ToString("D"), ReadXData(e).type);
                        fixedCt++;
                    }
                }
                Log($"FIX_DUPLICATES={fixedCt}");
            });
        }

        [CommandMethod("RCS_QA_RUN")]
        public void RunQa()
        {
            EnsureDir();
            try { File.AppendAllText(LogPath, "\n=== QA RUN START ===\n"); } catch { }

            RunTransaction((tr, db) =>
            {
                foreach (var ent in GetTargetEntities(tr, db))
                {
                    var d = ReadXData(ent);
                    if (string.IsNullOrWhiteSpace(d.type)) continue;

                    // Images don't have text content to validate via Regex, usually passed or checked for existence
                    if (ent is RasterImage) continue;

                    string txt = GetText(ent);
                    var rule = ValidationRules.FirstOrDefault(r => r.Type == d.type);
                    if (rule == null) continue;

                    bool fail =
                        (rule.Required && string.IsNullOrWhiteSpace(txt)) ||
                        (rule.Regex != null && !rule.Regex.IsMatch(txt));

                    if (fail)
                    {
                        try
                        {
                            File.AppendAllText(LogPath, $"FAIL | TYPE={d.type} | ID={d.id} | HANDLE={ent.Handle} | TEXT=\"{txt}\"\n");
                        }
                        catch { }

                        ent.UpgradeOpen();
                        ent.ColorIndex = 1; // Red
                    }
                }
            });
            try { File.AppendAllText(LogPath, "=== QA RUN END ===\n"); } catch { }
        }

        // =====================================================
        // RULES
        // =====================================================

        class TypeRule
        {
            public string Type;
            public string Layer;
            public string Style;
            public Regex Content;
        }

        static readonly List<TypeRule> TypeRules = new()
        {
            new() { Type="SHEET_NO", Content=new Regex(@"SHEET\s*\d+",RegexOptions.IgnoreCase)},
            new() { Type="DATE", Content=new Regex(@"\d{1,2}/\d{1,2}/\d{4}")},
            new() { Type="SCALE", Content=new Regex(@"SCALE\s*:",RegexOptions.IgnoreCase)},
            new() { Type="CLIENT", Style="RCS_CLIENT"}
        };

        static string InferType(Entity e)
        {
            string t = GetText(e);
            string l = e.Layer;
            string s = e is DBText d ? d.TextStyleName :
                       e is MText m ? m.TextStyleName : "";

            foreach (var r in TypeRules)
            {
                if (r.Layer != null && !l.Contains(r.Layer)) continue;
                if (r.Style != null && !s.Equals(r.Style, StringComparison.OrdinalIgnoreCase)) continue;
                if (r.Content != null && !r.Content.IsMatch(t)) continue;
                return r.Type;
            }
            return "GEN";
        }

        class ValRule
        {
            public string Type;
            public bool Required;
            public Regex Regex;
        }

        static readonly List<ValRule> ValidationRules = new()
        {
            new(){Type="SHEET_NO",Required=true,Regex=new Regex(@"^\d+(\sOF\s\d+)?$",RegexOptions.IgnoreCase)},
            new(){Type="DATE",Required=true,Regex=new Regex(@"\d{1,2}/\d{1,2}/\d{4}")}
        };

        // =====================================================
        // CORE HELPERS
        // =====================================================

        static void RunOnLayoutEntities(Action<Entity> action)
            => RunTransaction((tr, db) =>
            {
                foreach (var ent in GetTargetEntities(tr, db))
                {
                    ent.UpgradeOpen();
                    action(ent);
                }
            });

        // UPDATED: Now returns DBText, MText, and RasterImage
        static IEnumerable<Entity> GetTargetEntities(Transaction tr, Database db)
        {
            // db.CurrentSpaceId handles Model vs Paper automatically
            var btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

            foreach (ObjectId id in btr)
            {
                // Quick type check before opening improves performance slightly, 
                // but usually we just open. Safe pattern:
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                if (ent is DBText || ent is MText || ent is RasterImage)
                {
                    yield return ent;
                }
            }
        }

        static (string id, string type) ReadXData(Entity e)
        {
            var rb = e.GetXDataForApplication(AppName);
            if (rb == null) return ("", "");
            string id = "", type = "";
            foreach (TypedValue tv in rb)
            {
                if (tv.Value is string s)
                {
                    if (s.StartsWith("ID=")) id = s[3..];
                    if (s.StartsWith("TYPE=")) type = s[5..];
                }
            }
            return (id, type);
        }

        static void WriteXData(Entity e, string id, string type)
        {
            e.XData = new ResultBuffer(
                new TypedValue(1001, AppName),
                new TypedValue(1000, "VER=1"),
                new TypedValue(1000, $"ID={id}"),
                new TypedValue(1000, $"TYPE={type}")
            );
        }

        static bool HasXData(Entity e) => e.GetXDataForApplication(AppName) != null;

        static string GetText(Entity e) => e is DBText d ? d.TextString :
                                           e is MText m ? m.Contents : "";

        static void RunTransaction(Action<Transaction, Database> a)
        {
            var d = Application.DocumentManager.MdiActiveDocument;
            using (d.LockDocument())
            using (var tr = d.Database.TransactionManager.StartTransaction())
            {
                EnsureRegApp(tr, d.Database);
                a(tr, d.Database);
                tr.Commit();
            }
        }

        static void EnsureRegApp(Transaction tr, Database db)
        {
            var t = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!t.Has(AppName))
            {
                t.UpgradeOpen();
                var r = new RegAppTableRecord { Name = AppName };
                t.Add(r); tr.AddNewlyCreatedDBObject(r, true);
            }
        }

        static void EnsureDir()
        {
            try { if (!Directory.Exists(@"C:\temp")) Directory.CreateDirectory(@"C:\temp"); } catch { }
        }

        static void Log(string s)
        {
            EnsureDir();
            try { File.AppendAllText(LogPath, $"[{DateTime.Now}] {s}\n"); } catch { }
        }
    }
}