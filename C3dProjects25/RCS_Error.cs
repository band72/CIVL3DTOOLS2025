#nullable enable
using System;
using System.Diagnostics;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace RCS.C3D2025
{
    internal static class RcsError
    {
        public const string LogPath = @"C:\temp\c3doutput.txt";

        public static void Info(Editor? ed, string message) => Write("INFO", message, null, ed);
        public static void Warn(Editor? ed, string message) => Write("WARN", message, null, ed);
        public static void Fail(Editor? ed, string message, Exception? ex = null) => Write("ERROR", message, ex, ed);

        /// <summary>
        /// Wrap a command so it never hard-crashes Civil 3D. Writes detailed errors to LogPath.
        /// </summary>
        public static void RunCommandSafe(string commandName, Action action)
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            var sw = Stopwatch.StartNew();
            try
            {
                Info(ed, $"--- {commandName} START --- Drawing={doc?.Name ?? "(null)"}");

                action();

                Info(ed, $"--- {commandName} END OK ({sw.Elapsed.TotalSeconds:0.00}s) ---");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception aex)
            {
                Fail(ed, $"{commandName} AutoCAD Exception: {aex.ErrorStatus} - {aex.Message}", null);
                if (ed != null) ed.WriteMessage($"\n{commandName} FAILED: {aex.ErrorStatus} (see {LogPath})");
            }
            catch (IOException ioex)
            {
                Fail(ed, $"{commandName} IO Exception: {ioex.Message}", null);
                if (ed != null) ed.WriteMessage($"\n{commandName} FAILED (I/O). See {LogPath}");
            }
            catch (UnauthorizedAccessException uaex)
            {
                Fail(ed, $"{commandName} Unauthorized: {uaex.Message}", null);
                if (ed != null) ed.WriteMessage($"\n{commandName} FAILED (permissions). See {LogPath}");
            }
            catch (System.Exception ex)
            {
                Fail(ed, $"{commandName} Exception: {ex.Message}", null);
                if (ed != null) ed.WriteMessage($"\n{commandName} FAILED. See {LogPath}");
            }
        }

        private static void Write(string level, string message, Exception? ex, Editor? ed)
        {
            try
            {
                var dir = Path.GetDirectoryName(LogPath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var doc = Application.DocumentManager.MdiActiveDocument;
                var dwg = doc?.Name ?? "(no active doc)";

                var log = $"[{now}] [{level}] [DWG={dwg}] {message}";
                File.AppendAllText(LogPath, log + Environment.NewLine);

                if (ex is not null)
                {
                    File.AppendAllText(
                        LogPath,
                        $"[{now}] [{level}] EXCEPTION TYPE: {ex.GetType().FullName}{Environment.NewLine}" +
                        $"[{now}] [{level}] STACK:{Environment.NewLine}{ex}{Environment.NewLine}"
                    );
                }

                // Minimal command-line feedback (avoid spam)
                if (ed != null && (level == "ERROR" || level == "WARN"))
                    ed.WriteMessage($"\n{level}: {message}");
            }
            catch
            {
                // Never throw from logging.
            }
        }
    }
}
