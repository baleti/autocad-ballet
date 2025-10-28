using Autodesk.AutoCAD.Runtime;
using System.Diagnostics;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(AutoCADBallet.KillCurrentProcess))]

namespace AutoCADBallet
{
    public class KillCurrentProcess
    {
        [CommandMethod("kill-current-process", CommandFlags.Modal)]
        public void KillCurrentProcessCommand()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            var ed = doc?.Editor;

            try
            {
                // Get the current process (this AutoCAD instance)
                Process currentProcess = Process.GetCurrentProcess();
                int currentPid = currentProcess.Id;

                // Log what we're about to do
                if (ed != null)
                {
                    ed.WriteMessage($"\nKilling current AutoCAD process (PID: {currentPid})...\n");
                }

                // Forcefully terminate the current process
                // This will immediately kill AutoCAD without saving
                currentProcess.Kill();
            }
            catch (System.Exception ex)
            {
                if (ed != null)
                {
                    ed.WriteMessage($"\nError killing current process: {ex.Message}\n");
                }
            }
        }
    }
}
