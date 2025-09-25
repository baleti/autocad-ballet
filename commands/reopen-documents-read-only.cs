using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(AutoCADBallet.ReopenDocumentsReadOnlyCommand))]

namespace AutoCADBallet
{
    public class ReopenDocumentsReadOnlyCommand
    {
        [CommandMethod("reopen-documents-read-only", CommandFlags.Session)]
        public void ReopenDocumentsReadOnly()
        {
            ReopenDocumentsUtility.ReopenDocuments(forceReadOnly: true);
        }
    }
}