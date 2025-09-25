using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(AutoCADBallet.ReopenDocumentsReadWriteCommand))]

namespace AutoCADBallet
{
    public class ReopenDocumentsReadWriteCommand
    {
        [CommandMethod("reopen-documents-read-write", CommandFlags.Session)]
        public void ReopenDocumentsReadWrite()
        {
            ReopenDocumentsUtility.ReopenDocuments(forceReadOnly: false);
        }
    }
}