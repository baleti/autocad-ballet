using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

// Add alias to avoid ambiguity
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutocadBallet.TestCommands
{
    public class HelloWorldCommands
    {
        private static int _counter = 0;

        [CommandMethod("HELLO")]
        public void HelloWorld()
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            _counter++;

            ed.WriteMessage($"\n🎈 eurbgskbgueskr from autocad-ballet! (Call #{_counter})");
            ed.WriteMessage($"\n   Assembly loaded at: {DateTime.Now:HH:mm:ss.fff}");
            ed.WriteMessage($"\n   You can modify and rebuild this DLL while AutoCAD is running!");
        }

        [CommandMethod("TESTINPUT")]
        public void TestInput()
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            var result = ed.GetString("\nEnter some text: ");
            if (result.Status == PromptStatus.OK)
            {
                ed.WriteMessage($"\nYou entered: {result.StringResult}");
                ed.WriteMessage($"\n   Reversed: {ReverseString(result.StringResult)}");
            }
        }

        [CommandMethod("TIMESTAMP")]
        public void ShowTimestamp()
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            ed.WriteMessage($"\n⏰ Current time: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            ed.WriteMessage($"\n   Ticks: {DateTime.Now.Ticks}");
        }

        private string ReverseString(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }
    }

    public class MathCommands
    {
        [CommandMethod("FIBONACCI")]
        public void Fibonacci()
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;

            var result = ed.GetInteger("\nEnter number of Fibonacci terms to generate: ");
            if (result.Status == PromptStatus.OK && result.Value > 0)
            {
                ed.WriteMessage($"\nFibonacci sequence ({result.Value} terms):");

                long a = 0, b = 1;
                for (int i = 0; i < result.Value; i++)
                {
                    if (i == 0)
                        ed.WriteMessage($"\n  {i + 1}: {a}");
                    else if (i == 1)
                        ed.WriteMessage($"\n  {i + 1}: {b}");
                    else
                    {
                        long temp = a + b;
                        a = b;
                        b = temp;
                        ed.WriteMessage($"\n  {i + 1}: {b}");
                    }
                }
            }
        }
    }
}
