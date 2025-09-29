using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(FilterSelectionDocumentElements))]

/// <summary>
/// Command that always uses document scope for filtering selection, regardless of current selection scope setting
/// </summary>
public class FilterSelectionDocumentElements
{
    [CommandMethod("filter-selection-in-document", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void FilterSelectionDocumentCommand()
    {
        var command = new FilterSelectionDocumentImpl();
        command.Execute();
    }
}

public class FilterSelectionDocumentImpl : FilterElementsBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => true; // Use stored selection from document scope
    public override bool IncludeProperties => true;
    public override SelectionScope Scope => SelectionScope.document;
}