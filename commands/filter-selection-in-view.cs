using Autodesk.AutoCAD.Runtime;
using AutoCADBallet;

[assembly: CommandClass(typeof(FilterSelectionViewElements))]

/// <summary>
/// Command that always uses view scope for filtering selection, regardless of current selection scope setting
/// </summary>
public class FilterSelectionViewElements
{
    [CommandMethod("filter-selection-in-view", CommandFlags.UsePickSet | CommandFlags.Redraw | CommandFlags.Modal)]
    public void FilterSelectionViewCommand()
    {
        var command = new FilterSelectionViewImpl();
        command.Execute();
    }
}

public class FilterSelectionViewImpl : FilterElementsBase
{
    public override bool SpanAllScreens => false;
    public override bool UseSelectedElements => false; // Always use view scope - prompt for selection
    public override bool IncludeProperties => true;
    public override SelectionScope Scope => SelectionScope.view;
}