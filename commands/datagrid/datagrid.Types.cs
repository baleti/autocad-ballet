using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

public partial class CustomGUIs
{
    // Helper types extracted from DataGrid2_Helpers.cs

    private struct ColumnValueFilter
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public string       Value;         // value to look for in the cell
        public bool         IsExclusion;   // true ⇒ "must NOT contain"
        public bool         IsGlobPattern; // true if value contains wildcards
        public bool         IsExactMatch;  // true if value should match exactly
        public bool         IsColumnExactMatch; // true if column should match exactly
    }

    private enum ComparisonOperator
    {
        GreaterThan,
        LessThan
    }

    private struct ComparisonFilter
    {
        public List<string> ColumnParts;       // column-header fragments to match (null = all columns)
        public ComparisonOperator Operator;    // > or <
        public double Value;                   // numeric value to compare against
        public bool IsExclusion;               // true ⇒ "must NOT match comparison"
    }

    /// <summary>
    /// Represents column ordering information
    /// </summary>
    private struct ColumnOrderInfo
    {
        public List<string> ColumnParts;   // column-header fragments to match
        public int Position;               // desired position (1-based)
        public bool IsExactMatch;          // true if column should match exactly
    }

    /// <summary>
    /// Represents a group of filters that use AND logic internally.
    /// Multiple FilterGroups are combined with OR logic.
    /// </summary>
    private class FilterGroup
    {
        public List<List<string>> ColVisibilityFilters { get; set; }
        public List<ColumnValueFilter> ColValueFilters { get; set; }
        public List<string> GeneralFilters { get; set; }
        public List<ComparisonFilter> ComparisonFilters { get; set; }
        public List<string> GeneralGlobPatterns { get; set; } // glob patterns
        public List<ColumnOrderInfo> ColumnOrdering { get; set; } // column ordering
        public List<bool> ColVisibilityExactMatch { get; set; } // exact match for visibility
        public List<string> GeneralExactFilters { get; set; } // exact match general filters

        public FilterGroup()
        {
            ColVisibilityFilters = new List<List<string>>();
            ColValueFilters = new List<ColumnValueFilter>();
            GeneralFilters = new List<string>();
            ComparisonFilters = new List<ComparisonFilter>();
            GeneralGlobPatterns = new List<string>();
            ColumnOrdering = new List<ColumnOrderInfo>();
            ColVisibilityExactMatch = new List<bool>();
            GeneralExactFilters = new List<string>();
        }
    }

    /// <summary>
    /// Comparer for List<string> to use in HashSet
    /// </summary>
    private class ListStringComparer : IEqualityComparer<List<string>>
    {
        public bool Equals(List<string> x, List<string> y)
        {
            if (x == null && y == null) return true;
            if (x == null || y == null) return false;
            if (x.Count != y.Count) return false;
            return x.SequenceEqual(y);
        }

        public int GetHashCode(List<string> obj)
        {
            if (obj == null) return 0;
            int hash = 17;
            foreach (string s in obj)
                hash = hash * 31 + (s != null ? s.GetHashCode() : 0);
            return hash;
        }
    }

    private class SortCriteria
    {
        public string ColumnName { get; set; }
        public ListSortDirection Direction { get; set; }
    }

    /// <summary>
    /// Represents a cell edit operation
    /// </summary>
    private class CellEdit
    {
        public int RowIndex { get; set; }
        public string ColumnName { get; set; }
        public object OriginalValue { get; set; }
        public object NewValue { get; set; }
    }
}

