using System;
using System.Text.RegularExpressions;

public partial class CustomGUIs
{
    // Utility helpers extracted from DataGrid2_Helpers.cs

    private static string StripQuotes(string s)
    {
        return s.StartsWith("\"") && s.EndsWith("\"") && s.Length > 1
            ? s.Substring(1, s.Length - 2)
            : s;
    }

    /// <summary>Check if a string contains glob wildcards</summary>
    private static bool ContainsGlobWildcards(string pattern)
    {
        return pattern != null && pattern.Contains("*");
    }

    /// <summary>Convert glob pattern to regex pattern</summary>
    private static string GlobToRegexPattern(string globPattern)
    {
        // Escape special regex characters except *
        string escaped = Regex.Escape(globPattern).Replace("\\*", ".*");
        return "^" + escaped + "$";
    }

    /// <summary>Check if a value matches a glob pattern</summary>
    private static bool MatchesGlobPattern(string value, string pattern)
    {
        if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(pattern))
            return false;

        // Convert to lowercase for case-insensitive matching
        value = value.ToLowerInvariant();
        pattern = pattern.ToLowerInvariant();

        // If no wildcards, use simple contains (backward compatibility)
        if (!pattern.Contains("*"))
            return value.Contains(pattern);

        // Convert glob to regex and match
        string regexPattern = GlobToRegexPattern(pattern);
        return Regex.IsMatch(value, regexPattern);
    }
}

