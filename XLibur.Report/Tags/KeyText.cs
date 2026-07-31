using System;
using System.Globalization;

namespace XLibur.Report.Tags;

/// <summary>
/// Renders the value an expression produced as the text a report reads it as.
/// </summary>
/// <remarks>
/// Under the report's culture — the one the engine was constructed with — rather than the machine's.
/// A group label is text a human reads, so a date or a decimal in one should format the way the rest
/// of the report does, and a default-constructed template stays invariant end to end. The comparer
/// renders the same way for keys it can only order as text, so a label and the position it appears
/// in agree about what the key says.
/// </remarks>
internal static class KeyText
{
    public static string For(object? key, CultureInfo culture) => key switch
    {
        null => string.Empty,
        string text => text,
        IFormattable formattable => formattable.ToString(null, culture),
        _ => key.ToString() ?? string.Empty,
    };
}
