using System;
using System.Collections.Generic;

namespace XLibur.Report.Tags;

/// <summary>A tag as written in a cell: its name and its parameters.</summary>
public sealed class TagToken
{
    internal TagToken(string name, IReadOnlyDictionary<string, string> parameters)
    {
        Name = name;
        Parameters = parameters;
    }

    /// <summary>The tag name, as written.</summary>
    public string Name { get; }

    /// <summary>
    /// The tag's parameters. A bare parameter — <c>desc</c> rather than <c>desc=true</c> — is
    /// present with an empty value.
    /// </summary>
    public IReadOnlyDictionary<string, string> Parameters { get; }

    /// <summary>Whether <paramref name="name"/> was given, with or without a value.</summary>
    public bool Has(string name) => Parameters.ContainsKey(name);

    /// <summary>The value of <paramref name="name"/>, or <paramref name="fallback"/> if absent or empty.</summary>
    public string Value(string name, string fallback = "") =>
        Parameters.TryGetValue(name, out var value) && value.Length > 0 ? value : fallback;

    /// <summary>The value of <paramref name="name"/> as a number, or <paramref name="fallback"/>.</summary>
    public double Number(string name, double fallback)
    {
        var text = Value(name);
        return double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var value)
            ? value
            : fallback;
    }

    /// <summary>
    /// Whether a flag is on: present with no value, or present with a value that reads as true.
    /// </summary>
    public bool Flag(string name)
    {
        if (!Parameters.TryGetValue(name, out var value))
        {
            return false;
        }

        return value.Length == 0
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "1", StringComparison.Ordinal);
    }
}
