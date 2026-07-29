using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace XLibur.Report.Tags;

/// <summary>
/// Reads <c>&lt;&lt;Tag param=value&gt;&gt;</c> markers out of cell text.
/// </summary>
/// <remarks>
/// Parameters may be bare (<c>desc</c>), assigned (<c>over=item.Total</c>) or quoted
/// (<c>title="Grand total"</c>). Quoting is what allows a value to contain spaces or a
/// <c>&gt;</c>.
/// </remarks>
internal static class TagParser
{
    private static readonly Regex Pattern = new(
        @"<<\s*(?<name>[A-Za-z_][A-Za-z0-9_]*)(?<params>(?:[^>""]|""[^""]*"")*)>>",
        RegexOptions.Compiled | RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(5));

    /// <summary>Whether <paramref name="text"/> holds at least one tag.</summary>
    public static bool Contains(string? text) => !string.IsNullOrEmpty(text) && Pattern.IsMatch(text);

    /// <summary>Reads every tag in <paramref name="text"/>, in the order written.</summary>
    public static List<TagToken> Parse(string? text)
    {
        var tokens = new List<TagToken>();

        if (string.IsNullOrEmpty(text))
        {
            return tokens;
        }

        foreach (Match match in Pattern.Matches(text!))
        {
            tokens.Add(new TagToken(match.Groups["name"].Value, ParseParameters(match.Groups["params"].Value)));
        }

        return tokens;
    }

    /// <summary>Returns <paramref name="text"/> with its tags removed.</summary>
    public static string Strip(string? text) =>
        string.IsNullOrEmpty(text) ? string.Empty : Pattern.Replace(text!, string.Empty).Trim();

    private static Dictionary<string, string> ParseParameters(string text)
    {
        var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var index = 0;

        while (index < text.Length)
        {
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            var nameStart = index;
            while (index < text.Length && !char.IsWhiteSpace(text[index]) && text[index] != '=')
            {
                index++;
            }

            if (index == nameStart)
            {
                index++;
                continue;
            }

            var name = text.Substring(nameStart, index - nameStart);

            if (index < text.Length && text[index] == '=')
            {
                index++;
                parameters[name] = ReadValue(text, ref index);
            }
            else
            {
                // A bare parameter is a flag: <<Sort desc>>.
                parameters[name] = string.Empty;
            }
        }

        return parameters;
    }

    private static string ReadValue(string text, ref int index)
    {
        if (index < text.Length && text[index] == '"')
        {
            index++;
            var value = new StringBuilder();
            while (index < text.Length && text[index] != '"')
            {
                value.Append(text[index]);
                index++;
            }

            index++;
            return value.ToString();
        }

        var start = index;
        while (index < text.Length && !char.IsWhiteSpace(text[index]))
        {
            index++;
        }

        return text.Substring(start, index - start);
    }
}
