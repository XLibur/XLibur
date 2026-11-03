using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The model behind the formula conformance corpus (spec 59): case files holding the value Excel
/// produced for a formula, a reader for them, the comparison rules, and the XLibur side that
/// evaluates a case. The corpus is recorded with <c>tools/excel-conformance/Record-Corpus.ps1</c>.
/// </summary>
public static class FormulaConformance
{
    internal const string ResourceMarker = ".FormulaConformance.";
    public const string KnownDiffsFile = "known-diffs.tsv";
    private const string StampPrefix = "# excel_build:";
    private const string FormulaCell = "K1";

    internal static readonly string[] Types = ["number", "text", "bool", "error", "empty", "rejected"];

    /// <summary>
    /// A formula and what Excel computed for it. The value is stored the way the recorder writes it:
    /// a round-trip number, text with <c>\\ \t \n \r</c> escaped, <c>TRUE</c>/<c>FALSE</c>, or an
    /// error code.
    /// </summary>
    public sealed record Case(
        string Function, string Family, string Setup, string Formula,
        string ExcelValue, string ExcelType, string Note, string Id)
    {
        public override string ToString() => $"{Id} {Formula}";
    }

    /// <summary>A value in the same encoding as a stored Excel answer.</summary>
    public readonly record struct Answer(string Type, string Value)
    {
        public override string ToString() => Type + ":" + Value;
    }

    public sealed record KnownDiff(string Function, string Id, string XliburAnswer, string ExcelAnswer, string Reason, string Issue);

    public sealed class CorpusFileException(string message) : Exception(message);

    internal static IEnumerable<string> CorpusFileNames()
        => typeof(FormulaConformance).Assembly.GetManifestResourceNames()
            .Where(x => x.Contains(ResourceMarker, StringComparison.Ordinal) && x.EndsWith(".tsv", StringComparison.Ordinal))
            .Where(x => !x.EndsWith(KnownDiffsFile, StringComparison.Ordinal))
            .Order(StringComparer.Ordinal);

    private static StreamReader Open(string resourceName)
        => new(typeof(FormulaConformance).Assembly.GetManifestResourceStream(resourceName)!);

    /// <summary>All recorded cases, ids assigned per function in file order. Append, never reorder.</summary>
    internal static List<Case> ReadAllCases()
    {
        var ordinals = new Dictionary<string, int>(StringComparer.Ordinal);
        var cases = new List<Case>();
        foreach (var name in CorpusFileNames())
        {
            using var reader = Open(name);
            cases.AddRange(ReadCases(name, reader, ordinals));
        }

        return cases;
    }

    /// <summary>
    /// Reads one corpus file. A file with no Excel build stamp, or a row with no recorded answer,
    /// is rejected: the stored answers must come from Excel, never from XLibur.
    /// </summary>
    internal static List<Case> ReadCases(string fileName, TextReader reader, Dictionary<string, int> ordinals)
    {
        var stamped = false;
        var cases = new List<Case>();
        var lineNumber = 0;
        while (reader.ReadLine() is { } line)
        {
            lineNumber++;
            if (line.Length == 0)
                continue;

            if (line[0] == '#')
            {
                stamped |= line.StartsWith(StampPrefix, StringComparison.Ordinal)
                           && line.Length > StampPrefix.Length + 1;
                continue;
            }

            var f = line.Split('\t');
            if (f.Length < 7)
                throw new CorpusFileException($"{fileName}:{lineNumber} has {f.Length} columns, expected 7.");

            if (!Types.Contains(f[5]))
                throw new CorpusFileException($"{fileName}:{lineNumber} has no recorded Excel answer (type '{f[5]}'). Run the recorder.");

            ordinals[f[0]] = ordinals.GetValueOrDefault(f[0]) + 1;
            cases.Add(new Case(f[0], f[1], f[2], f[3], f[4], f[5], f[6], $"{f[0]}#{ordinals[f[0]]}"));
        }

        if (!stamped)
            throw new CorpusFileException($"{fileName} has no '# excel_build:' stamp, so its answers are not known to come from Excel.");

        return cases;
    }

    internal static List<KnownDiff> ReadKnownDiffs()
    {
        var name = typeof(FormulaConformance).Assembly.GetManifestResourceNames()
            .Single(x => x.EndsWith(KnownDiffsFile, StringComparison.Ordinal));
        using var reader = Open(name);
        var rows = new List<KnownDiff>();
        while (reader.ReadLine() is { } line)
        {
            if (line.Length == 0 || line[0] == '#')
                continue;

            var f = line.Split('\t');
            if (f.Length < 6)
                throw new CorpusFileException($"{KnownDiffsFile} row '{line}' has {f.Length} columns, expected 6.");

            rows.Add(new KnownDiff(f[0], f[1], f[2], f[3], f[4], f[5]));
        }

        return rows;
    }

    /// <summary>
    /// Whether XLibur's answer agrees with the stored Excel answer. Numbers match within relative
    /// 1e-12 (absolute 1e-300 near zero). Everything else must be identical: the type, the text
    /// ordinally, and the error code, so empty, zero and the empty string stay distinct. The one
    /// exception is a formula Excel rejects, which matches a rejection or any error from XLibur.
    /// </summary>
    internal static bool Matches(Answer excel, Answer actual)
    {
        // A formula Excel refuses at entry cannot exist in a workbook Excel wrote, so XLibur refusing
        // it or evaluating it to an error value are the same outcome. Evaluating it to a value is not.
        if (excel.Type == "rejected")
            return actual.Type is "rejected" or "error";

        if (excel.Type != actual.Type)
            return false;

        if (excel.Type != "number")
            return string.Equals(excel.Value, actual.Value, StringComparison.Ordinal);

        var e = double.Parse(excel.Value, CultureInfo.InvariantCulture);
        var a = double.Parse(actual.Value, CultureInfo.InvariantCulture);
        if (e.Equals(a))
            return true;

        return Math.Abs(e - a) <= Math.Max(1e-12 * Math.Max(Math.Abs(e), Math.Abs(a)), 1e-300);
    }

    internal static Answer Stored(Case c) => new(c.ExcelType, c.ExcelValue);

    /// <summary>
    /// Whether XLibur's answer is the one a known-differences row recorded (<c>type:value</c>).
    /// Numbers match within the comparison tolerance, because the maths library rounds the last
    /// digit differently between platforms; everything else must be identical.
    /// </summary>
    public static bool IsRecordedAnswer(string recorded, Answer actual)
    {
        var split = recorded.IndexOf(':', StringComparison.Ordinal);
        var known = split < 0
            ? new Answer(recorded, string.Empty)
            : new Answer(recorded[..split], recorded[(split + 1)..]);

        if (known.Type == "number" && actual.Type == "number")
            return Matches(known, actual);

        return known.Type == actual.Type && string.Equals(known.Value, actual.Value, StringComparison.Ordinal);
    }

    /// <summary>
    /// Evaluates a case in XLibur through the public API: the setup cells, the formula in
    /// <c>K1</c>, then the cell's value. A formula the parser refuses, at either step, is
    /// <c>rejected</c>, the counterpart of Excel refusing the entry. Any other exception is
    /// <c>throws</c>.
    /// </summary>
    internal static Answer Evaluate(Case c)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ApplySetup(ws, c.Setup);
        var cell = ws.Cell(FormulaCell);

        try
        {
            cell.FormulaA1 = c.Formula;
            return Encode(cell.Value);
        }
        catch (ExpressionParseException)
        {
            return new Answer("rejected", string.Empty);
        }
        catch (Exception ex)
        {
            return new Answer("throws", ex.GetType().Name);
        }
    }

    private static Answer Encode(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => new Answer("empty", string.Empty),
        XLDataType.Boolean => new Answer("bool", value.GetBoolean() ? "TRUE" : "FALSE"),
        XLDataType.Text => new Answer("text", Escape(value.GetText())),
        XLDataType.Error => new Answer("error", value.ToString(CultureInfo.InvariantCulture)),
        _ => new Answer("number", value.GetUnifiedNumber().ToString("R", CultureInfo.InvariantCulture)),
    };

    /// <summary>
    /// Escapes <c>\ tab LF CR</c>, and a surrogate that is not half of a valid pair as
    /// <c>\uXXXX</c>, which UTF-8 could not store.
    /// </summary>
    internal static string Escape(string text)
    {
        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            switch (ch)
            {
                case '\\': sb.Append("\\\\"); break;
                case '\t': sb.Append("\\t"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                default:
                    if (char.IsHighSurrogate(ch) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                        sb.Append(ch).Append(text[++i]);
                    else if (char.IsSurrogate(ch))
                        sb.Append("\\u").Append(((int)ch).ToString("X4", CultureInfo.InvariantCulture));
                    else
                        sb.Append(ch);
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// <c>A1=5;A2="x";A3=#N/A;A4=TRUE;A5==SUM(1,2)</c>. A value left empty is an empty cell, and
    /// the empty string is written as a formula so both sides build it the same way.
    /// </summary>
    internal static void ApplySetup(IXLWorksheet ws, string setup)
    {
        var i = 0;
        while (i < setup.Length)
        {
            var eq = setup.IndexOf('=', i);
            var cell = ws.Cell(setup[i..eq].Trim());
            i = eq + 1;

            if (i < setup.Length && setup[i] == '"')
            {
                var text = new StringBuilder();
                i++;
                while (true)
                {
                    if (setup[i] == '"')
                    {
                        if (i + 1 < setup.Length && setup[i + 1] == '"')
                        {
                            text.Append('"');
                            i += 2;
                            continue;
                        }

                        i++;
                        break;
                    }

                    text.Append(setup[i++]);
                }

                if (text.Length == 0)
                    cell.FormulaA1 = "=\"\"";
                else
                    cell.Value = text.ToString();
            }
            else
            {
                var end = setup.IndexOf(';', i);
                if (end < 0)
                    end = setup.Length;

                var raw = setup[i..end];
                i = end;
                if (raw.Length == 0)
                {
                    // An empty cell.
                }
                else if (raw is "TRUE" or "FALSE" || raw[0] is '=' or '#')
                    cell.FormulaA1 = raw[0] == '=' ? raw : "=" + raw;
                else
                    cell.Value = double.Parse(raw, CultureInfo.InvariantCulture);
            }

            if (i < setup.Length && setup[i] == ';')
                i++;
        }
    }
}
