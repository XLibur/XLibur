using System.Text;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Fuzz;

/// <summary>The targets a fuzzing run can drive, and what each of them counts as acceptable.</summary>
internal static class FuzzTargets
{
    public const string Workbook = "workbook";
    public const string StructuredWorkbook = "workbook-structured";
    public const string Formula = "formula";
    public const string Address = "address";

    public static readonly string[] All = [Workbook, StructuredWorkbook, Formula, Address];

    public static bool IsKnown(string target)
    {
        return Array.Exists(All, t => t.Equals(target, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Run one input through a target. The returned string describes what happened when nothing
    /// went wrong; fuzzing discards it, replay prints it.
    ///
    /// It exists because "no failure" is ambiguous in a way that matters. A package XLibur
    /// *rejected* and a package XLibur *round-tripped* both leave the pipeline without throwing,
    /// but they say opposite things about the generator: if the structure-aware target only ever
    /// produces packages that are rejected on sight, it is reaching no more of the library than
    /// the blind one and is worthless. That distinction has to be visible to be checked.
    /// </summary>
    public static string Run(string target, ReadOnlySpan<byte> data)
    {
        switch (target.ToLowerInvariant())
        {
            case Workbook:
                return Describe(WorkbookPipeline.Run(Workbook, data.ToArray(), out var blindRejection), blindRejection);

            case StructuredWorkbook:
                return Describe(
                    WorkbookPipeline.Run(StructuredWorkbook, WorkbookPackageGenerator.Generate(new FuzzBytes(data)), out var generatedRejection),
                    generatedRejection);

            case Formula:
                RunFormula(data);
                return "evaluated";

            case Address:
                RunAddress(data);
                return "checked";

            default:
                throw new ArgumentException($"Unknown fuzz target '{target}'.", nameof(target));
        }
    }

    private static string Describe(WorkbookOutcome outcome, string? rejection)
    {
        return rejection is null ? outcome.ToString() : $"{outcome} ({rejection})";
    }

    /// <summary>
    /// Evaluate a formula against a sheet holding a value of every type.
    ///
    /// The previous version of this target evaluated against a freshly constructed, empty
    /// workbook, so every reference resolved to blank and the coercion, error-propagation and
    /// range-argument paths were nearly unreachable. A dozen lines of fixture opens all of them.
    /// </summary>
    private static void RunFormula(ReadOnlySpan<byte> data)
    {
        var formula = Encoding.UTF8.GetString(data);

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");
        sheet.Cell("A1").Value = 42;
        sheet.Cell("A2").Value = -1.5;
        sheet.Cell("A3").Value = "text";
        sheet.Cell("A4").Value = true;
        sheet.Cell("A5").Value = new DateTime(2020, 1, 31, 0, 0, 0, DateTimeKind.Unspecified);
        sheet.Cell("A6").Value = XLError.DivisionByZero;
        // A7 is deliberately left blank: an empty cell is a distinct argument kind.
        sheet.Cell("B1").Value = 1;
        sheet.Cell("B2").Value = 2;
        sheet.Cell("B3").Value = 3;

        try
        {
            _ = sheet.Evaluate(formula);
        }
        catch (ExpressionParseException)
        {
            // Not a formula. The corpus is mostly not formulas; that is expected.
        }
        catch (ArgumentException)
        {
            // A function called with arguments it cannot accept.
        }
    }

    /// <summary>
    /// Check what the address predicates <em>claim</em>, not merely that they return.
    ///
    /// The previous version called all three and discarded every result with <c>_ =</c>, so it
    /// could detect a hard crash and nothing else — a predicate answering the wrong question was
    /// invisible to it. Two properties are checked here instead.
    /// </summary>
    private static void RunAddress(ReadOnlySpan<byte> data)
    {
        var address = Encoding.UTF8.GetString(data);

        var isA1 = XLHelper.IsValidA1Address(address);
        var isRc = XLHelper.IsValidRCAddress(address);
        var isRange = XLHelper.IsValidRangeAddress(address);

        // Property 1 — consistency. A single-cell A1 address is also a valid range address;
        // a range of one cell is what it is. If A1 accepts it and range rejects it, one of the
        // two is wrong and only comparing them can say so.
        if (isA1 && !isRange)
            throw new FuzzAssertionException($"'{address}' is a valid A1 address but not a valid range address.");

        if (!isA1)
        {
            _ = isRc;
            return;
        }

        // Property 2 — round trip. An address XLibur accepts must survive being rendered back to
        // text and parsed again. This is the check that can see a wrong answer rather than a
        // crash, and address parsing is load-bearing for the formula shifter.
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");

        IXLAddress parsed;
        try
        {
            parsed = sheet.Cell(address).Address;
        }
        catch (ArgumentException)
        {
            throw new FuzzAssertionException($"'{address}' passed IsValidA1Address but could not be used as a cell address.");
        }

        var rendered = parsed.ToStringRelative();
        if (!XLHelper.IsValidA1Address(rendered))
            throw new FuzzAssertionException($"'{address}' rendered as '{rendered}', which IsValidA1Address rejects.");

        var reparsed = sheet.Cell(rendered).Address;
        if (reparsed.RowNumber != parsed.RowNumber ||
            !reparsed.ColumnLetter.Equals(parsed.ColumnLetter, StringComparison.Ordinal))
        {
            throw new FuzzAssertionException(
                $"'{address}' parsed to {parsed.ColumnLetter}{parsed.RowNumber}, rendered as '{rendered}', " +
                $"which parsed to {reparsed.ColumnLetter}{reparsed.RowNumber}.");
        }
    }
}

/// <summary>
/// Raised when a target's own property is violated. Distinct from anything the library throws,
/// so a violated property can never be mistaken for a library exception the oracle tolerates.
/// </summary>
internal sealed class FuzzAssertionException : Exception
{
    public FuzzAssertionException(string message)
        : base(message)
    {
    }
}
