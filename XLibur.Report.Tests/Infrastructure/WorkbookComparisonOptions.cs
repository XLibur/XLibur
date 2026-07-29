namespace XLibur.Report.Tests.Infrastructure;

/// <summary>
/// Selects which dimensions <see cref="WorkbookComparer"/> checks.
/// </summary>
/// <remarks>
/// Everything is on by default: a golden-file test is worth most when it notices changes nobody
/// was looking for. Individual dimensions can be switched off for a fixture that legitimately
/// varies in one of them.
/// </remarks>
public sealed class WorkbookComparisonOptions
{
    /// <summary>All dimensions.</summary>
    public static WorkbookComparisonOptions Default => new();

    /// <summary>Cell values and their types.</summary>
    public bool Values { get; set; } = true;

    /// <summary>Cell formulas, compared in A1 notation.</summary>
    public bool Formulas { get; set; } = true;

    /// <summary>Cell styles, compared through their string form.</summary>
    public bool Styles { get; set; } = true;

    /// <summary>Merged cell ranges.</summary>
    public bool MergedRanges { get; set; } = true;

    /// <summary>
    /// Conditional formatting rules — their count, applied ranges and types. The count matters
    /// as much as the ranges: duplicating a rule per generated cell is the upstream behaviour
    /// this library deliberately does not reproduce.
    /// </summary>
    public bool ConditionalFormats { get; set; } = true;

    /// <summary>Cell comments.</summary>
    public bool Comments { get; set; } = true;

    /// <summary>Cell hyperlinks.</summary>
    public bool Hyperlinks { get; set; } = true;

    /// <summary>Row heights, column widths and outline levels.</summary>
    public bool Dimensions { get; set; } = true;

    /// <summary>Print areas, page breaks and autofilter state.</summary>
    public bool PageSetup { get; set; } = true;

    /// <summary>How many differences to collect before giving up on a comparison.</summary>
    public int MaxDifferences { get; set; } = 50;
}
