using XLibur.Excel;

namespace XLibur.Report.Tags;

/// <summary>
/// A tag that can summarise a run of generated rows into a single cell.
/// </summary>
/// <remarks>
/// The options row is not the only place a summary lands: once a range is grouped, every group's
/// subtotal row repeats whatever summaries the options row declares, over that group's rows alone.
/// Both go through this one method, so a summary is written the same way wherever it appears — and
/// a custom tag that implements it joins group subtotals without further wiring.
/// </remarks>
public interface IRangeSummaryTag
{
    /// <summary>
    /// Writes a summary of rows <paramref name="firstRow"/> to <paramref name="lastRow"/> into
    /// <paramref name="target"/>, and reports whether it could.
    /// </summary>
    /// <remarks>
    /// A run with no rows in it — <paramref name="lastRow"/> below <paramref name="firstRow"/> — is
    /// not an error: an empty range still has a total, and it is zero.
    /// </remarks>
    bool TryWriteSummary(IXLCell target, int firstRow, int lastRow, ProcessingContext context);
}
