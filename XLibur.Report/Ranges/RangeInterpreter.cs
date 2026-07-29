using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Report.Expressions;
using XLibur.Report.Rewriting;

namespace XLibur.Report.Ranges;

/// <summary>
/// Drives one generation run: resolves the bound ranges, evaluates everything outside them, then
/// expands each range.
/// </summary>
/// <remarks>
/// Order matters. Cells outside a bound range see only the workbook-wide variables, so they are
/// evaluated first, while the sheet still has its template addresses. Expansion then moves them,
/// which is harmless because their values are already resolved. Cells <em>inside</em> a bound
/// range are skipped by that first pass — they refer to <c>item</c>, which only exists once
/// expansion puts a row's data in scope.
/// </remarks>
internal sealed class RangeInterpreter
{
    private readonly CellEvaluator _evaluator;
    private readonly IExpressionEngine _engine;
    private readonly TemplateErrors _errors;

    public RangeInterpreter(IExpressionEngine engine, TemplateErrors errors)
    {
        _evaluator = new CellEvaluator(engine, errors);
        _engine = engine;
        _errors = errors;
    }

    /// <summary>Expansions performed by the last <see cref="Evaluate"/> call, in the order they ran.</summary>
    public List<ExpansionRecord> Expansions { get; } = new();

    public void Evaluate(IXLWorkbook workbook, IReadOnlyDictionary<string, object?> variables)
    {
        Expansions.Clear();

        var globalScope = new ExpressionScope(variables);
        var bound = RangeBinder.Resolve(workbook, variables, _errors);

        EvaluateOutsideBoundRanges(workbook, bound, globalScope);

        var pendingPivots = new List<PivotRequest>();
        var expander = new RangeExpander(_evaluator, _engine, _errors, pendingPivots);
        foreach (var range in bound)
        {
            var record = expander.Expand(range, globalScope);
            if (record is not null)
            {
                Expansions.Add(record);
            }
        }

        // Once for the workbook, and only now: a chart or a pivot may read a range on another sheet, so
        // nothing can be re-pointed until every sheet has settled.
        ReferenceRewriter.Rewrite(workbook, Expansions);
        PivotRewriter.Rewrite(workbook, Expansions, _errors);

        // Last of all. A <<Pivot>> is built from live ranges that have finished moving, so waiting
        // costs nothing; and being built after the rewriter is what keeps the rewriter from re-pointing
        // a source that is already where it should be, and lets the new pivot share a cache with a
        // template pivot that has been re-pointed to the same rows.
        PivotBuilder.Build(pendingPivots, _errors);
    }

    private void EvaluateOutsideBoundRanges(IXLWorkbook workbook, List<BoundRange> bound, ExpressionScope globalScope)
    {
        foreach (var worksheet in workbook.Worksheets.Where(IsGenerated).ToList())
        {
            var areas = bound
                .Where(b => ReferenceEquals(b.Worksheet, worksheet))
                .Select(b => b.Area)
                .ToList();

            foreach (var cell in worksheet.CellsUsed(XLCellsUsedOptions.Contents).ToList())
            {
                var address = cell.Address;
                if (areas.Any(a => a.Contains(address.RowNumber, address.ColumnNumber)))
                {
                    continue;
                }

                _evaluator.Evaluate(cell, globalScope);
            }
        }
    }

    internal static bool IsGenerated(IXLWorksheet worksheet) => worksheet.Visibility == XLWorksheetVisibility.Visible;
}
