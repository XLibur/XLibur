using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Report.Rewriting;

namespace XLibur.Report.Tags;

/// <summary>
/// Builds a pivot table over the rows the range generated. Written in the options row as
/// <c>&lt;&lt;Pivot dest="Summary!A3"&gt;&gt;</c>, with
/// <see cref="PivotFieldTag">field tags</see> in the options row saying what each column is for.
/// </summary>
/// <remarks>
/// <para>
/// The static pattern — a pivot authored in the template over a named range, which the engine
/// re-points and refreshes — is the primary one, and the one to reach for. This tag is for the other
/// case: a pivot whose shape is decided by the template rather than drawn by hand, so that one
/// template can produce a pivot over whatever the data turns out to be.
/// </para>
/// <para>
/// The source is the generated block <em>plus the row above it</em>, because a pivot names its fields
/// from a heading row and a bound range's headings sit above it — the same convention
/// <see cref="AutoFilterTag"/> follows. A range starting in row 1 has no heading row and reports an
/// error rather than inventing field names.
/// </para>
/// <para>
/// Parameters: <c>dest</c> (required — where to put it, as a sheet-qualified cell reference or the
/// name of a defined name covering one cell), <c>name</c> (the pivot table's name, defaulting to
/// <c>Pivot</c> and a number).
/// </para>
/// <para>
/// <c>dest</c> is read in template coordinates, like everything else a template writes: a cell named
/// as being below the range is below the <em>generated</em> range in the report, however many rows
/// that turned out to be. That works by holding the destination as a one-cell range from before the
/// range expands and letting the core library carry it — the same mechanism that moves a picture, and
/// the reason a picture needed no code. The pivot itself is built once every range has been
/// generated, because a pivot table's position is a plain rectangle that would not follow anything
/// that happened after it was created.
/// </para>
/// <para>
/// Not supported, with an error rather than a surprise: a range that repeats across (a pivot names its
/// fields from a heading row, which such a range does not have), and a grouped range (its subtotal
/// rows are part of the generated block, and a pivot would count them as data on top of the rows they
/// already total).
/// </para>
/// </remarks>
public sealed class PivotTag : OptionTag
{
    private IXLRange? _destination;
    private List<PivotField> _fields = new();

    /// <summary>
    /// Fixes what the tag will need, while the addresses the template wrote still mean what they said.
    /// </summary>
    /// <remarks>
    /// Nothing is reported here. Expansion has not run, so a template mistake found now would be
    /// reported ahead of every other tag's, and the order errors are listed in is the order a report
    /// author reads them in.
    /// </remarks>
    public override IReadOnlyList<object?> TransformItems(IReadOnlyList<object?> items, ProcessingContext context)
    {
        var headingRow = context.GeneratedRange.RangeAddress.FirstAddress.RowNumber - 1;

        _destination = Resolve(context, Token.Value("dest"));

        if (headingRow >= 1 && !context.IsHorizontal)
        {
            // The heading cell answers both of a field tag's questions — which column, and what the
            // field is called — so holding it live is enough to survive a column being deleted out
            // from under the tag, or moved.
            _fields = context.Tags
                .OfType<PivotFieldTag>()
                .Where(tag => !tag.InRepeatedRow)
                .OrderBy(tag => tag.Line)
                .Select(tag => new PivotField(
                    tag,
                    context.Worksheet.Range(headingRow, tag.Line, headingRow, tag.Line)))
                .ToList();
        }

        return items;
    }

    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        if (context.IsHorizontal)
        {
            Report(context, "<<Pivot>> is not supported in a range that repeats across: a pivot table "
                + "names its fields from a heading row, which such a range does not have.");
            return;
        }

        if (context.Tags.Any(tag => tag is GroupTag))
        {
            Report(context, "<<Pivot>> cannot be used in a grouped range: the subtotal rows "
                + "<<Group>> writes are part of the generated block, and the pivot would count them as "
                + "data alongside the rows they total. Pivot the range or group it, not both.");
            return;
        }

        var source = Source(context);
        if (source is null)
        {
            return;
        }

        if (_destination is null)
        {
            var dest = Token.Value("dest");

            Report(context, dest.Length == 0
                ? "<<Pivot>> needs a dest, saying where to put the pivot table — "
                  + "<<Pivot dest=\"Summary!A3\">>."
                : $"<<Pivot dest=\"{dest}\">> is not a cell this library can find. Give a "
                  + "sheet-qualified cell reference, or the name of a defined name covering one cell.");
            return;
        }

        if (_fields.Count == 0)
        {
            Report(context, "<<Pivot>> was given no fields to lay out. Put <<Row>>, <<Column>>, "
                + "<<Page>> or <<Data>> in the options row under the columns the pivot should use.");
            return;
        }

        // Built once every range has been generated, not here: see the remarks on this type.
        context.PendingPivots.Add(new PivotRequest(
            source,
            _destination,
            Token.Value("name"),
            _fields,
            context.Worksheet.Name,
            context.Worksheet.Cell(Row, Column).Address.ToString()));
    }

    /// <summary>
    /// The heading row plus the generated rows, held as a live range so that a range generated above
    /// it afterwards carries it. <c>null</c> once the reason there is no source has been reported.
    /// </summary>
    private IXLRange? Source(ProcessingContext context)
    {
        var address = context.GeneratedRange.RangeAddress;
        var firstRow = address.FirstAddress.RowNumber;
        var lastRow = address.LastAddress.RowNumber;

        if (lastRow < firstRow)
        {
            return null;
        }

        if (firstRow == 1)
        {
            Report(context, "<<Pivot>> needs a heading row above the range to name the pivot's fields, "
                + "and the range starts in row 1.");
            return null;
        }

        return context.Worksheet.Range(
            firstRow - 1,
            address.FirstAddress.ColumnNumber,
            lastRow,
            address.LastAddress.ColumnNumber);
    }

    /// <summary>
    /// The one cell <paramref name="dest"/> names, as a live range: a defined name covering a single
    /// cell, or an A1 reference on its own sheet or on the range's.
    /// </summary>
    private static IXLRange? Resolve(ProcessingContext context, string dest)
    {
        if (dest.Length == 0)
        {
            return null;
        }

        return ResolveName(context.Worksheet.Workbook, dest) ?? ResolveReference(context, dest);
    }

    private static IXLRange? ResolveName(IXLWorkbook workbook, string dest)
    {
        if (!workbook.DefinedNames.TryGetValue(dest, out var definedName))
        {
            return null;
        }

        var range = definedName.Ranges.FirstOrDefault();
        if (range is null)
        {
            return null;
        }

        var address = range.RangeAddress;

        return address.FirstAddress.RowNumber == address.LastAddress.RowNumber
            && address.FirstAddress.ColumnNumber == address.LastAddress.ColumnNumber
            ? range
            : null;
    }

    private static IXLRange? ResolveReference(ProcessingContext context, string dest)
    {
        if (!SheetReference.TryParse(dest, out var reference))
        {
            return null;
        }

        var worksheet = context.Worksheet;

        if (reference.SheetName is { Length: > 0 } sheetName)
        {
            // Typed explicitly: the internals grant makes an XLWorksheet overload visible too, and
            // `var` cannot choose between them.
            if (!context.Worksheet.Workbook.TryGetWorksheet(sheetName, out IXLWorksheet? named))
            {
                return null;
            }

            worksheet = named;
        }

        return worksheet.Range(
            reference.FirstRow,
            reference.FirstColumn,
            reference.FirstRow,
            reference.FirstColumn);
    }

    private void Report(ProcessingContext context, string message) =>
        context.Errors.Add(new TemplateError(
            message,
            context.Worksheet.Name,
            context.Worksheet.Cell(Row, Column).Address.ToString()));
}

/// <summary>
/// Says what one of a pivoted range's columns is for. Written in the options row under that column as
/// <c>&lt;&lt;Row&gt;&gt;</c>, <c>&lt;&lt;Column&gt;&gt;</c> (or <c>&lt;&lt;Col&gt;&gt;</c>),
/// <c>&lt;&lt;Page&gt;&gt;</c> or <c>&lt;&lt;Data&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// <para>
/// The field's name is the column's heading — the cell above the range — because that is the name the
/// pivot cache reads out of the source, and a template should not have to write it twice.
/// </para>
/// <para>
/// <c>&lt;&lt;Data&gt;&gt;</c> sums by default. Name another summary as a flag —
/// <c>&lt;&lt;Data avg&gt;&gt;</c>, <c>&lt;&lt;Data count&gt;&gt;</c>, <c>&lt;&lt;Data max&gt;&gt;</c>
/// — or as a value, <c>&lt;&lt;Data func=min&gt;&gt;</c>. The recognised names are <c>sum</c>,
/// <c>count</c>, <c>counta</c>, <c>countnums</c>, <c>avg</c>/<c>average</c>, <c>min</c>, <c>max</c>,
/// <c>product</c>, <c>stddev</c>, <c>stddevp</c>, <c>var</c> and <c>varp</c>. Give <c>title</c> to
/// label the value column something other than its field name.
/// </para>
/// <para>
/// The tag does nothing when it runs; <see cref="PivotTag"/> reads it. Without a
/// <c>&lt;&lt;Pivot&gt;&gt;</c> in the same options row it has no effect at all, which is reported.
/// </para>
/// </remarks>
public sealed class PivotFieldTag : OptionTag
{
    private static readonly Dictionary<string, XLPivotSummary> Summaries =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["sum"] = XLPivotSummary.Sum,
            ["count"] = XLPivotSummary.Count,
            ["counta"] = XLPivotSummary.Count,
            ["countnums"] = XLPivotSummary.CountNumbers,
            ["avg"] = XLPivotSummary.Average,
            ["average"] = XLPivotSummary.Average,
            ["min"] = XLPivotSummary.Minimum,
            ["max"] = XLPivotSummary.Maximum,
            ["product"] = XLPivotSummary.Product,
            ["stddev"] = XLPivotSummary.StandardDeviation,
            ["stdev"] = XLPivotSummary.StandardDeviation,
            ["stddevp"] = XLPivotSummary.PopulationStandardDeviation,
            ["stdevp"] = XLPivotSummary.PopulationStandardDeviation,
            ["var"] = XLPivotSummary.Variance,
            ["varp"] = XLPivotSummary.PopulationVariance,
        };

    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        if (context.Tags.Any(tag => tag is PivotTag))
        {
            return;
        }

        context.Errors.Add(new TemplateError(
            $"<<{Token.Name}>> only means something to a <<Pivot>>, and there is none in this range's "
            + "options row.",
            context.Worksheet.Name,
            context.Worksheet.Cell(Row, Column).Address.ToString()));
    }

    /// <summary>Adds the field this tag describes to <paramref name="pivot"/>.</summary>
    internal void AddTo(IXLPivotTable pivot, string fieldName)
    {
        if (string.Equals(Token.Name, "Data", StringComparison.OrdinalIgnoreCase))
        {
            pivot.Values.Add(fieldName, Token.Value("title", fieldName)).SetSummaryFormula(Summary());
            return;
        }

        var axis = Token.Name.ToUpperInvariant() switch
        {
            "ROW" => pivot.RowLabels,
            "PAGE" => pivot.ReportFilters,
            _ => pivot.ColumnLabels,
        };

        axis.Add(fieldName);
    }

    /// <summary>
    /// Which summary a <c>&lt;&lt;Data&gt;&gt;</c> field uses. An unrecognised name sums, on the
    /// grounds that a total is what a report reader expects of a value column.
    /// </summary>
    private XLPivotSummary Summary()
    {
        if (Summaries.TryGetValue(Token.Value("func"), out var named))
        {
            return named;
        }

        // The first summary named as a flag, in the order the template wrote them, so that
        // <<Data avg>> and <<Data func=avg>> are the same tag written two ways.
        foreach (var parameter in Token.Parameters.Keys)
        {
            if (Summaries.TryGetValue(parameter, out var summary) && Token.Flag(parameter))
            {
                return summary;
            }
        }

        return XLPivotSummary.Sum;
    }
}
