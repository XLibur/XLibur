using XLibur.Excel;

namespace XLibur.Report.Tags;

/// <summary>
/// Declares that a range repeats <em>across</em> rather than downwards: one column per item instead
/// of one row. Written in the range's last column as <c>&lt;&lt;Horizontal&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// <para>
/// A horizontal range is the mirror of a vertical one. Its last <em>column</em> is the options
/// column, the columns before it are the ones repeated, and the rows are its lines: a tag sits in a
/// row and acts on that row, a summary totals across the generated columns, and the labels a reader
/// reads down the side sit in a column outside the range. So a template with the range B4:C6 and
/// <c>&lt;&lt;Horizontal&gt;&gt;</c> in C4 repeats column B once per item and keeps C for its
/// options.
/// </para>
/// <para>
/// The tag has to be found before the range can be read at all, which is why it is looked for in the
/// last column specifically rather than anywhere in the range — that is the one place that means the
/// same thing whichever way the range turns out to run.
/// </para>
/// <para>
/// Grouping is not supported across: <c>&lt;&lt;Group&gt;&gt;</c> in a horizontal range reports an
/// error rather than doing something surprising.
/// </para>
/// </remarks>
public sealed class HorizontalTag : OptionTag
{
}

/// <summary>
/// Turns on Excel's autofilter over the generated rows. Written anywhere in the options row as
/// <c>&lt;&lt;AutoFilter&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// The filter covers the row above the generated block as well, because that is where a template
/// puts its column headings and a filter without headings is not much use. Give
/// <c>noheader</c> to filter the generated rows alone.
/// <para>
/// Excel filters rows, so the tag does nothing in a range that repeats across.
/// </para>
/// </remarks>
public sealed class AutoFilterTag : OptionTag
{
    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        if (context.IsHorizontal)
        {
            context.Errors.Add(new TemplateError(
                "<<AutoFilter>> has no meaning in a range that repeats across: Excel filters rows, not columns.",
                context.Worksheet.Name));
            return;
        }

        var address = context.GeneratedRange.RangeAddress;
        var firstRow = address.FirstAddress.RowNumber;
        var lastRow = address.LastAddress.RowNumber;

        if (lastRow < firstRow)
        {
            return;
        }

        if (!Token.Flag("noheader") && firstRow > 1)
        {
            firstRow--;
        }

        context.Worksheet
            .Range(firstRow, address.FirstAddress.ColumnNumber, lastRow, address.LastAddress.ColumnNumber)
            .SetAutoFilter();
    }
}

/// <summary>
/// Widens the range's columns to fit what was generated into them. Written in the options row as
/// <c>&lt;&lt;ColsFit&gt;&gt;</c>.
/// </summary>
public sealed class ColumnsFitTag : OptionTag
{
    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        var address = context.GeneratedRange.RangeAddress;

        context.Worksheet
            .Columns(address.FirstAddress.ColumnNumber, address.LastAddress.ColumnNumber)
            .AdjustToContents();
    }
}

/// <summary>
/// Sets the generated rows' heights to fit their contents. Written in the options row as
/// <c>&lt;&lt;RowsFit&gt;&gt;</c>.
/// </summary>
public sealed class RowsFitTag : OptionTag
{
    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        var address = context.GeneratedRange.RangeAddress;
        var firstRow = address.FirstAddress.RowNumber;
        var lastRow = address.LastAddress.RowNumber;

        if (lastRow < firstRow)
        {
            return;
        }

        context.Worksheet.Rows(firstRow, lastRow).AdjustToContents();
    }
}

/// <summary>
/// Hides the line the tag is written in — the column for a range that repeats downwards, the row for
/// one that repeats across. For a line a template needs in order to sort or total, but that the
/// reader should not see. Written as <c>&lt;&lt;Hidden&gt;&gt;</c>.
/// </summary>
public sealed class HiddenTag : OptionTag
{
    /// <inheritdoc />
    public override void Execute(ProcessingContext context) => context.Axis.HideLine(context.Worksheet, Line);
}

/// <summary>
/// Removes the line the tag is written in, once everything else has run. Written as
/// <c>&lt;&lt;Delete&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// Runs last, so a line may be sorted or totalled by and then removed. Give
/// <c>keep</c> with a truthy value to leave it in place — which is how a template makes the
/// removal conditional: <c>&lt;&lt;Delete keep="{{ ShowWorkings }}"&gt;&gt;</c>. Bare
/// <c>&lt;&lt;Delete keep&gt;&gt;</c> keeps it outright.
/// </remarks>
public sealed class DeleteTag : OptionTag
{
    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        // Present-but-empty is the bare flag form, which means yes.
        if (Token.Has("keep") && context.IsTrue(Token.Value("keep", "true")))
        {
            return;
        }

        context.Axis.DeleteLine(context.Worksheet, Line);
    }
}
