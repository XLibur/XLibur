using XLibur.Excel;

namespace XLibur.Report.Tags;

/// <summary>
/// Turns on Excel's autofilter over the generated rows. Written anywhere in the options row as
/// <c>&lt;&lt;AutoFilter&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// The filter covers the row above the generated block as well, because that is where a template
/// puts its column headings and a filter without headings is not much use. Give
/// <c>noheader</c> to filter the generated rows alone.
/// </remarks>
public sealed class AutoFilterTag : OptionTag
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
/// Hides the column the tag is written in — for a column a template needs in order to sort or
/// total, but that the reader should not see. Written as <c>&lt;&lt;Hidden&gt;&gt;</c>.
/// </summary>
public sealed class HiddenTag : OptionTag
{
    /// <inheritdoc />
    public override void Execute(ProcessingContext context) => context.Worksheet.Column(Column).Hide();
}

/// <summary>
/// Removes the column the tag is written in, once everything else has run. Written as
/// <c>&lt;&lt;Delete&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// Runs last, so a column may be sorted or totalled by and then removed. Give
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

        context.Worksheet.Column(Column).Delete();
    }
}
