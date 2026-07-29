using XLibur.Excel;

namespace XLibur.Report.Expressions;

/// <summary>
/// Whether the value an expression produced counts as true.
/// </summary>
/// <remarks>
/// <para>
/// Tags that ask a question of the data — <c>&lt;&lt;If test=…&gt;&gt;</c>, <c>&lt;&lt;Delete
/// keep=…&gt;&gt;</c> — get an <see cref="object"/> back from whichever engine the template uses, so
/// the rule for reading it as a yes or a no belongs here rather than in an engine. Both engines then
/// answer the same question the same way, which is the point of the seam.
/// </para>
/// <para>
/// The rule is Scriban's: only <c>null</c> and <c>false</c> are false. <strong>Zero and the empty
/// string are true</strong>, which surprises people, so a template that means "greater than nothing"
/// has to say so: <c>test="item.Quantity &gt; 0"</c>, not <c>test="item.Quantity"</c>.
/// </para>
/// </remarks>
internal static class ExpressionTruth
{
    public static bool IsTruthy(object? value) => value switch
    {
        null => false,
        bool logical => logical,
        Blank => false,
        XLError => false,
        XLCellValue cellValue => IsTruthy(Unwrap(cellValue)),
        _ => true,
    };

    /// <summary>
    /// Reads a cell value back out as a plain object, so a value that arrived through a cell is
    /// judged the same way one that came straight from an expression is.
    /// </summary>
    private static object? Unwrap(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => null,
        XLDataType.Boolean => value.GetBoolean(),
        XLDataType.Error => value.GetError(),
        XLDataType.Text => value.GetText(),
        _ => value.ToString(),
    };
}
