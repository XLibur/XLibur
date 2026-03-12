using System;

namespace XLibur.Excel;

public interface IXLStyle : IEquatable<IXLStyle>
{
    IXLAlignment Alignment { get; set; }

    IXLBorder Border { get; set; }

    IXLNumberFormat DateFormat { get; }

    IXLFill Fill { get; set; }

    IXLFont Font { get; set; }

    /// <summary>
    /// Should the text values of a cell saved to the file be prefixed by a quote (<c>'</c>) character?
    /// Has no effect if cell values is not a <see cref="XLDataType.Text"/>. Doesn't affect values during runtime,
    /// text values are returned without quote.
    /// </summary>
    bool IncludeQuotePrefix { get; set; }

    IXLNumberFormat NumberFormat { get; set; }

    IXLProtection Protection { get; set; }

    IXLStyle SetIncludeQuotePrefix(bool includeQuotePrefix = true);

    /// <summary>
    /// Apply multiple style changes as a single operation. For cell containers, only one repository
    /// lookup and one style-slice write occurs regardless of how many properties change.
    /// </summary>
    /// <param name="modifications">Action that receives an <see cref="IXLStyle"/> and sets desired properties.</param>
    /// <returns>This style instance for fluent chaining.</returns>
    IXLStyle Batch(Action<IXLStyle> modifications);
}
