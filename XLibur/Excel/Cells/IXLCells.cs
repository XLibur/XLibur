using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLCells : IEnumerable<IXLCell>
{
#pragma warning disable S2376 // Write-only properties: intentional batch-set on collection items
    /// <summary>
    /// Sets the cells' value.
    /// <para>
    /// Setter will clear a formula if the cell contains a formula.
    /// If the value is a text that starts with a single quote, setter will prefix the value with a single quote through
    /// <see cref="IXLStyle.IncludeQuotePrefix"/> in Excel too and the value of cell is set to to non-quoted text.
    /// </para>
    /// </summary>
    XLCellValue Value { set; }
#pragma warning restore S2376

    /// <summary>
    /// Clears the contents of these cells.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    IXLCells Clear(XLClearOptions clearOptions = XLClearOptions.All);

    /// <summary>
    /// Delete the comments of these cells.
    /// </summary>
    void DeleteComments();

    /// <summary>
    /// Delete the sparklines of these cells.
    /// </summary>
    void DeleteSparklines();

#pragma warning disable S2376 // Write-only properties: intentional batch-set on collection items
    /// <summary>
    /// Sets the cells' formula with A1 references.
    /// </summary>
    /// <value>The formula with A1 references.</value>
    string FormulaA1 { set; }
#pragma warning restore S2376

#pragma warning disable S2376 // Write-only properties: intentional batch-set on collection items
    /// <summary>
    /// Sets the cells' formula with R1C1 references.
    /// </summary>
    /// <value>The formula with R1C1 references.</value>
    string FormulaR1C1 { set; }
#pragma warning restore S2376

    IXLStyle Style { get; set; }

    void Select();
}
