#nullable disable

using System;

namespace XLibur.Excel;

public interface IXLPivotTable
{
    XLPivotTableTheme Theme { get; set; }

    IXLPivotFields ReportFilters { get; }

    /// <summary>
    /// Labels displayed in columns (i.e. horizontal axis) of the pivot table.
    /// </summary>
    IXLPivotFields ColumnLabels { get; }

    /// <summary>
    /// Labels displayed in rows (i.e. vertical axis) of the pivot table.
    /// </summary>
    IXLPivotFields RowLabels { get; }
    IXLPivotValues Values { get; }

    string Name { get; set; }
    string Title { get; set; }
    string Description { get; set; }

    string ColumnHeaderCaption { get; set; }
    string RowHeaderCaption { get; set; }

    /// <summary>
    /// Top left corner cell of a pivot table. If the pivot table contains filters fields, the target cell is top
    /// left cell of the first filter field.
    /// </summary>
    IXLCell TargetCell { get; set; }

    /// <summary>
    /// The cache of data for the pivot table. The pivot table is created
    /// from cached data, not up-to-date data in a worksheet.
    /// </summary>
    IXLPivotCache PivotCache { get; set; }

    bool MergeAndCenterWithLabels { get; set; } // MergeItem
    int RowLabelIndent { get; set; } // Indent

    /// <summary>
    /// Filter fields layout setting that indicates layout order of filter fields. The layout
    /// uses <see cref="FilterFieldsPageWrap"/> to determine when to break to a new row or
    /// column. Default value is <see cref="XLFilterAreaOrder.DownThenOver"/>.
    /// </summary>
    XLFilterAreaOrder FilterAreaOrder { get; set; }

    /// <summary>
    /// Specifies the number of page fields to display before starting another row or column.
    /// Value = 0 means unlimited.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">If value &lt; 0.</exception>
    int FilterFieldsPageWrap { get; set; } // PageWrap
    string ErrorValueReplacement { get; set; } // ErrorCaption
    string EmptyCellReplacement { get; set; } // MissingCaption
    bool AutofitColumns { get; set; } //UseAutoFormatting
    bool PreserveCellFormatting { get; set; } // PreserveFormatting

    bool ShowGrandTotalsRows { get; set; } // RowGrandTotals
    bool ShowGrandTotalsColumns { get; set; } // ColumnGrandTotals
    bool FilteredItemsInSubtotals { get; set; } // Subtotal filtered page items
    bool AllowMultipleFilters { get; set; } // MultipleFieldFilters
    bool UseCustomListsForSorting { get; set; } // CustomListSort

    bool ShowExpandCollapseButtons { get; set; }
    bool ShowContextualTooltips { get; set; }
    bool ShowPropertiesInTooltips { get; set; }
    bool DisplayCaptionsAndDropdowns { get; set; }
    bool ClassicPivotTableLayout { get; set; }
    bool ShowValuesRow { get; set; }
    bool ShowEmptyItemsOnRows { get; set; }
    bool ShowEmptyItemsOnColumns { get; set; }
    bool DisplayItemLabels { get; set; }
    bool SortFieldsAtoZ { get; set; }

    bool PrintExpandCollapsedButtons { get; set; }
    bool RepeatRowLabels { get; set; }
    bool PrintTitles { get; set; }

    bool EnableShowDetails { get; set; }
    bool EnableCellEditing { get; set; }

    IXLPivotTable CopyTo(IXLCell targetCell);

    IXLPivotTable SetName(string value);

    IXLPivotTable SetTitle(string value);

    IXLPivotTable SetDescription(string value);

    IXLPivotTable SetMergeAndCenterWithLabels(); IXLPivotTable SetMergeAndCenterWithLabels(bool value);

    IXLPivotTable SetRowLabelIndent(int value);

    IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value);

    IXLPivotTable SetFilterFieldsPageWrap(int value);

    IXLPivotTable SetErrorValueReplacement(string value);

    IXLPivotTable SetEmptyCellReplacement(string value);

    IXLPivotTable SetAutofitColumns(); IXLPivotTable SetAutofitColumns(bool value);

    IXLPivotTable SetPreserveCellFormatting(); IXLPivotTable SetPreserveCellFormatting(bool value);

    IXLPivotTable SetShowGrandTotalsRows(); IXLPivotTable SetShowGrandTotalsRows(bool value);

    /// <summary>
    /// Should pivot table display a grand total for each row in the last column of a pivot
    /// table (it will enlarge pivot table for extra column).
    /// </summary>
    /// <remarks>
    /// This API has inverse row/column names than the Excel. Excel: <em>On for rows
    /// </em> should use this method <em>ShowGrandTotalsColumns</em>.
    /// </remarks>
    IXLPivotTable SetShowGrandTotalsColumns(); IXLPivotTable SetShowGrandTotalsColumns(bool value);

    IXLPivotTable SetFilteredItemsInSubtotals(); IXLPivotTable SetFilteredItemsInSubtotals(bool value);

    IXLPivotTable SetAllowMultipleFilters(); IXLPivotTable SetAllowMultipleFilters(bool value);

    IXLPivotTable SetUseCustomListsForSorting(); IXLPivotTable SetUseCustomListsForSorting(bool value);

    IXLPivotTable SetShowExpandCollapseButtons(); IXLPivotTable SetShowExpandCollapseButtons(bool value);

    IXLPivotTable SetShowContextualTooltips(); IXLPivotTable SetShowContextualTooltips(bool value);

    IXLPivotTable SetShowPropertiesInTooltips(); IXLPivotTable SetShowPropertiesInTooltips(bool value);

    IXLPivotTable SetDisplayCaptionsAndDropdowns(); IXLPivotTable SetDisplayCaptionsAndDropdowns(bool value);

    IXLPivotTable SetClassicPivotTableLayout(); IXLPivotTable SetClassicPivotTableLayout(bool value);

    IXLPivotTable SetShowValuesRow(); IXLPivotTable SetShowValuesRow(bool value);

    IXLPivotTable SetShowEmptyItemsOnRows(); IXLPivotTable SetShowEmptyItemsOnRows(bool value);

    IXLPivotTable SetShowEmptyItemsOnColumns(); IXLPivotTable SetShowEmptyItemsOnColumns(bool value);

    IXLPivotTable SetDisplayItemLabels(); IXLPivotTable SetDisplayItemLabels(bool value);

    IXLPivotTable SetSortFieldsAtoZ(); IXLPivotTable SetSortFieldsAtoZ(bool value);

    IXLPivotTable SetPrintExpandCollapsedButtons(); IXLPivotTable SetPrintExpandCollapsedButtons(bool value);

    IXLPivotTable SetRepeatRowLabels(); IXLPivotTable SetRepeatRowLabels(bool value);

    IXLPivotTable SetPrintTitles(); IXLPivotTable SetPrintTitles(bool value);


    IXLPivotTable SetEnableShowDetails(); IXLPivotTable SetEnableShowDetails(bool value);



    IXLPivotTable SetEnableCellEditing(); IXLPivotTable SetEnableCellEditing(bool value);

    IXLPivotTable SetColumnHeaderCaption(string value);

    IXLPivotTable SetRowHeaderCaption(string value);

    bool ShowRowHeaders { get; set; }
    bool ShowColumnHeaders { get; set; }
    bool ShowRowStripes { get; set; }
    bool ShowColumnStripes { get; set; }
    XLPivotSubtotals Subtotals { get; set; }

    /// <summary>
    /// Set the layout of the pivot table. It also changes layout of all pivot fields.
    /// </summary>
    XLPivotLayout Layout { set; }
    bool InsertBlankLines { set; }

    IXLPivotTable SetShowRowHeaders(); IXLPivotTable SetShowRowHeaders(bool value);

    IXLPivotTable SetShowColumnHeaders(); IXLPivotTable SetShowColumnHeaders(bool value);

    IXLPivotTable SetShowRowStripes(); IXLPivotTable SetShowRowStripes(bool value);

    IXLPivotTable SetShowColumnStripes(); IXLPivotTable SetShowColumnStripes(bool value);

    IXLPivotTable SetSubtotals(XLPivotSubtotals value);

    IXLPivotTable SetLayout(XLPivotLayout value);

    IXLPivotTable SetInsertBlankLines(); IXLPivotTable SetInsertBlankLines(bool value);

    IXLWorksheet Worksheet { get; }

    IXLPivotTableStyleFormats StyleFormats { get; }
}
