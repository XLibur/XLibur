using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A slicer: the button panel Excel draws on a worksheet to filter a pivot table or a table.
/// </summary>
/// <remarks>
/// <para>
/// A slicer is owned by the worksheet it is drawn on — see <see cref="IXLWorksheet.Slicers"/>. What
/// it filters is a separate relationship, held by its cache: <see cref="IXLPivotTable.Slicers"/> is
/// a view over the slicers whose cache lists that pivot table, and one cache may list several.
/// </para>
/// <para>
/// A slicer read from a file is reported here in full, including styling XLibur has no model for.
/// Editing one patches the change into the part it was read from rather than regenerating it, so
/// everything alongside the edited attribute survives; a slicer nobody assigns to is not written to
/// at all. The selection is read-only for now — changing it has to move the pivot table's item
/// visibility with it, and that is not modelled yet.
/// </para>
/// </remarks>
public interface IXLSlicer
{
    /// <summary>
    /// The slicer's internal name, unique within the workbook. This is what the drawing anchor
    /// refers to, not what the user sees — see <see cref="Caption"/> for that.
    /// </summary>
    string Name { get; }

    /// <summary>
    /// The heading shown above the slicer's buttons. Defaults to <see cref="Name"/> when the file
    /// does not say otherwise.
    /// </summary>
    string Caption { get; set; }

    /// <summary>
    /// Whether <see cref="Caption"/> is displayed. <c>true</c> unless the file says otherwise.
    /// </summary>
    bool ShowCaption { get; set; }

    /// <summary>
    /// The name of the slicer style, for example <c>SlicerStyleDark3</c>. <c>null</c> means the
    /// workbook default.
    /// </summary>
    /// <remarks>
    /// Deliberately a string rather than an enumeration of the built-in styles: a workbook may name
    /// a custom style, and a read model that could only report the styles it knows about would
    /// silently lose the rest.
    /// </remarks>
    string? Style { get; set; }

    /// <summary>
    /// How many columns of buttons the slicer is laid out in. 1 unless the file says otherwise.
    /// </summary>
    uint ColumnCount { get; set; }

    /// <summary>
    /// The height of one button row, in points. <c>null</c> when the file does not say, which
    /// leaves it to Excel.
    /// </summary>
    double? RowHeightPt { get; set; }

    /// <summary>
    /// The cell the slicer's top-left corner is anchored to. Setting it moves the slicer.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A slicer is a drawing, so it is anchored to the grid the same way a picture or a chart is,
    /// and it moves when rows or columns are inserted above or to the left of it.
    /// </para>
    /// <para>
    /// Moving a slicer read from a file shifts both of its corners together, so it keeps the size
    /// it had. Reading reports the cell the corner sits in; a file may also place the corner some
    /// distance into that cell, and that offset is preserved through a save but is not reported
    /// here.
    /// </para>
    /// </remarks>
    IXLCell Position { get; set; }

    /// <summary>
    /// Whether this slicer filters pivot tables or a table.
    /// </summary>
    XLSlicerSourceKind SourceKind { get; }

    /// <summary>
    /// The name of the field the slicer filters on — a pivot cache field for a pivot slicer, a
    /// table column for a table slicer.
    /// </summary>
    string SourceFieldName { get; }

    /// <summary>
    /// The worksheet the slicer is drawn on.
    /// </summary>
    IXLWorksheet Worksheet { get; }

    /// <summary>
    /// The pivot tables this slicer filters. Empty when <see cref="SourceKind"/> is
    /// <see cref="XLSlicerSourceKind.Table"/>.
    /// </summary>
    /// <remarks>
    /// More than one pivot table may share a slicer cache, which is how a dashboard drives several
    /// pivot tables from one set of buttons. A pivot table listed in the cache but missing from the
    /// workbook is omitted rather than reported as null.
    /// </remarks>
    IReadOnlyList<IXLPivotTable> PivotTables { get; }

    /// <summary>
    /// The table this slicer filters, or <c>null</c> when <see cref="SourceKind"/> is
    /// <see cref="XLSlicerSourceKind.PivotTable"/>.
    /// </summary>
    IXLTable? Table { get; }

    /// <summary>
    /// Whether the slicer records an explicit selection. <c>false</c> means every item is showing,
    /// which is how Excel represents a slicer nobody has clicked.
    /// </summary>
    bool HasSelection { get; }

    /// <summary>
    /// The items currently selected in the slicer. Empty when <see cref="HasSelection"/> is
    /// <c>false</c>.
    /// </summary>
    /// <remarks>
    /// The two source kinds keep their selection in different places, and both are read here. A
    /// pivot slicer's cache holds indices into the pivot cache field's shared items; a table
    /// slicer's selection is the value filter on the bound column of the table's auto filter. A
    /// table column filtered by something other than a list of values — a custom or top-ten filter,
    /// which a slicer cannot produce but a user can apply by hand — reports no selection.
    /// </remarks>
    IReadOnlyList<XLCellValue> SelectedItems { get; }
}
