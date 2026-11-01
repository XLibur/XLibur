using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A timeline: the date scrubber Excel draws on a worksheet to filter a pivot table by date.
/// </summary>
/// <remarks>
/// <para>
/// A timeline is owned by the worksheet it is drawn on — see <see cref="IXLWorksheet.Timelines"/>.
/// What it filters is a separate relationship, held by its cache: <see cref="IXLPivotTable.Timelines"/>
/// is a view over the timelines whose cache lists that pivot table.
/// </para>
/// <para>
/// A timeline read from a file is reported here in full, including attributes XLibur has no model
/// for. Editing one patches the change into the part it was read from rather than regenerating it,
/// so everything alongside the edited attribute survives; a timeline nobody assigns to is not
/// written to at all. The selection is read-only — changing it has to move the pivot table's
/// <c>dateBetween</c> filter and item visibility with it, and that is not modelled.
/// </para>
/// </remarks>
public interface IXLTimeline
{
    /// <summary>
    /// The timeline's internal name, unique within the workbook. This is what the drawing anchor
    /// refers to, not what the user sees — see <see cref="Caption"/> for that.
    /// </summary>
    string Name { get; }

    /// <summary>The heading shown above the band. Defaults to <see cref="Name"/>.</summary>
    string Caption { get; set; }

    /// <summary>Whether <see cref="Caption"/> is displayed. <c>true</c> unless the file says otherwise.</summary>
    bool ShowHeader { get; set; }

    /// <summary>Whether the selected range is written out under the header.</summary>
    bool ShowSelectionLabel { get; set; }

    /// <summary>Whether the level chooser (Years / Quarters / Months / Days) is shown.</summary>
    bool ShowTimeLevel { get; set; }

    /// <summary>Whether the scrollbar under the band is shown.</summary>
    bool ShowHorizontalScrollbar { get; set; }

    /// <summary>
    /// The name of the timeline style, for example <c>TimeSlicerStyleLight2</c>. <c>null</c> means
    /// the workbook default.
    /// </summary>
    /// <remarks>
    /// Deliberately a string rather than an enumeration of the built-in styles: a workbook may name
    /// a custom style, and a read model that could only report the styles it knows about would
    /// silently lose the rest.
    /// </remarks>
    string? Style { get; set; }

    /// <summary>How finely the band is divided.</summary>
    XLTimelineLevel Level { get; set; }

    /// <summary>The cell the timeline's top-left corner is anchored to. Setting it moves the timeline.</summary>
    /// <remarks>
    /// Moving a timeline read from a file shifts both of its corners together, so it keeps the size
    /// it had. Reading reports the cell the corner sits in; a file may also place the corner some
    /// distance into that cell, and that offset is preserved through a save but is not reported here.
    /// </remarks>
    IXLCell Position { get; set; }

    /// <summary>The pivot cache field the band is drawn from.</summary>
    string SourceFieldName { get; }

    /// <summary>The worksheet the timeline is drawn on.</summary>
    IXLWorksheet Worksheet { get; }

    /// <summary>
    /// The pivot tables this timeline filters. A pivot table listed in the cache but missing from
    /// the workbook is omitted rather than reported as null.
    /// </summary>
    IReadOnlyList<IXLPivotTable> PivotTables { get; }

    /// <summary>
    /// The extent of the scrubber — the date field's range, rounded outward. <c>null</c> when the
    /// file records no bounds.
    /// </summary>
    /// <remarks>
    /// Read-only: Excel recomputes the extent when the pivot cache refreshes, so a settable bound
    /// would be honest in only one direction.
    /// </remarks>
    DateTime? BoundsStart { get; }

    /// <inheritdoc cref="BoundsStart"/>
    DateTime? BoundsEnd { get; }

    /// <summary>
    /// Whether the timeline records an explicit range. <c>false</c> means every date is showing,
    /// which is how Excel represents a timeline nobody has scrubbed.
    /// </summary>
    bool HasSelection { get; }

    /// <summary>
    /// The first date of the selected range, or <c>null</c> when <see cref="HasSelection"/> is
    /// <c>false</c>.
    /// </summary>
    /// <remarks>
    /// Read-only. Excel records a timeline's range in three places at once — the cache's state, a
    /// <c>dateBetween</c> filter on the pivot table, and hidden-item flags on the pivot field — and
    /// a model that wrote one without the others would produce a workbook that disagrees with
    /// itself in a way no validator can see.
    /// </remarks>
    DateTime? SelectionStart { get; }

    /// <inheritdoc cref="SelectionStart"/>
    DateTime? SelectionEnd { get; }
}
