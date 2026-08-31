using System;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel.Drawings;

namespace XLibur.Excel;

[DebuggerDisplay("{Name} ({SourceFieldName})")]
internal sealed class XLTimeline : IXLTimeline
{
    /// <summary>
    /// The size Excel gives a new timeline, in pixels — measured off the round-trip fixture's frame,
    /// 3,333,750 × 1,371,600 EMU at 96 dpi. A timeline is wider and shorter than a slicer.
    /// </summary>
    internal const int DefaultWidthPx = 350;

    /// <inheritdoc cref="DefaultWidthPx"/>
    internal const int DefaultHeightPx = 144;

    private readonly XLWorksheet _worksheet;
    private string _caption;
    private bool _showHeader = true;
    private bool _showSelectionLabel = true;
    private bool _showTimeLevel = true;
    private bool _showHorizontalScrollbar = true;
    private string? _style;
    private uint _level;

    internal XLTimeline(XLWorksheet worksheet, XLTimelineCache cache, string name)
    {
        _worksheet = worksheet;
        Cache = cache;
        Name = name;
        _caption = name;
    }

    /// <summary>The cache that binds the timeline to what it filters and holds its range.</summary>
    internal XLTimelineCache Cache { get; }

    /// <summary>
    /// The id of the relationship from the worksheet part to the timelines part this timeline was
    /// read from. Together with <see cref="Name"/> this is how the write path finds the element
    /// again to patch it. Null for a timeline not read from a package.
    /// </summary>
    internal string? PartRelId { get; set; }

    /// <summary>
    /// Whether the timeline was created through the API rather than read from a package. A new
    /// timeline is generated on save; a loaded one is only ever patched.
    /// </summary>
    internal bool IsNew { get; set; }

    /// <summary>Which properties the caller has assigned since the timeline was loaded.</summary>
    internal XLTimelineFormat AssignedFormat { get; private set; }

    /// <summary>
    /// The raw <c>@selectionLevel</c>, carried through untouched. XLibur does not model it — it only
    /// means anything alongside a selection, which is read-only.
    /// </summary>
    internal uint? SelectionLevelRaw { get; set; }

    /// <summary>
    /// The raw <c>@scrollPosition</c>, carried through untouched. It records where the user had
    /// scrolled the band, which XLibur has no reason to change.
    /// </summary>
    internal DateTime? ScrollPosition { get; set; }

    public string Name { get; }

    public string Caption
    {
        get => _caption;
        set
        {
            _caption = value ?? throw new ArgumentNullException(nameof(value));
            AssignedFormat |= XLTimelineFormat.Caption;
        }
    }

    public bool ShowHeader
    {
        get => _showHeader;
        set
        {
            _showHeader = value;
            AssignedFormat |= XLTimelineFormat.ShowHeader;
        }
    }

    public bool ShowSelectionLabel
    {
        get => _showSelectionLabel;
        set
        {
            _showSelectionLabel = value;
            AssignedFormat |= XLTimelineFormat.ShowSelectionLabel;
        }
    }

    public bool ShowTimeLevel
    {
        get => _showTimeLevel;
        set
        {
            _showTimeLevel = value;
            AssignedFormat |= XLTimelineFormat.ShowTimeLevel;
        }
    }

    public bool ShowHorizontalScrollbar
    {
        get => _showHorizontalScrollbar;
        set
        {
            _showHorizontalScrollbar = value;
            AssignedFormat |= XLTimelineFormat.ShowHorizontalScrollbar;
        }
    }

    public string? Style
    {
        get => _style;
        set
        {
            _style = value;
            AssignedFormat |= XLTimelineFormat.Style;
        }
    }

    /// <summary>
    /// The level, as the enumeration. The raw number is what is stored and what is written back, so
    /// a file carrying a value outside the enumeration round-trips its number rather than being
    /// narrowed to the nearest modelled one.
    /// </summary>
    public XLTimelineLevel Level
    {
        get => (XLTimelineLevel)_level;
        set
        {
            _level = (uint)value;
            AssignedFormat |= XLTimelineFormat.Level;
        }
    }

    /// <inheritdoc cref="Level"/>
    internal uint LevelRaw => _level;

    public IXLCell Position
    {
        get => FromMarker?.Cell ?? _worksheet.Cell(1, 1);
        set
        {
            ArgumentNullException.ThrowIfNull(value);

            // A fresh marker rather than a mutated one, because a marker registers itself with the
            // workbook's range repository so that inserting rows above the timeline moves it.
            // Setting the position drops any offset within the old cell: the caller named a cell,
            // so the corner goes to that cell's corner.
            FromMarker = new XLMarker(value);
            AssignedFormat |= XLTimelineFormat.Position;
        }
    }

    /// <summary>
    /// The timeline's top-left anchor point, with the offset within the cell that a file may carry.
    /// </summary>
    internal XLMarker? FromMarker { get; set; }

    /// <summary>
    /// The timeline's bottom-right anchor point, when it was read from a two-cell anchor. Kept so
    /// that moving a loaded timeline shifts both corners together and leaves its size alone.
    /// </summary>
    internal XLMarker? ToMarker { get; set; }

    internal int WidthPx { get; set; } = DefaultWidthPx;

    internal int HeightPx { get; set; } = DefaultHeightPx;

    /// <summary>
    /// Sets the properties read from a package without marking them as assigned.
    /// </summary>
    /// <remarks>
    /// This is what keeps <see cref="AssignedFormat"/> honest. It has to stay the only way the
    /// reader populates a timeline: assigning through the properties instead would mark every loaded
    /// timeline as edited, and the patcher would then rewrite parts nobody touched.
    /// </remarks>
    internal void SeedLoadedFormat(
        string caption,
        bool showHeader,
        bool showSelectionLabel,
        bool showTimeLevel,
        bool showHorizontalScrollbar,
        string? style,
        uint level)
    {
        _caption = caption;
        _showHeader = showHeader;
        _showSelectionLabel = showSelectionLabel;
        _showTimeLevel = showTimeLevel;
        _showHorizontalScrollbar = showHorizontalScrollbar;
        _style = style;
        _level = level;
    }

    public string SourceFieldName => Cache.SourceName;

    public IXLWorksheet Worksheet => _worksheet;

    public IReadOnlyList<IXLPivotTable> PivotTables => Cache.PivotTables;

    public DateTime? BoundsStart => Cache.BoundsStart;

    public DateTime? BoundsEnd => Cache.BoundsEnd;

    public bool HasSelection => Cache.SelectionStart is not null || Cache.SelectionEnd is not null;

    public DateTime? SelectionStart => Cache.SelectionStart;

    public DateTime? SelectionEnd => Cache.SelectionEnd;
}
