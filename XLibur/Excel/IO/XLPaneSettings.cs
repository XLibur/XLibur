using XLibur.Excel.Coordinates;

namespace XLibur.Excel.IO;

/// <summary>Which corner of a split view holds the active pane.</summary>
internal enum XLPaneCorner
{
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight,
}

/// <summary>
/// <c>ST_PaneState</c>. XLibur writes <see cref="Frozen"/> and <see cref="Split"/>, chosen by
/// <c>IXLSheetView.FreezePanes</c>. <see cref="FrozenSplit"/> — a pane frozen out of an existing
/// manual split — is accepted by the reader and normalised to <see cref="Frozen"/>, because the
/// model carries a boolean and has never had a third state to put it in.
/// </summary>
internal enum XLPaneState
{
    Frozen,
    FrozenSplit,
    Split,
}

/// <summary>
/// The <c>&lt;pane&gt;</c> attributes both write paths emit, with defaults applied.
/// </summary>
/// <remarks>
/// <para>
/// This type owns the decision; the two writers own only the emission. Before spec 29 the decision
/// was made twice and the two copies disagreed on <c>state</c> — the DOM path wrote
/// <c>frozenSplit</c> for every pane while the streaming path wrote <c>frozen</c>, and the reader
/// mapped both back to the same model so no round-trip test could see it.
/// </para>
/// <para>
/// The enums are XLibur's own rather than the SDK's <c>PaneStateValues</c> / <c>PaneValues</c>, so
/// the streaming path does not take an OpenXML dependency to read a decision it renders as a raw
/// string. Each writer maps at the point of emission.
/// </para>
/// </remarks>
internal readonly struct XLPaneSettings
{
    /// <summary><c>xSplit</c>, or <c>null</c> when the axis is not split and the attribute is omitted.</summary>
    internal required int? SplitColumn { get; init; }

    /// <summary><c>ySplit</c>, or <c>null</c> when the axis is not split.</summary>
    internal required int? SplitRow { get; init; }

    /// <summary><c>topLeftCell</c>, the first cell of the scrollable pane.</summary>
    internal required string TopLeftCell { get; init; }

    /// <summary><c>activePane</c>.</summary>
    internal required XLPaneCorner ActivePane { get; init; }

    /// <summary><c>state</c>.</summary>
    internal required XLPaneState State { get; init; }

    /// <summary><c>false</c> when no <c>&lt;pane&gt;</c> element should be written at all.</summary>
    internal bool HasPane => SplitColumn is not null || SplitRow is not null;

    /// <param name="splitColumn">Split position on the column axis, 0 for none.</param>
    /// <param name="splitRow">Split position on the row axis, 0 for none.</param>
    /// <param name="frozen">
    /// <c>IXLSheetView.FreezePanes</c>: <c>true</c> for a frozen pane, <c>false</c> for a draggable
    /// split bar. The streaming path always passes <c>true</c>; it only offers <c>FreezePanes</c>.
    /// </param>
    /// <param name="paneTopLeftCell">
    /// An explicit pane scroll position, or <c>null</c> to anchor at split + 1 for a freeze and at
    /// A1 for a split. The streaming path always passes <c>null</c>; it exposes no equivalent API.
    /// </param>
    /// <param name="activeCell">
    /// The active cell, or <c>null</c>. When set it decides which corner owns the active pane;
    /// otherwise the split shape does. Streaming always passes <c>null</c>.
    /// </param>
    internal static XLPaneSettings Resolve(
        int splitColumn, int splitRow, bool frozen, XLAddress? paneTopLeftCell, Point? activeCell)
    {
        return new XLPaneSettings
        {
            // S125 reads the xSplit line below as commented-out code; it is prose quoting the
            // attribute that used to be written.
#pragma warning disable S125
            // Excel omits the unused axis rather than writing 0. Before spec 29 the DOM path wrote
            // xSplit="0" for a rows-only freeze where the streaming path omitted the attribute;
            // task 1's harness read both packages and confirmed it.
            SplitColumn = splitColumn > 0 ? splitColumn : null,
#pragma warning restore S125
            SplitRow = splitRow > 0 ? splitRow : null,
            TopLeftCell = ResolveTopLeftCell(splitColumn, splitRow, frozen, paneTopLeftCell),
            ActivePane = ResolveCorner(splitColumn, splitRow, activeCell),
            // XLibur never produces a split-then-frozen pane, so the choice is the sheet view's
            // own: a freeze, or the draggable split bar a caller gets by setting SplitRow and
            // SplitColumn without freezing.
            State = frozen ? XLPaneState.Frozen : XLPaneState.Split,
        };
    }

    /// <summary>
    /// The cell the unfrozen pane scrolls to: the caller's, else split + 1 for a freeze, else A1.
    /// </summary>
    /// <remarks>
    /// split + 1 is the first unfrozen cell, and only means that for a freeze: an unfrozen split
    /// states its position in twentieths of a point, so the same arithmetic would name a cell
    /// hundreds of columns away, or none at all. Such a pane anchors at A1 unless the caller
    /// named a cell.
    /// </remarks>
    private static string ResolveTopLeftCell(
        int splitColumn, int splitRow, bool frozen, XLAddress? paneTopLeftCell)
    {
        if (paneTopLeftCell is { IsValid: true } p)
            return p.ToStringRelative(false);

        return frozen
            ? XLHelper.GetColumnLetterFromNumber(splitColumn + 1) + (splitRow + 1)
            : "A1";
    }

    /// <summary>
    /// The active pane (and therefore the selection's pane) must name the pane that actually owns
    /// the active cell. When no active cell is set, fall back to the split-derived default.
    /// </summary>
    private static XLPaneCorner ResolveCorner(int splitColumn, int splitRow, Point? activeCell)
    {
        if (activeCell is not { } active)
        {
            if (splitRow == 0 && splitColumn == 0)
                return XLPaneCorner.TopLeft;
            if (splitRow == 0)
                return XLPaneCorner.TopRight;
            return splitColumn == 0 ? XLPaneCorner.BottomLeft : XLPaneCorner.BottomRight;
        }

        var bottom = splitRow > 0 && active.Row > splitRow;
        var right = splitColumn > 0 && active.Column > splitColumn;

        if (bottom && right)
            return XLPaneCorner.BottomRight;
        if (bottom)
            return XLPaneCorner.BottomLeft;
        return right ? XLPaneCorner.TopRight : XLPaneCorner.TopLeft;
    }
}
