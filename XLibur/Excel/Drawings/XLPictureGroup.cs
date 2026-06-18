namespace XLibur.Excel.Drawings;

/// <summary>
/// Metadata describing a picture that lives inside a drawing group shape
/// (<c>xdr:grpSp</c>), possibly nested several levels deep. Children of a group are laid
/// out in the group's <em>child</em> coordinate space and scaled to the sheet via the
/// group's transform (<c>a:xfrm</c> / <c>chOff</c> / <c>chExt</c> vs <c>off</c> / <c>ext</c>).
/// <para>
/// For nested groups the per-level transforms are composed, so <see cref="ScaleX"/> /
/// <see cref="ScaleY"/> are the <em>total</em> scale from the picture's own child coordinate
/// space to the sheet. The picture is linked back to its <c>xdr:pic</c> element on save via
/// the picture's drawing id (<see cref="XLPicture.Id"/>) — which is unique across the whole
/// drawing regardless of nesting depth — because the load and save DOMs are independent
/// re-parses of the package.
/// </para>
/// </summary>
internal sealed class XLPictureGroup
{
    /// <summary>Total horizontal scale from the picture's child coordinate space to the sheet.</summary>
    public double ScaleX { get; init; } = 1.0;

    /// <summary>Total vertical scale from the picture's child coordinate space to the sheet.</summary>
    public double ScaleY { get; init; } = 1.0;

    /// <summary>
    /// Composed horizontal offset (EMU) of the affine child→sheet mapping: a child x maps to the
    /// sheet at <c>OffsetX + childX · ScaleX</c>. Used to convert a moved picture's sheet position
    /// back to its child <c>a:off</c>.
    /// </summary>
    public double OffsetX { get; init; }

    /// <summary>Composed vertical offset (EMU) of the affine child→sheet mapping.</summary>
    public double OffsetY { get; init; }

    /// <summary>The picture's width in pixels as computed at load time (baseline for change detection).</summary>
    public int LoadedWidthPx { get; init; }

    /// <summary>The picture's height in pixels as computed at load time (baseline for change detection).</summary>
    public int LoadedHeightPx { get; init; }

    /// <summary>The picture's left position in pixels as computed at load time (baseline for change detection).</summary>
    public int LoadedLeftPx { get; init; }

    /// <summary>The picture's top position in pixels as computed at load time (baseline for change detection).</summary>
    public int LoadedTopPx { get; init; }

    /// <summary>
    /// The <c>cNvPr</c> id of the immediate parent group shape. Used to locate the group element
    /// when adding to or removing from a group; the picture itself is still located by its own id.
    /// Null for a group created in this session that hasn't been saved yet.
    /// </summary>
    public uint? GroupId { get; init; }

    /// <summary>
    /// Stable per-worksheet identity of the group this picture belongs to. Unlike <see cref="GroupId"/>
    /// (the drawing's <c>cNvPr</c> id, unknown until save for new groups), this is assigned when the
    /// group is loaded or created and never changes, so the public <c>IXLPictureGroup</c> can identify
    /// a group's members consistently before and after a save.
    /// </summary>
    public long GroupKey { get; init; }

    /// <summary>
    /// True for a picture that has been added to a group but not yet written: the writer must create
    /// a new <c>xdr:pic</c> inside the group element rather than updating an existing one. The
    /// composed transform (<see cref="ScaleX"/>/<see cref="ScaleY"/>/<see cref="OffsetX"/>/<see
    /// cref="OffsetY"/>) is the target group's, copied from an existing member.
    /// </summary>
    public bool IsNew { get; init; }
}
