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

    /// <summary>The picture's width in pixels as computed at load time (baseline for change detection).</summary>
    public int LoadedWidthPx { get; init; }

    /// <summary>The picture's height in pixels as computed at load time (baseline for change detection).</summary>
    public int LoadedHeightPx { get; init; }

    /// <summary>
    /// The <c>cNvPr</c> id of the immediate parent group shape. Used to locate the group element
    /// when adding to or removing from a group (a later phase); the picture itself is still located
    /// by its own id.
    /// </summary>
    public uint? GroupId { get; init; }
}
