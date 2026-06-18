namespace XLibur.Excel.Drawings;

/// <summary>
/// Metadata describing a picture that lives inside a drawing group shape
/// (<c>xdr:grpSp</c>). Children of a group are laid out in the group's <em>child</em>
/// coordinate space and scaled to the sheet via the group's transform
/// (<c>a:xfrm</c> / <c>chOff</c> / <c>chExt</c> vs <c>off</c> / <c>ext</c>).
/// <para>
/// Geometry is stored in EMUs exactly as read from the file so that an unedited
/// picture can be written back without rounding drift. The picture is linked back
/// to its <c>xdr:pic</c> element on save via the picture's drawing id
/// (<see cref="XLPicture.Id"/>), because the load and save DOMs are independent
/// re-parses of the package.
/// </para>
/// </summary>
internal sealed class XLPictureGroup
{
    /// <summary>Group offset on the sheet (EMU), from <c>grpSpPr/a:xfrm/a:off</c>.</summary>
    public long GroupOffsetX { get; init; }

    public long GroupOffsetY { get; init; }

    /// <summary>Group extent on the sheet (EMU), from <c>grpSpPr/a:xfrm/a:ext</c>.</summary>
    public long GroupExtentCx { get; init; }

    public long GroupExtentCy { get; init; }

    /// <summary>Child coordinate origin (EMU), from <c>grpSpPr/a:xfrm/a:chOff</c>.</summary>
    public long ChildOffsetX { get; init; }

    public long ChildOffsetY { get; init; }

    /// <summary>Child coordinate extent (EMU), from <c>grpSpPr/a:xfrm/a:chExt</c>.</summary>
    public long ChildExtentCx { get; init; }

    public long ChildExtentCy { get; init; }

    /// <summary>The picture's own extent in child coordinate space (EMU) at load time.</summary>
    public long PictureChildExtentCx { get; init; }

    public long PictureChildExtentCy { get; init; }

    /// <summary>The picture's width in pixels as computed at load time (the baseline for change detection).</summary>
    public int LoadedWidthPx { get; init; }

    /// <summary>The picture's height in pixels as computed at load time (the baseline for change detection).</summary>
    public int LoadedHeightPx { get; init; }

    /// <summary>Horizontal scale applied to children to map child space onto the sheet.</summary>
    public double ScaleX => ChildExtentCx == 0 ? 1.0 : (double)GroupExtentCx / ChildExtentCx;

    /// <summary>Vertical scale applied to children to map child space onto the sheet.</summary>
    public double ScaleY => ChildExtentCy == 0 ? 1.0 : (double)GroupExtentCy / ChildExtentCy;
}
