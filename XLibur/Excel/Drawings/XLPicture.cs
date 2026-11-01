using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using XLibur.Graphics;

namespace XLibur.Excel.Drawings;

[DebuggerDisplay("{Name}")]
internal sealed partial class XLPicture : IXLPicture
{
    private int _height;
    private int _id;
    private string _name = string.Empty;
    private XLPicturePlacement _placement;
    private int _width;

    internal XLPicture(XLWorksheet worksheet, Stream stream)
        : this(worksheet, stream, XLPictureFormat.Unknown)
    {
    }

    internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
        : this(worksheet)
    {
        ArgumentNullException.ThrowIfNull(stream);

        var info = worksheet.Workbook.GraphicEngine.GetPictureInfo(stream, format);
        Init(info);

        _image = new XLPictureImageStream();
        stream.Position = 0;
        stream.CopyTo(_image);
        _image.Seek(0, SeekOrigin.Begin);
        _image.MarkUnmodified();
    }

    private XLPicture(IXLWorksheet worksheet)
    {
        Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        Placement = XLPicturePlacement.MoveAndSize;
        Markers = new Dictionary<XLMarkerPosition, XLMarker?>
        {
            [XLMarkerPosition.TopLeft] = null,
            [XLMarkerPosition.BottomRight] = null
        };

        _id = XLPictures.NextFreeId(worksheet.Workbook);
    }

    public IXLCell BottomRightCell
    {
        get => Markers[XLMarkerPosition.BottomRight]!.Cell;

        private set
        {
            if (!value.Worksheet.Equals(Worksheet))
                throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

            Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
        }
    }

    public XLPictureFormat Format { get; private set; } = XLPictureFormat.Unknown;

    public int Height
    {
        get => _height;
        set
        {
            if (Placement == XLPicturePlacement.MoveAndSize)
                throw new ArgumentException(
                    $"Cannot set the height when placement is '{Placement}'. Change the placement to FreeFloating or Move first.");
            _height = value;
            KeptAnchor = null;
        }
    }

    public int Id
    {
        get => _id;
        internal set
        {
            if (Collection.IsIdTaken(value, this))
                throw new ArgumentException($"The picture ID '{value}' already exists.");

            var oldId = _id;
            _id = value;
            OnIdChanged(oldId);
        }
    }

    /// <summary>The pictures of <see cref="Worksheet"/>, which this picture is one of while <see cref="IsInCollection"/>.</summary>
    private XLPictures Collection => (XLPictures)Worksheet.Pictures;

    /// <summary>
    /// Whether this picture is one of <see cref="Collection"/>: from when it is added until it is
    /// deleted. Set by the collection, which keeps the names and ids its pictures hold.
    /// </summary>
    internal bool IsInCollection { get; set; }

    /// <summary>Tell <see cref="Collection"/> that <see cref="Id"/> was <paramref name="oldId"/>.</summary>
    private void OnIdChanged(int oldId)
    {
        if (IsInCollection && oldId != _id)
            Collection.OnIdChanged(this, oldId);
    }

    /// <summary>
    /// Take the drawing id and image link a save wrote for this picture, once that save has
    /// succeeded. A save gives a new picture its drawing id, and can give a picture an image part of
    /// its own, and the next save finds the picture's anchor by both. The check in <see cref="Id"/>
    /// is skipped: the pictures take their new ids one at a time, so two of them can briefly hold
    /// the same one. The picture is then as the file holds it (<see cref="MarkSaved"/>).
    /// </summary>
    internal void KeepSavedIds(int id, string relId)
    {
        var oldId = _id;
        _id = id;
        OnIdChanged(oldId);
        RelId = relId;
        MarkSaved();
    }

    private readonly XLPictureImageStream _image = null!;

    public MemoryStream ImageStream => _image;

    /// <summary>
    /// What the picture was when it was loaded or last saved, or <c>null</c> for a picture no save
    /// has written yet. A save writes a picture only when it differs from this (spec 17).
    /// </summary>
    private XLPictureSavedState? _saved;

    /// <summary>Whether the image was written to since the picture was loaded or last saved.</summary>
    internal bool IsImageModified => _saved is null || _image.IsModified;

    /// <summary>
    /// Whether the placement, a marker, the size or the position differs from what the picture was
    /// loaded or last saved with. A picture no save has written yet has changed.
    /// </summary>
    internal bool IsGeometryChanged => _saved is not { } saved || saved.Geometry != CaptureGeometry();

    /// <summary>Whether the width or the height differs from what the picture was loaded or last saved with.</summary>
    internal bool IsSizeChanged =>
        _saved is not { } saved || saved.Geometry.Width != _width || saved.Geometry.Height != _height;

    /// <summary>
    /// Whether a save has to write this picture: it is new, or anything the file holds for it has
    /// changed since it was loaded or last saved. <paramref name="layout"/> is the sheet's layout
    /// now; a kept anchor whose rows or columns were resized has to be written from the model.
    /// </summary>
    internal bool HasUnsavedChanges(Lazy<GridLayout> layout) =>
        _saved is not { } saved ||
        _image.IsModified ||
        AssignedStyle != XLPictureStyleProperties.None ||
        !string.Equals(saved.Name, _name, StringComparison.Ordinal) ||
        saved.Geometry != CaptureGeometry() ||
        (KeptAnchor is { } kept && !kept.IsLayoutUnchanged(layout.Value));

    /// <summary>
    /// Record the picture as the file now holds it: once it is loaded, and once a save that wrote it
    /// has succeeded.
    /// </summary>
    internal void MarkSaved()
    {
        _saved = new XLPictureSavedState(_name, CaptureGeometry());
        _image.MarkUnmodified();
        AssignedStyle = XLPictureStyleProperties.None;
    }

    private XLPictureGeometry CaptureGeometry()
    {
        var from = Markers[XLMarkerPosition.TopLeft];
        var to = Markers[XLMarkerPosition.BottomRight];
        return new XLPictureGeometry(Placement, _width, _height,
            from?.Anchor, from?.Offset, to?.Anchor, to?.Offset);
    }

    public int Left
    {
        get => Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0;
        set
        {
            if (Placement != XLPicturePlacement.FreeFloating)
                throw new ArgumentException(
                    $"Cannot set the left-hand offset when placement is '{Placement}'. Change the placement to FreeFloating first.");

            Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(value, Top));
            KeptAnchor = null;
        }
    }

    public string Name
    {
        get => _name;
        set
        {
            ArgumentException.ThrowIfNullOrEmpty(value);
            if (_name == value) return;

            // A name a kept shape has, such as a linked picture the loader does not read, is taken
            // as a picture's is (#796).
            Collection.ThrowIfNameTaken(value, this);

            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Picture name cannot be empty or consist only of whitespace.",
                    nameof(value));

            var oldName = _name;
            _name = value;
            OnRenamed(oldName);
        }
    }

    public int OriginalHeight { get; private set; }

    public int OriginalWidth { get; private set; }

    public XLPicturePlacement Placement
    {
        get => _placement;
        set
        {
            _placement = value;
            KeptAnchor = null;
        }
    }

    public int Top
    {
        get => Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0;
        set
        {
            if (Placement != XLPicturePlacement.FreeFloating)
                throw new ArgumentException(
                    $"Cannot set the top offset when placement is '{Placement}'. Change the placement to FreeFloating first.");

            Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(Left, value));
            KeptAnchor = null;
        }
    }

    public IXLCell TopLeftCell
    {
        get => Markers[XLMarkerPosition.TopLeft]!.Cell;

        private set
        {
            if (!value.Worksheet.Equals(Worksheet))
                throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

            Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
        }
    }

    public int Width
    {
        get => _width;
        set
        {
            if (Placement == XLPicturePlacement.MoveAndSize)
                throw new ArgumentException(
                    $"Cannot set the width when placement is '{Placement}'. Change the placement to FreeFloating or Move first.");
            _width = value;
            KeptAnchor = null;
        }
    }

    public IXLWorksheet Worksheet { get; }

    internal IDictionary<XLMarkerPosition, XLMarker?> Markers { get; private set; }

    internal string? RelId { get; set; }

    /// <summary>
    /// The loaded two-cell anchor, when the model cannot say it back. Set by the loader after it has
    /// placed the picture, and dropped by any edit of the placement, size or position. See
    /// <see cref="XLKeptAnchor"/>.
    /// </summary>
    internal XLKeptAnchor? KeptAnchor { get; set; }

    /// <summary>
    /// When set, this picture lives inside a drawing group shape (<c>xdr:grpSp</c>).
    /// Grouped pictures are written back in place (their <c>xdr:pic</c> element is updated)
    /// rather than re-anchored, so sibling pictures, connectors and shapes in the group are
    /// preserved. See <see cref="XLPictureGroup"/>.
    /// </summary>
    internal XLPictureGroup? GroupInfo { get; set; }

    /// <summary>Whether this picture is part of a drawing group shape.</summary>
    public bool IsInGroup => GroupInfo is not null;

    /// <summary>The group this picture belongs to, or <c>null</c> if it is not grouped.</summary>
    public IXLPictureGroup? Group =>
        GroupInfo is null ? null : new XLPictureGroupView((XLWorksheet)Worksheet, GroupInfo.GroupKey);

    /// <summary>
    /// Create a copy of the picture on a different worksheet.
    /// </summary>
    /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
    /// <returns>A created copy of the picture.</returns>
    public IXLPicture CopyTo(IXLWorksheet targetSheet)
    {
        return CopyTo((XLWorksheet)targetSheet);
    }

    public void Delete()
    {
        Worksheet.Pictures.Delete(Name);
    }

    #region IDisposable

    public void Dispose()
    {
        ImageStream.Dispose();
    }

    #endregion IDisposable

    /// <summary>
    /// Create a copy of the picture on the same worksheet.
    /// </summary>
    /// <returns>A created copy of the picture.</returns>
    public IXLPicture Duplicate()
    {
        return CopyTo(Worksheet);
    }

    public Point GetOffset(XLMarkerPosition position)
    {
        return Markers[position]!.Offset;
    }

    public IXLPicture MoveTo(int left, int top)
    {
        Placement = XLPicturePlacement.FreeFloating;
        Left = left;
        Top = top;
        return this;
    }

    public IXLPicture MoveTo(IXLCell cell)
    {
        return MoveTo(cell, 0, 0);
    }

    public IXLPicture MoveTo(IXLCell cell, int xOffset, int yOffset)
    {
        return MoveTo(cell, new Point(xOffset, yOffset));
    }

    public IXLPicture MoveTo(IXLCell cell, Point offset)
    {
        Placement = XLPicturePlacement.Move;
        TopLeftCell = cell ?? throw new ArgumentNullException(nameof(cell));
        Markers[XLMarkerPosition.TopLeft]!.Offset = offset;
        return this;
    }

    public IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell)
    {
        return MoveTo(fromCell, 0, 0, toCell, 0, 0);
    }

    public IXLPicture MoveTo(IXLCell fromCell, int fromCellXOffset, int fromCellYOffset, IXLCell toCell,
        int toCellXOffset, int toCellYOffset)
    {
        return MoveTo(fromCell, new Point(fromCellXOffset, fromCellYOffset), toCell,
            new Point(toCellXOffset, toCellYOffset));
    }

    public IXLPicture MoveTo(IXLCell fromCell, Point fromOffset, IXLCell toCell, Point toOffset)
    {
        Placement = XLPicturePlacement.MoveAndSize;

        TopLeftCell = fromCell ?? throw new ArgumentNullException(nameof(fromCell));
        Markers[XLMarkerPosition.TopLeft]!.Offset = fromOffset;

        BottomRightCell = toCell ?? throw new ArgumentNullException(nameof(toCell));
        Markers[XLMarkerPosition.BottomRight]!.Offset = toOffset;

        return this;
    }

    public IXLPicture Scale(double factor, bool relativeToOriginal = false)
    {
        return ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);
    }

    public IXLPicture ScaleHeight(double factor, bool relativeToOriginal = false)
    {
        Height = Convert.ToInt32((relativeToOriginal ? OriginalHeight : Height) * factor);
        return this;
    }

    public IXLPicture ScaleWidth(double factor, bool relativeToOriginal = false)
    {
        Width = Convert.ToInt32((relativeToOriginal ? OriginalWidth : Width) * factor);
        return this;
    }

    public IXLPicture WithPlacement(XLPicturePlacement value)
    {
        Placement = value;
        return this;
    }

    public IXLPicture WithSize(int width, int height)
    {
        Width = width;
        Height = height;
        return this;
    }

    private IXLPicture CopyTo(XLWorksheet targetSheet)
    {
        var newPicture = targetSheet == Worksheet
            ? targetSheet.AddPicture(ImageStream, Format)
            : targetSheet.AddPicture(ImageStream, Format, Name);

        newPicture = newPicture
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .WithSize(Width, Height)
            .WithPlacement(Placement);

        switch (Placement)
        {
            case XLPicturePlacement.FreeFloating:
                newPicture.MoveTo(Left, Top);
                break;

            case XLPicturePlacement.Move:
                newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), GetOffset(XLMarkerPosition.TopLeft));
                break;

            case XLPicturePlacement.MoveAndSize:
                newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), GetOffset(XLMarkerPosition.TopLeft),
                    targetSheet.Cell(BottomRightCell.Address),
                    GetOffset(XLMarkerPosition.BottomRight));
                break;
        }

        ((XLPicture)newPicture).TakeStyleOf(this);
        return newPicture;
    }

    internal void SetName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Picture name cannot be empty or consist only of whitespace.", nameof(value));

        var oldName = _name;
        _name = value;
        OnRenamed(oldName);
    }

    /// <summary>Tell <see cref="Collection"/> that <see cref="Name"/> was <paramref name="oldName"/>.</summary>
    private void OnRenamed(string oldName)
    {
        if (IsInCollection)
            Collection.OnRenamed(this, oldName);
    }

    private void Init(XLPictureInfo info)
    {
        Format = info.Format;
        var size = info.GetSizePx(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
        _width = OriginalWidth = size.Width;
        _height = OriginalHeight = size.Height;
    }
}

/// <summary>Where a picture is and how big, as the model holds it: what decides its anchor.</summary>
internal readonly record struct XLPictureGeometry(
    XLPicturePlacement Placement,
    int Width,
    int Height,
    XLibur.Excel.Coordinates.Point? From,
    Point? FromOffset,
    XLibur.Excel.Coordinates.Point? To,
    Point? ToOffset);

/// <summary>What a picture was when it was loaded or last saved.</summary>
internal sealed record XLPictureSavedState(string Name, XLPictureGeometry Geometry);

