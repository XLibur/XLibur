using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using XLibur.Graphics;

namespace XLibur.Excel.Drawings;

[DebuggerDisplay("{Name}")]
internal sealed class XLPicture : IXLPicture
{
    private int _height;
    private int _id;
    private string _name = string.Empty;
    private int _width;

    internal XLPicture(XLWorksheet worksheet, Stream stream)
        : this(worksheet, stream, XLPictureFormat.Unknown)
    {
    }

    internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
        : this(worksheet)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        var info = worksheet.Workbook.GraphicEngine.GetPictureInfo(stream, format);
        Init(info);

        ImageStream = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(ImageStream);
        ImageStream.Seek(0, SeekOrigin.Begin);
    }

    private XLPicture(IXLWorksheet worksheet)
    {
        Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        Placement = XLPicturePlacement.MoveAndSize;
        Markers = new Dictionary<XLMarkerPosition, XLMarker?>()
        {
            [XLMarkerPosition.TopLeft] = null,
            [XLMarkerPosition.BottomRight] = null
        };

        // Calculate default picture ID
        var allPictures = worksheet.Workbook.Worksheets.SelectMany(ws => ws.Pictures);
        var freeId = allPictures.Select(x => x.Id).DefaultIfEmpty(0).Max() + 1;
        _id = freeId;
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
        }
    }

    public int Id
    {
        get => _id;
        internal set
        {
            if ((Worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
                throw new ArgumentException($"The picture ID '{value}' already exists.");

            _id = value;
        }
    }

    public MemoryStream ImageStream { get; private set; } = null!;

    public int Left
    {
        get => Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0;
        set
        {
            if (Placement != XLPicturePlacement.FreeFloating)
                throw new ArgumentException(
                    $"Cannot set the left-hand offset when placement is '{Placement}'. Change the placement to FreeFloating first.");

            Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(value, Top));
        }
    }

    public string Name
    {
        get => _name;
        set
        {
            ArgumentException.ThrowIfNullOrEmpty(value);
            if (_name == value) return;

            if ((Worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ??
                 this) != this)
                throw new ArgumentException($"The picture name '{value}' already exists.");

            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Picture name cannot be empty or consist only of whitespace.",
                    nameof(value));

            _name = value;
        }
    }

    public int OriginalHeight { get; private set; }

    public int OriginalWidth { get; private set; }

    public XLPicturePlacement Placement { get; set; }

    public int Top
    {
        get => Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0;
        set
        {
            if (Placement != XLPicturePlacement.FreeFloating)
                throw new ArgumentException(
                    $"Cannot set the top offset when placement is '{Placement}'. Change the placement to FreeFloating first.");

            Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(Left, value));
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
        }
    }

    public IXLWorksheet Worksheet { get; }

    internal IDictionary<XLMarkerPosition, XLMarker?> Markers { get; private set; }

    internal string? RelId { get; set; }

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

        return newPicture;
    }

    internal void SetName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Picture name cannot be empty or consist only of whitespace.", nameof(value));

        _name = value;
    }

    private void Init(XLPictureInfo info)
    {
        Format = info.Format;
        var size = info.GetSizePx(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
        _width = OriginalWidth = size.Width;
        _height = OriginalHeight = size.Height;
    }
}

