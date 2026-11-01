using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;

namespace XLibur.Excel.Drawings;

/// <summary>The formatting properties of a shape a save writes, when they were assigned.</summary>
[Flags]
internal enum XLShapeProperties
{
    None = 0,
    FillColor = 1 << 0,
    NoFill = 1 << 1,
    LineColor = 1 << 2,
    LineWidth = 1 << 3,
    LineDash = 1 << 4,
    NoLine = 1 << 5,
}

/// <summary>
/// A drawing shape, <c>xdr:sp</c> (spec 15). A new shape is written whole by the save; a shape loaded
/// from a file is written back by patching the <c>xdr:sp</c> it was read from, which the save finds
/// by <see cref="Id"/>, with only what was assigned since the load or the last save.
/// </summary>
[DebuggerDisplay("{Name}")]
internal class XLShape : IXLShape
{
    /// <summary>The size a shape is given when it is added: two inches by one at 96 dpi.</summary>
    internal const int DefaultWidth = 192;

    internal const int DefaultHeight = 96;

    private const double MaxLineWidthPt = 1584;

    private string _name = string.Empty;
    private XLPicturePlacement _placement;
    private int _width;
    private int _height;
    private XLColor? _fillColor;
    private bool _noFill;
    private XLColor? _lineColor;
    private double? _lineWidthPt;
    private XLLineDashStyle? _lineDash;
    private bool _noLine;

    /// <summary>What the shape was when it was loaded or last saved; <c>null</c> for one no save has written.</summary>
    private (string Name, XLPictureGeometry Geometry)? _saved;

    internal XLShape(XLWorksheet worksheet, string presetGeometry)
    {
        Worksheet = worksheet;
        PresetGeometry = presetGeometry;
        Body = new XLTextBody();
        Markers = new Dictionary<XLMarkerPosition, XLMarker?>
        {
            [XLMarkerPosition.TopLeft] = null,
            [XLMarkerPosition.BottomRight] = null,
        };
        _placement = XLPicturePlacement.Move;
        _width = DefaultWidth;
        _height = DefaultHeight;
    }

    public int Id { get; internal set; }

    public string Name
    {
        get => _name;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("A shape name cannot be empty or consist only of whitespace.", nameof(value));

            if (_name == value)
                return;

            var oldName = _name;
            _name = value;
            Collection?.OnRenamed(this, oldName);
        }
    }

    /// <summary>
    /// Whether <c>cNvSpPr/@txBox</c> is set. Fixed for the life of the shape, so a shape that is one
    /// is an <see cref="XLTextBox"/>.
    /// </summary>
    internal virtual bool IsTextBox => false;

    public XLShapeType ShapeType => XLShapeTypes.FromPreset(PresetGeometry, IsTextBox);

    public string PresetGeometry { get; }

    public IXLWorksheet Worksheet { get; }

    /// <summary>The collection the shape is in, from when it is added until it is deleted.</summary>
    internal XLShapes? Collection { get; set; }

    /// <summary>
    /// Whether the file the save writes into holds this shape: it was loaded, or a save wrote it. Such
    /// a shape is patched, found by <see cref="Id"/>.
    /// </summary>
    internal bool IsInFile { get; private set; }

    internal IDictionary<XLMarkerPosition, XLMarker?> Markers { get; }

    internal XLShapeProperties Assigned { get; private set; }

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

    public int Width
    {
        get => _width;
        set
        {
            ThrowIfSizedByCells("width");
            ArgumentOutOfRangeException.ThrowIfNegative(value);
            _width = value;
        }
    }

    public int Height
    {
        get => _height;
        set
        {
            ThrowIfSizedByCells("height");
            ArgumentOutOfRangeException.ThrowIfNegative(value);
            _height = value;
        }
    }

    public XLPicturePlacement Placement
    {
        get => _placement;
        set
        {
            if (!Enum.IsDefined(value))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown placement.");

            _placement = value;
        }
    }

    public IXLCell TopLeftCell => Markers[XLMarkerPosition.TopLeft]?.Cell ?? Worksheet.Cell(1, 1);

    public IXLCell BottomRightCell =>
        Markers[XLMarkerPosition.BottomRight]?.Cell
        ?? throw new InvalidOperationException($"Shape '{Name}' has no bottom-right cell; its placement is '{Placement}'.");

    public XLColor? FillColor
    {
        get => _fillColor;
        set
        {
            _fillColor = value;
            Assign(XLShapeProperties.FillColor);
            if (value is not null && _noFill)
            {
                _noFill = false;
                Assign(XLShapeProperties.NoFill);
            }
        }
    }

    public bool NoFill
    {
        get => _noFill;
        set
        {
            _noFill = value;
            Assign(XLShapeProperties.NoFill);
            if (value && _fillColor is not null)
            {
                _fillColor = null;
                Assign(XLShapeProperties.FillColor);
            }
        }
    }

    public XLColor? LineColor
    {
        get => _lineColor;
        set
        {
            _lineColor = value;
            Assign(XLShapeProperties.LineColor);
            if (value is not null && _noLine)
            {
                _noLine = false;
                Assign(XLShapeProperties.NoLine);
            }
        }
    }

    public double? LineWidthPt
    {
        get => _lineWidthPt;
        set
        {
            if (value is { } width && (double.IsNaN(width) || width < 0 || width > MaxLineWidthPt))
                throw new ArgumentOutOfRangeException(nameof(value), value, $"A line is 0 to {MaxLineWidthPt} points wide.");

            _lineWidthPt = value;
            Assign(XLShapeProperties.LineWidth);
        }
    }

    public XLLineDashStyle? LineDash
    {
        get => _lineDash;
        set
        {
            if (value is { } dash && !Enum.IsDefined(dash))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown dash style.");

            _lineDash = value;
            Assign(XLShapeProperties.LineDash);
        }
    }

    public bool NoLine
    {
        get => _noLine;
        set
        {
            _noLine = value;
            Assign(XLShapeProperties.NoLine);
            if (value && _lineColor is not null)
            {
                _lineColor = null;
                Assign(XLShapeProperties.LineColor);
            }
        }
    }

    public IXLTextBody TextBody => Body;

    /// <summary>The text body, as its own type.</summary>
    internal XLTextBody Body { get; }

    public IXLShape MoveTo(int left, int top)
    {
        Placement = XLPicturePlacement.FreeFloating;
        Markers[XLMarkerPosition.BottomRight] = null;
        Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(left, top));
        return this;
    }

    public IXLShape MoveTo(IXLCell cell) => MoveTo(cell, 0, 0);

    public IXLShape MoveTo(IXLCell cell, int xOffset, int yOffset)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ThrowIfOtherSheet(cell);
        Placement = XLPicturePlacement.Move;
        Markers[XLMarkerPosition.BottomRight] = null;
        Markers[XLMarkerPosition.TopLeft] = new XLMarker(cell, new Point(xOffset, yOffset));
        return this;
    }

    public IXLShape MoveTo(IXLCell fromCell, IXLCell toCell)
    {
        ArgumentNullException.ThrowIfNull(fromCell);
        ArgumentNullException.ThrowIfNull(toCell);
        ThrowIfOtherSheet(fromCell);
        ThrowIfOtherSheet(toCell);
        Placement = XLPicturePlacement.MoveAndSize;
        Markers[XLMarkerPosition.TopLeft] = new XLMarker(fromCell);
        Markers[XLMarkerPosition.BottomRight] = new XLMarker(toCell);
        return this;
    }

    public IXLShape WithSize(int width, int height)
    {
        Width = width;
        Height = height;
        return this;
    }

    public IXLShape WithPlacement(XLPicturePlacement value)
    {
        Placement = value;
        return this;
    }

    public void Delete() =>
        (Collection ?? throw new InvalidOperationException($"Shape '{Name}' has already been deleted.")).Delete(this);

    /// <summary>Sets the name without telling the collection, for a shape not yet in it.</summary>
    internal void SetNameUnchecked(string name) => _name = name;

    /// <summary>Sets the size as the loader reads it, whatever the placement.</summary>
    internal void SetLoadedSize(int width, int height)
    {
        _width = width;
        _height = height;
    }

    /// <summary>Sets the placement as the loader reads it, before the shape is in the file.</summary>
    internal void SetLoadedPlacement(XLPicturePlacement placement) => _placement = placement;

    /// <summary>Sets the formatting as the loader reads it, without assigning anything.</summary>
    internal void SetLoadedFormat(XLColor? fillColor, bool noFill, XLColor? lineColor, double? lineWidthPt,
        XLLineDashStyle? lineDash, bool noLine)
    {
        _fillColor = fillColor;
        _noFill = noFill;
        _lineColor = lineColor;
        _lineWidthPt = lineWidthPt;
        _lineDash = lineDash;
        _noLine = noLine;
    }

    /// <summary>
    /// Whether a save has to write this shape: it is new, or something the file holds for it changed
    /// since it was loaded or last saved.
    /// </summary>
    internal bool HasUnsavedChanges =>
        _saved is null || Assigned != XLShapeProperties.None || IsNameChanged || IsGeometryChanged || Body.HasChanges;

    internal bool IsNameChanged => _saved is not { } saved || !string.Equals(saved.Name, _name, StringComparison.Ordinal);

    internal bool IsGeometryChanged => _saved is not { } saved || saved.Geometry != CaptureGeometry();

    internal bool IsSizeChanged =>
        _saved is not { } saved || saved.Geometry.Width != _width || saved.Geometry.Height != _height;

    /// <summary>Whether the top-left corner, its cell or its offset, differs from what was loaded or last saved.</summary>
    internal bool IsTopLeftChanged
    {
        get
        {
            var from = Markers[XLMarkerPosition.TopLeft];
            return _saved is not { } saved || saved.Geometry.From != from?.Anchor || saved.Geometry.FromOffset != from?.Offset;
        }
    }

    /// <summary>Whether the bottom-right corner, its cell or its offset, differs from what was loaded or last saved.</summary>
    internal bool IsBottomRightChanged
    {
        get
        {
            var to = Markers[XLMarkerPosition.BottomRight];
            return _saved is not { } saved || saved.Geometry.To != to?.Anchor || saved.Geometry.ToOffset != to?.Offset;
        }
    }
    /// <summary>
    /// Records the shape as the file now holds it, under drawing id <paramref name="id"/>: once it is
    /// loaded, and once a save that wrote it has succeeded.
    /// </summary>
    internal void MarkSaved(int id)
    {
        if (Id != id)
        {
            var oldId = Id;
            Id = id;
            Collection?.OnIdChanged(oldId, id);
        }

        IsInFile = true;
        Assigned = XLShapeProperties.None;
        _saved = (_name, CaptureGeometry());
        Body.MarkSaved();
    }

    private XLPictureGeometry CaptureGeometry()
    {
        var from = Markers[XLMarkerPosition.TopLeft];
        var to = Markers[XLMarkerPosition.BottomRight];
        return new XLPictureGeometry(_placement, _width, _height, from?.Anchor, from?.Offset, to?.Anchor, to?.Offset);
    }

    private void Assign(XLShapeProperties property) => Assigned |= property;

    private void ThrowIfSizedByCells(string dimension)
    {
        if (Placement == XLPicturePlacement.MoveAndSize)
            throw new ArgumentException(
                $"Cannot set the {dimension} when placement is '{Placement}'. The shape is sized by its cells; move its corners instead.");
    }

    private void ThrowIfOtherSheet(IXLCell cell)
    {
        if (!ReferenceEquals(cell.Worksheet, Worksheet))
            throw new ArgumentException("A shape and its anchor cells must be on the same worksheet.", nameof(cell));
    }
}

/// <summary>A shape whose <c>cNvSpPr/@txBox</c> is set.</summary>
internal sealed class XLTextBox : XLShape, IXLTextBox
{
    internal XLTextBox(XLWorksheet worksheet, string presetGeometry)
        : base(worksheet, presetGeometry)
    {
    }

    internal override bool IsTextBox => true;

    public string Text
    {
        get => TextBody.Text;
        set => TextBody.Text = value;
    }
}
