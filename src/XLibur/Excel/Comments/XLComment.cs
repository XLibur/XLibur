using System;
using XLibur.Excel.RichText;

namespace XLibur.Excel;

internal sealed class XLComment : XLFormattedText<IXLComment>, IXLComment
{
    /// <summary>
    /// The cell the note was last known to sit on. Only a hint for <see cref="Delete"/>: shifting
    /// rows or columns moves the note's entry within the misc slice without telling the note, so the
    /// address can name a cell the note has since moved off. The worksheet behind the reference is
    /// not a hint — a note never changes sheet, because copying builds a new one.
    /// </summary>
    private XLCell _lastKnownCell = null!;

    public XLComment(XLCell cell, IXLFontBase? defaultFont = null, int? shapeId = null)
        : base(defaultFont ?? XLFont.DefaultCommentFont)
    {
        Initialize(cell, shapeId: shapeId);
    }

    public XLComment(XLCell cell, XLFormattedText<IXLComment> defaultComment, IXLFontBase defaultFont,
        IXLDrawingStyle style)
        : base(defaultComment, defaultFont)
    {
        Initialize(cell, style);
    }

    public XLComment(XLCell cell, string text, IXLFontBase defaultFont)
        : base(text, defaultFont)
    {
        Initialize(cell);
    }

    #region IXLComment Members

    public string Author { get; set; } = string.Empty;

    public IXLComment SetAuthor(string value)
    {
        Author = value;
        return this;
    }

    public IXLRichString AddSignature()
    {
        AddText(Author + ":").SetBold();
        return AddText(Environment.NewLine);
    }

    public void Delete()
    {
        // Confirm the note is still the one living at the remembered address before clearing it: a
        // shift may have moved this note elsewhere, or moved another note in. The sheet's notes are
        // only walked when that hint has gone stale, so the common case stays a single slice read.
        if (ReferenceEquals(_lastKnownCell.SliceComment, this))
        {
            _lastKnownCell.SliceComment = null;
            return;
        }

        var cells = _lastKnownCell.Worksheet.Internals.CellsCollection;
        if (cells.FindNote(this) is { } point)
            cells.GetCell(point).SliceComment = null;
    }

    #endregion IXLComment Members

    #region IXLDrawing

    public string Name { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    /// <summary>
    /// How the note is tied to the grid. The same value as
    /// <c>Style.Properties.Positioning</c> and stored there — one field, so the two can no longer
    /// disagree the way they did before D17 was fixed. This is the name the writers and the shift
    /// listener read; the style property is the name the public API exposes.
    /// </summary>
    public XLDrawingAnchor Anchor
    {
        get => Style.Properties.Positioning;
        set => Style.Properties.Positioning = value;
    }

    public bool HorizontalFlip { get; set; }

    public bool VerticalFlip { get; set; }

    public int Rotation { get; set; }

    public int ExtentLength { get; set; }

    public int ExtentWidth { get; set; }

    public int ShapeId { get; internal set; }

    public bool Visible { get; set; }

    public IXLComment SetVisible()
    {
        Visible = true;
        return Container;
    }

    public IXLComment SetVisible(bool hidden)
    {
        Visible = hidden;
        return Container;
    }

    public IXLDrawingPosition Position { get; private set; } = null!;

    public int ZOrder { get; set; }

    public IXLComment SetZOrder(int zOrder)
    {
        ZOrder = zOrder;
        return Container;
    }

    public IXLDrawingStyle Style { get; private set; } = null!;

    public IXLComment SetName(string name)
    {
        Name = name;
        return Container;
    }

    public IXLComment SetDescription(string description)
    {
        Description = description;
        return Container;
    }

    public IXLComment SetHorizontalFlip()
    {
        HorizontalFlip = true;
        return Container;
    }

    public IXLComment SetHorizontalFlip(bool horizontalFlip)
    {
        HorizontalFlip = horizontalFlip;
        return Container;
    }

    public IXLComment SetVerticalFlip()
    {
        VerticalFlip = true;
        return Container;
    }

    public IXLComment SetVerticalFlip(bool verticalFlip)
    {
        VerticalFlip = verticalFlip;
        return Container;
    }

    public IXLComment SetRotation(int rotation)
    {
        Rotation = rotation;
        return Container;
    }

    public IXLComment SetExtentLength(int extentLength)
    {
        ExtentLength = extentLength;
        return Container;
    }

    public IXLComment SetExtentWidth(int extentWidth)
    {
        ExtentWidth = extentWidth;
        return Container;
    }

    #endregion IXLDrawing

    private void Initialize(XLCell cell, IXLDrawingStyle? style = null, int? shapeId = null)
    {
        style ??= XLDrawingStyle.DefaultCommentStyle;
        shapeId ??= cell.Worksheet.Workbook.ShapeIdManager.GetNext();

        Author = cell.Worksheet.Author;
        Container = this;
        Style = new XLDrawingStyle();
        var previousRowNumber = cell.Address.RowNumber;
        double previousRowOffset = 0;

        if (previousRowNumber > 1)
        {
            previousRowNumber--;

            previousRowOffset =
                cell.Worksheet.Internals.RowsCollection.TryGetValue(previousRowNumber, out var previousRow)
                    ? Math.Max(0, previousRow.Height - 7)
                    : Math.Max(0, cell.Worksheet.RowHeight - 7);
        }

        Position = new XLDrawingPosition
        {
            Column = cell.Address.ColumnNumber + 1,
            ColumnOffset = 2,
            Row = previousRowNumber,
            RowOffset = previousRowOffset
        };

        ZOrder = cell.Worksheet.ZOrder++;
        Style
            .Margins.SetLeft(style.Margins.Left)
            .Margins.SetRight(style.Margins.Right)
            .Margins.SetTop(style.Margins.Top)
            .Margins.SetBottom(style.Margins.Bottom)
            .Margins.SetAutomatic(style.Margins.Automatic)
            .Size.SetHeight(style.Size.Height)
            .Size.SetWidth(style.Size.Width)
            .ColorsAndLines.SetLineColor(style.ColorsAndLines.LineColor)
            .ColorsAndLines.SetFillColor(style.ColorsAndLines.FillColor)
            .ColorsAndLines.SetLineDash(style.ColorsAndLines.LineDash)
            .ColorsAndLines.SetLineStyle(style.ColorsAndLines.LineStyle)
            .ColorsAndLines.SetLineWeight(style.ColorsAndLines.LineWeight)
            .ColorsAndLines.SetFillTransparency(style.ColorsAndLines.FillTransparency)
            .ColorsAndLines.SetLineTransparency(style.ColorsAndLines.LineTransparency)
            .Alignment.SetHorizontal(style.Alignment.Horizontal)
            .Alignment.SetVertical(style.Alignment.Vertical)
            .Alignment.SetDirection(style.Alignment.Direction)
            .Alignment.SetOrientation(style.Alignment.Orientation)
            .Alignment.SetAutomaticSize(style.Alignment.AutomaticSize)
            .Properties.SetPositioning(style.Properties.Positioning)
            .Protection.SetLocked(style.Protection.Locked)
            .Protection.SetLockText(style.Protection.LockText);

        _lastKnownCell = cell;
        ShapeId = shapeId.Value;
    }
}
