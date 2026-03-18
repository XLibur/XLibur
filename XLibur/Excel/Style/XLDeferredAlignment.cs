using System;

namespace XLibur.Excel;

/// <summary>
/// Lightweight <see cref="IXLAlignment"/> that accumulates property changes into an <see cref="XLAlignmentKey"/>
/// without triggering repository lookups or style-slice writes. Used by <see cref="XLDeferredStyle"/>
/// to support batch style mutations.
/// </summary>
internal sealed class XLDeferredAlignment : IXLAlignment
{
    private readonly XLDeferredStyle _style;
    internal XLAlignmentKey Key;

    internal XLDeferredAlignment(XLDeferredStyle style, XLAlignmentKey key)
    {
        _style = style;
        Key = key;
    }

    public XLAlignmentHorizontalValues Horizontal
    {
        get => Key.Horizontal;
        set
        {
            Key = Key with { Horizontal = value };

            bool updateIndent = !(
                value == XLAlignmentHorizontalValues.Left
                || value == XLAlignmentHorizontalValues.Right
                || value == XLAlignmentHorizontalValues.Distributed
            );
            if (updateIndent)
                Key = Key with { Indent = 0 };
        }
    }

    public XLAlignmentVerticalValues Vertical
    {
        get => Key.Vertical;
        set => Key = Key with { Vertical = value };
    }

    public int Indent
    {
        get => Key.Indent;
        set
        {
            if (Indent != value)
            {
                if (Horizontal == XLAlignmentHorizontalValues.General)
                    Horizontal = XLAlignmentHorizontalValues.Left;

                if (value > 0 && !(
                        Horizontal == XLAlignmentHorizontalValues.Left
                        || Horizontal == XLAlignmentHorizontalValues.Right
                        || Horizontal == XLAlignmentHorizontalValues.Distributed
                    ))
                {
                    throw new ArgumentException(
                        "For indents, only left, right, and distributed horizontal alignments are supported.");
                }
            }
            Key = Key with { Indent = value };
        }
    }

    public bool JustifyLastLine
    {
        get => Key.JustifyLastLine;
        set => Key = Key with { JustifyLastLine = value };
    }

    public XLAlignmentReadingOrderValues ReadingOrder
    {
        get => Key.ReadingOrder;
        set => Key = Key with { ReadingOrder = value };
    }

    public int RelativeIndent
    {
        get => Key.RelativeIndent;
        set => Key = Key with { RelativeIndent = value };
    }

    public bool ShrinkToFit
    {
        get => Key.ShrinkToFit;
        set => Key = Key with { ShrinkToFit = value };
    }

    public int TextRotation
    {
        get => Key.TextRotation;
        set
        {
            if (value != 255 && (value < -90 || value > 90))
                throw new ArgumentException("TextRotation must be between -90 and 90 degrees, or 255.");
            Key = Key with { TextRotation = value };
        }
    }

    public bool WrapText
    {
        get => Key.WrapText;
        set => Key = Key with { WrapText = value };
    }

    public bool TopToBottom
    {
        get => TextRotation == 255;
        set => TextRotation = value ? 255 : 0;
    }

    public IXLStyle SetHorizontal(XLAlignmentHorizontalValues value) { Horizontal = value; return _style; }
    public IXLStyle SetVertical(XLAlignmentVerticalValues value) { Vertical = value; return _style; }
    public IXLStyle SetIndent(int value) { Indent = value; return _style; }
    public IXLStyle SetJustifyLastLine() { JustifyLastLine = true; return _style; }
    public IXLStyle SetJustifyLastLine(bool value) { JustifyLastLine = value; return _style; }
    public IXLStyle SetReadingOrder(XLAlignmentReadingOrderValues value) { ReadingOrder = value; return _style; }
    public IXLStyle SetRelativeIndent(int value) { RelativeIndent = value; return _style; }
    public IXLStyle SetShrinkToFit() { ShrinkToFit = true; return _style; }
    public IXLStyle SetShrinkToFit(bool value) { ShrinkToFit = value; return _style; }
    public IXLStyle SetTextRotation(int value) { TextRotation = value; return _style; }
    public IXLStyle SetWrapText() { WrapText = true; return _style; }
    public IXLStyle SetWrapText(bool value) { WrapText = value; return _style; }
    public IXLStyle SetTopToBottom() { TopToBottom = true; return _style; }
    public IXLStyle SetTopToBottom(bool value) { TopToBottom = value; return _style; }

    public bool Equals(IXLAlignment? other) => other is XLDeferredAlignment da ? Key == da.Key : false;
}
