using System;

namespace XLibur.Excel;

/// <summary>
/// Lightweight <see cref="IXLBorder"/> that accumulates property changes into an <see cref="XLBorderKey"/>
/// without triggering repository lookups or style-slice writes. Used by <see cref="XLDeferredStyle"/>
/// to support batch style mutations.
/// </summary>
internal sealed class XLDeferredBorder : IXLBorder
{
    private readonly XLDeferredStyle _style;
    internal XLBorderKey Key;

    internal XLDeferredBorder(XLDeferredStyle style, XLBorderKey key)
    {
        _style = style;
        Key = key;
    }

    public XLBorderStyleValues OutsideBorder
    {
        set => Key = Key with
        {
            TopBorder = value,
            BottomBorder = value,
            LeftBorder = value,
            RightBorder = value,
        };
    }

    public XLColor OutsideBorderColor
    {
        set => Key = Key with
        {
            TopBorderColor = value.Key,
            BottomBorderColor = value.Key,
            LeftBorderColor = value.Key,
            RightBorderColor = value.Key,
        };
    }

    public XLBorderStyleValues InsideBorder
    {
        set => Key = Key with
        {
            TopBorder = value,
            BottomBorder = value,
            LeftBorder = value,
            RightBorder = value,
        };
    }

    public XLColor InsideBorderColor
    {
        set => Key = Key with
        {
            TopBorderColor = value.Key,
            BottomBorderColor = value.Key,
            LeftBorderColor = value.Key,
            RightBorderColor = value.Key,
        };
    }

    public XLBorderStyleValues LeftBorder
    {
        get => Key.LeftBorder;
        set => Key = Key with { LeftBorder = value };
    }

    public XLColor LeftBorderColor
    {
        get
        {
            var colorKey = Key.LeftBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { LeftBorderColor = value.Key };
        }
    }

    public XLBorderStyleValues RightBorder
    {
        get => Key.RightBorder;
        set => Key = Key with { RightBorder = value };
    }

    public XLColor RightBorderColor
    {
        get
        {
            var colorKey = Key.RightBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { RightBorderColor = value.Key };
        }
    }

    public XLBorderStyleValues TopBorder
    {
        get => Key.TopBorder;
        set => Key = Key with { TopBorder = value };
    }

    public XLColor TopBorderColor
    {
        get
        {
            var colorKey = Key.TopBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { TopBorderColor = value.Key };
        }
    }

    public XLBorderStyleValues BottomBorder
    {
        get => Key.BottomBorder;
        set => Key = Key with { BottomBorder = value };
    }

    public XLColor BottomBorderColor
    {
        get
        {
            var colorKey = Key.BottomBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { BottomBorderColor = value.Key };
        }
    }

    public bool DiagonalUp
    {
        get => Key.DiagonalUp;
        set => Key = Key with { DiagonalUp = value };
    }

    public bool DiagonalDown
    {
        get => Key.DiagonalDown;
        set => Key = Key with { DiagonalDown = value };
    }

    public XLBorderStyleValues DiagonalBorder
    {
        get => Key.DiagonalBorder;
        set => Key = Key with { DiagonalBorder = value };
    }

    public XLColor DiagonalBorderColor
    {
        get
        {
            var colorKey = Key.DiagonalBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { DiagonalBorderColor = value.Key };
        }
    }

    public IXLStyle SetOutsideBorder(XLBorderStyleValues value) { OutsideBorder = value; return _style; }
    public IXLStyle SetOutsideBorderColor(XLColor value) { OutsideBorderColor = value; return _style; }
    public IXLStyle SetInsideBorder(XLBorderStyleValues value) { InsideBorder = value; return _style; }
    public IXLStyle SetInsideBorderColor(XLColor value) { InsideBorderColor = value; return _style; }
    public IXLStyle SetLeftBorder(XLBorderStyleValues value) { LeftBorder = value; return _style; }
    public IXLStyle SetLeftBorderColor(XLColor value) { LeftBorderColor = value; return _style; }
    public IXLStyle SetRightBorder(XLBorderStyleValues value) { RightBorder = value; return _style; }
    public IXLStyle SetRightBorderColor(XLColor value) { RightBorderColor = value; return _style; }
    public IXLStyle SetTopBorder(XLBorderStyleValues value) { TopBorder = value; return _style; }
    public IXLStyle SetTopBorderColor(XLColor value) { TopBorderColor = value; return _style; }
    public IXLStyle SetBottomBorder(XLBorderStyleValues value) { BottomBorder = value; return _style; }
    public IXLStyle SetBottomBorderColor(XLColor value) { BottomBorderColor = value; return _style; }
    public IXLStyle SetDiagonalUp() { DiagonalUp = true; return _style; }
    public IXLStyle SetDiagonalUp(bool value) { DiagonalUp = value; return _style; }
    public IXLStyle SetDiagonalDown() { DiagonalDown = true; return _style; }
    public IXLStyle SetDiagonalDown(bool value) { DiagonalDown = value; return _style; }
    public IXLStyle SetDiagonalBorder(XLBorderStyleValues value) { DiagonalBorder = value; return _style; }
    public IXLStyle SetDiagonalBorderColor(XLColor value) { DiagonalBorderColor = value; return _style; }

    public bool Equals(IXLBorder? other) => other is XLDeferredBorder db ? Key == db.Key : false;
    public override bool Equals(object? obj) => Equals(obj as IXLBorder);
    public override int GetHashCode() => Key.GetHashCode();
}
