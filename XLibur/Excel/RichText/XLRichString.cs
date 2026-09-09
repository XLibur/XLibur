using System;
using System.Diagnostics;

namespace XLibur.Excel.RichText;

[DebuggerDisplay("{Text}")]
internal sealed class XLRichString : IXLRichString, IEquatable<XLRichString>
{
    private readonly IXLWithRichString _withRichString;
    private readonly XLFont _font;
    private readonly Action _onChange;
    private string _text;

    public XLRichString(string text, IXLFontBase font, IXLWithRichString withRichString, Action? onChange,
        bool inheritsContainerFont = false, XLStatedRunProperties statedProperties = XLStatedRunProperties.All)
    {
        _text = text;
        _font = new XLFont(font);
        _withRichString = withRichString;
        _onChange = onChange ?? (() => { });
        InheritsContainerFont = inheritsContainerFont;
        StatedProperties = statedProperties;
    }

    /// <summary>
    /// The run carries no formatting of its own and merely reflects the font of its container - it
    /// was read from an <c>&lt;r&gt;</c> with no <c>&lt;rPr&gt;</c>, which per ECMA-376 CT_RElt
    /// inherits the cell font. <see cref="_font"/> holds that inherited font so the run can still be
    /// read and measured, but nothing about it was ever stated by the source, so it must be written
    /// back without an <c>&lt;rPr&gt;</c>. Any change to the font makes the formatting the run's own
    /// and clears this.
    /// </summary>
    internal bool InheritsContainerFont { get; private set; }

    /// <summary>
    /// Which properties of the run's <c>&lt;rPr&gt;</c> the source actually stated. Only the ones a
    /// run would otherwise inherit unnoticed are tracked; see <see cref="XLStatedRunProperties"/>.
    /// Like <see cref="InheritsContainerFont"/>, any change to the font makes the whole formatting
    /// the run's own.
    /// </summary>
    internal XLStatedRunProperties StatedProperties { get; private set; }

    /// <summary>
    /// Records what the source stated, for a run the loader built through the ordinary
    /// <c>AddText</c> path. Must be called after the run's font has been applied, because applying
    /// it marks the formatting as the run's own. The container is notified like it is for any other
    /// change - the cell snapshots its rich text on every notification, and the snapshot taken while
    /// the font was being applied still says the run stated everything.
    /// </summary>
    internal void SetStatedProperties(XLStatedRunProperties statedProperties)
    {
        // Only when it actually changes: the container re-interns its whole rich text on every
        // notification, and a run that stated everything - which is every run XLibur itself wrote -
        // would otherwise churn the shared string table for no change at all.
        if (StatedProperties == statedProperties)
            return;

        StatedProperties = statedProperties;
        _onChange();
    }

    public string Text
    {
        get => _text;
        set
        {
            _text = value;
            _onChange();
        }
    }

    /// <summary>
    /// Signals a change to this run's font. Distinct from a text-only change, because setting any
    /// font property means the run now states its own formatting.
    /// </summary>
    private void OnFontChanged()
    {
        InheritsContainerFont = false;
        StatedProperties = XLStatedRunProperties.All;
        _onChange();
    }

    public IXLRichString AddText(string text)
    {
        return _withRichString.AddText(text);
    }

    public IXLRichString AddNewLine()
    {
        return AddText(Environment.NewLine);
    }

    public bool Bold
    {
        get => _font.Bold;
        set
        {
            _font.Bold = value;
            OnFontChanged();
        }
    }

    public bool Italic
    {
        get => _font.Italic;
        set
        {
            _font.Italic = value;
            OnFontChanged();
        }
    }

    public XLFontUnderlineValues Underline
    {
        get => _font.Underline;
        set
        {
            _font.Underline = value;
            OnFontChanged();
        }
    }

    public bool Strikethrough
    {
        get => _font.Strikethrough;
        set
        {
            _font.Strikethrough = value;
            OnFontChanged();
        }
    }

    public XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => _font.VerticalAlignment;
        set
        {
            _font.VerticalAlignment = value;
            OnFontChanged();
        }
    }

    public bool Shadow
    {
        get => _font.Shadow;
        set
        {
            _font.Shadow = value;
            OnFontChanged();
        }
    }

    public double FontSize
    {
        get => _font.FontSize;
        set
        {
            _font.FontSize = value;
            OnFontChanged();
        }
    }

    public XLColor FontColor
    {
        get => _font.FontColor;
        set
        {
            _font.FontColor = value;
            OnFontChanged();
        }
    }

    public string FontName
    {
        get => _font.FontName;
        set
        {
            _font.FontName = value;
            OnFontChanged();
        }
    }

    public XLFontFamilyNumberingValues FontFamilyNumbering
    {
        get => _font.FontFamilyNumbering;
        set
        {
            _font.FontFamilyNumbering = value;
            OnFontChanged();
        }
    }

    public XLFontCharSet FontCharSet
    {
        get => _font.FontCharSet;
        set
        {
            _font.FontCharSet = value;
            OnFontChanged();
        }
    }

    public XLFontScheme FontScheme
    {
        get => _font.FontScheme;
        set
        {
            _font.FontScheme = value;
            OnFontChanged();
        }
    }

    public IXLRichString SetBold()
    {
        Bold = true; return this;
    }

    public IXLRichString SetBold(bool value)
    {
        Bold = value; return this;
    }

    public IXLRichString SetItalic()
    {
        Italic = true; return this;
    }

    public IXLRichString SetItalic(bool value)
    {
        Italic = value; return this;
    }

    public IXLRichString SetUnderline()
    {
        Underline = XLFontUnderlineValues.Single; return this;
    }

    public IXLRichString SetUnderline(XLFontUnderlineValues value)
    {
        Underline = value; return this;
    }

    public IXLRichString SetStrikethrough()
    {
        Strikethrough = true; return this;
    }

    public IXLRichString SetStrikethrough(bool value)
    {
        Strikethrough = value; return this;
    }

    public IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        VerticalAlignment = value; return this;
    }

    public IXLRichString SetShadow()
    {
        Shadow = true; return this;
    }

    public IXLRichString SetShadow(bool value)
    {
        Shadow = value; return this;
    }

    public IXLRichString SetFontSize(double value)
    {
        FontSize = value; return this;
    }

    public IXLRichString SetFontColor(XLColor value)
    {
        FontColor = value; return this;
    }

    public IXLRichString SetFontName(string value)
    {
        FontName = value; return this;
    }

    public IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        FontFamilyNumbering = value; return this;
    }

    public IXLRichString SetFontCharSet(XLFontCharSet value)
    {
        FontCharSet = value; return this;
    }

    public IXLRichString SetFontScheme(XLFontScheme value)
    {
        FontScheme = value; return this;
    }

    public override bool Equals(object? obj) => Equals(obj as XLRichString);

    public bool Equals(IXLRichString? other) => Equals(other as XLRichString);

    public bool Equals(XLRichString? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return Text == other.Text && _font.Key.Equals(other._font.Key);
    }

    public override int GetHashCode()
    {
        // Since all properties of the type are mutable, can't have different hashcode for any instance.
        // Don't ever use this class in a dictionary, e.g., SST.
        return 4; // Chosen by fair dice roll. Guaranteed to be random.
    }
}
