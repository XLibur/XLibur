using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel.Drawings;

/// <summary>
/// The text of a shape. Each paragraph and run that the file holds remembers where it is in the
/// body, so a save edits it there and nothing beside it (spec 15, design rule 8).
/// </summary>
/// <remarks>
/// Only <see cref="Text"/>'s setter and <see cref="Clear"/> rebuild the body, and they are the
/// edits that say so. A body that holds a single paragraph with no text, as Excel writes an empty
/// one, reads as no paragraph at all; the save then writes the first paragraph added into it.
/// </remarks>
internal sealed class XLTextBody : IXLTextBody
{
    private readonly List<XLTextParagraph> _paragraphs = [];
    private XLTextVerticalAlignment _verticalAlignment;
    private bool _wrapText = true;

    public int Count => _paragraphs.Count;

    public IXLTextParagraph this[int index] => _paragraphs[index];

    internal IReadOnlyList<XLTextParagraph> Paragraphs => _paragraphs;

    /// <summary>Whether the paragraphs were replaced wholesale since the load or the last save.</summary>
    internal bool IsRebuilt { get; private set; }

    /// <summary>
    /// Whether the file holds one empty paragraph that the model does not show, which the first
    /// paragraph added is written into.
    /// </summary>
    internal bool HasPlaceholderParagraph { get; set; }

    internal bool IsVerticalAlignmentAssigned { get; private set; }

    internal bool IsWrapTextAssigned { get; private set; }

    internal bool HasChanges =>
        IsRebuilt || IsVerticalAlignmentAssigned || IsWrapTextAssigned || _paragraphs.Exists(p => p.HasChanges);

    public XLTextVerticalAlignment VerticalAlignment
    {
        get => _verticalAlignment;
        set
        {
            if (!Enum.IsDefined(value))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown vertical alignment.");

            _verticalAlignment = value;
            IsVerticalAlignmentAssigned = true;
        }
    }

    public bool WrapText
    {
        get => _wrapText;
        set
        {
            _wrapText = value;
            IsWrapTextAssigned = true;
        }
    }

    public string Text
    {
        get => string.Join("\n", _paragraphs.Select(p => p.Text));
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            ThrowIfNotWritable(value, nameof(value));
            Clear();
            if (value.Length == 0)
                return;

            foreach (var line in Lines(value))
                AddParagraph(line);
        }
    }

    /// <summary>Throws when <paramref name="text"/>, a whole body's text, has a line a paragraph cannot hold.</summary>
    internal static void ThrowIfNotWritable(string text, string paramName)
    {
        foreach (var line in Lines(text))
            XLTextRun.ThrowIfNotWritable(line, paramName);
    }

    private static string[] Lines(string text) => text.Replace("\r\n", "\n").Split('\n');

    public IXLTextParagraph AddParagraph(string text = "")
    {
        ArgumentNullException.ThrowIfNull(text);
        XLTextRun.ThrowIfNotWritable(text, nameof(text));
        var paragraph = new XLTextParagraph(domIndex: null);
        if (text.Length > 0)
            paragraph.AddText(text);

        _paragraphs.Add(paragraph);
        return paragraph;
    }

    public void Clear()
    {
        _paragraphs.Clear();
        IsRebuilt = true;
    }

    public IEnumerator<IXLTextParagraph> GetEnumerator() => _paragraphs.Cast<IXLTextParagraph>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Adds a paragraph the loader read, the <paramref name="domIndex"/>-th <c>a:p</c> of the body.</summary>
    internal XLTextParagraph AddLoadedParagraph(int domIndex)
    {
        var paragraph = new XLTextParagraph(domIndex);
        _paragraphs.Add(paragraph);
        return paragraph;
    }

    /// <summary>Sets the body properties as the loader reads them, without assigning anything.</summary>
    internal void SetLoaded(XLTextVerticalAlignment verticalAlignment, bool wrapText)
    {
        _verticalAlignment = verticalAlignment;
        _wrapText = wrapText;
    }

    /// <summary>
    /// Makes this empty body a copy of <paramref name="source"/>: its paragraphs and runs, where the
    /// file holds each, and every edit not saved yet, so a copy of the shape written from the
    /// shape's XML comes out as the shape would.
    /// </summary>
    internal void CopyFrom(XLTextBody source)
    {
        foreach (var paragraph in source._paragraphs)
            _paragraphs.Add(paragraph.Copy());

        _verticalAlignment = source._verticalAlignment;
        _wrapText = source._wrapText;
        IsRebuilt = source.IsRebuilt;
        HasPlaceholderParagraph = source.HasPlaceholderParagraph;
        IsVerticalAlignmentAssigned = source.IsVerticalAlignmentAssigned;
        IsWrapTextAssigned = source.IsWrapTextAssigned;
    }

    /// <summary>
    /// Records the body as a save wrote it: the paragraphs are the body's, in order, and an empty
    /// body is one empty paragraph.
    /// </summary>
    internal void MarkSaved()
    {
        for (var i = 0; i < _paragraphs.Count; i++)
            _paragraphs[i].MarkSaved(i);

        HasPlaceholderParagraph = _paragraphs.Count == 0;
        IsRebuilt = false;
        IsVerticalAlignmentAssigned = false;
        IsWrapTextAssigned = false;
    }
}

/// <summary>A paragraph of shape text, <c>a:p</c>.</summary>
internal sealed class XLTextParagraph : IXLTextParagraph
{
    // The runs, and the text of the line breaks and fields between them, in order. Only the runs
    // are the paragraph's IXLTextRun items; the others are kept so its text reads as the file's.
    private readonly List<object> _content = [];
    private readonly List<XLTextRun> _runs = [];
    private XLTextAlignment _alignment;

    internal XLTextParagraph(int? domIndex) => DomIndex = domIndex;

    /// <summary>Which <c>a:p</c> of the body the file holds this paragraph as, or <c>null</c> for a new one.</summary>
    internal int? DomIndex { get; private set; }

    internal bool IsAlignmentAssigned { get; private set; }

    internal IReadOnlyList<XLTextRun> Runs => _runs;

    internal bool HasChanges => DomIndex is null || IsAlignmentAssigned || _runs.Exists(r => r.HasChanges);

    public XLTextAlignment Alignment
    {
        get => _alignment;
        set
        {
            if (!Enum.IsDefined(value))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown alignment.");

            _alignment = value;
            IsAlignmentAssigned = true;
        }
    }

    public string Text => string.Concat(_content.Select(c => c is XLTextRun run ? run.Text : (string)c));

    public int Count => _runs.Count;

    public IXLTextRun this[int index] => _runs[index];

    public IXLTextRun AddText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        XLTextRun.ThrowIfNotWritable(text, nameof(text));
        var run = new XLTextRun(text, domIndex: null);
        _runs.Add(run);
        _content.Add(run);
        return run;
    }

    public IEnumerator<IXLTextRun> GetEnumerator() => _runs.Cast<IXLTextRun>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Sets the alignment as the loader reads it, without assigning it.</summary>
    internal void SetLoadedAlignment(XLTextAlignment alignment) => _alignment = alignment;

    /// <summary>Adds a run the loader read, the <paramref name="domIndex"/>-th <c>a:r</c> of the paragraph.</summary>
    internal XLTextRun AddLoadedRun(string text, int domIndex)
    {
        var run = new XLTextRun(text, domIndex);
        _runs.Add(run);
        _content.Add(run);
        return run;
    }

    /// <summary>Adds the text of a line break or a field the loader read, which is not a run.</summary>
    internal void AddLoadedText(string text) => _content.Add(text);

    internal XLTextParagraph Copy()
    {
        var copy = new XLTextParagraph(DomIndex)
        {
            _alignment = _alignment,
            IsAlignmentAssigned = IsAlignmentAssigned,
        };
        foreach (var content in _content)
        {
            if (content is XLTextRun run)
            {
                var runCopy = run.Copy();
                copy._runs.Add(runCopy);
                copy._content.Add(runCopy);
            }
            else
            {
                copy._content.Add(content);
            }
        }

        return copy;
    }

    internal void MarkSaved(int domIndex)
    {
        DomIndex = domIndex;
        IsAlignmentAssigned = false;
        for (var i = 0; i < _runs.Count; i++)
            _runs[i].MarkSaved(i);
    }
}

/// <summary>The font properties of a run a save writes, when they were assigned.</summary>
[Flags]
internal enum XLTextRunProperties
{
    None = 0,
    Text = 1 << 0,
    Bold = 1 << 1,
    Italic = 1 << 2,
    Underline = 1 << 3,
    Strikethrough = 1 << 4,
    FontSize = 1 << 5,
    FontColor = 1 << 6,
    FontName = 1 << 7,
}

/// <summary>A run of shape text, <c>a:r</c>.</summary>
internal sealed class XLTextRun : IXLTextRun
{
    /// <summary>The sizes <c>a:rPr/@sz</c> can hold, in points.</summary>
    private const double MinFontSize = 1;

    private const double MaxFontSize = 4000;

    private string _text;
    private bool? _bold;
    private bool? _italic;
    private XLShapeTextUnderline? _underline;
    private bool? _strikethrough;
    private double? _fontSize;
    private XLColor? _fontColor;
    private string? _fontName;

    internal XLTextRun(string text, int? domIndex)
    {
        _text = text;
        DomIndex = domIndex;
    }

    /// <summary>
    /// Throws when <paramref name="text"/> has a character an <c>a:t</c> cannot hold: a control
    /// character other than a tab, which includes the <c>\v</c> a line break reads as, and a line
    /// break, which starts a new paragraph instead. Checked where the text is set, not at the save.
    /// </summary>
    internal static void ThrowIfNotWritable(string text, string paramName)
    {
        foreach (var c in text)
        {
            if ((c < ' ' && c != '\t') || c is '￾' or '￿')
                throw new ArgumentException(
                    $"Shape text cannot hold the character U+{(int)c:X4}. Start a new paragraph for a line break.", paramName);
        }
    }

    /// <summary>Which <c>a:r</c> of its paragraph the file holds this run as, or <c>null</c> for a new one.</summary>
    internal int? DomIndex { get; private set; }

    internal XLTextRunProperties Assigned { get; private set; }

    internal bool HasChanges => DomIndex is null || Assigned != XLTextRunProperties.None;

    public string Text
    {
        get => _text;
        set
        {
            ArgumentNullException.ThrowIfNull(value);
            ThrowIfNotWritable(value, nameof(value));
            _text = value;
            Assigned |= XLTextRunProperties.Text;
        }
    }

    public bool? Bold
    {
        get => _bold;
        set
        {
            _bold = value;
            Assigned |= XLTextRunProperties.Bold;
        }
    }

    public bool? Italic
    {
        get => _italic;
        set
        {
            _italic = value;
            Assigned |= XLTextRunProperties.Italic;
        }
    }

    public XLShapeTextUnderline? Underline
    {
        get => _underline;
        set
        {
            if (value is { } underline && !Enum.IsDefined(underline))
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown underline.");

            _underline = value;
            Assigned |= XLTextRunProperties.Underline;
        }
    }

    public bool? Strikethrough
    {
        get => _strikethrough;
        set
        {
            _strikethrough = value;
            Assigned |= XLTextRunProperties.Strikethrough;
        }
    }

    public double? FontSize
    {
        get => _fontSize;
        set
        {
            if (value is { } size && (double.IsNaN(size) || size < MinFontSize || size > MaxFontSize))
                throw new ArgumentOutOfRangeException(nameof(value), value, $"A font size is {MinFontSize} to {MaxFontSize} points.");

            _fontSize = value;
            Assigned |= XLTextRunProperties.FontSize;
        }
    }

    public XLColor? FontColor
    {
        get => _fontColor;
        set
        {
            _fontColor = value;
            Assigned |= XLTextRunProperties.FontColor;
        }
    }

    public string? FontName
    {
        get => _fontName;
        set
        {
            if (value is { } name && string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("A font name cannot be empty or consist only of whitespace.", nameof(value));

            _fontName = value;
            Assigned |= XLTextRunProperties.FontName;
        }
    }

    public IXLTextRun SetBold(bool value = true)
    {
        Bold = value;
        return this;
    }

    public IXLTextRun SetItalic(bool value = true)
    {
        Italic = value;
        return this;
    }

    public IXLTextRun SetFontSize(double value)
    {
        FontSize = value;
        return this;
    }

    public IXLTextRun SetFontColor(XLColor value)
    {
        FontColor = value;
        return this;
    }

    public IXLTextRun SetFontName(string value)
    {
        FontName = value;
        return this;
    }

    /// <summary>Sets the font as the loader reads it, without assigning anything.</summary>
    internal void SetLoadedFont(bool? bold, bool? italic, XLShapeTextUnderline? underline, bool? strikethrough,
        double? fontSize, XLColor? fontColor, string? fontName)
    {
        _bold = bold;
        _italic = italic;
        _underline = underline;
        _strikethrough = strikethrough;
        _fontSize = fontSize;
        _fontColor = fontColor;
        _fontName = fontName;
    }

    internal XLTextRun Copy() => new(_text, DomIndex)
    {
        _bold = _bold,
        _italic = _italic,
        _underline = _underline,
        _strikethrough = _strikethrough,
        _fontSize = _fontSize,
        _fontColor = _fontColor,
        _fontName = _fontName,
        Assigned = Assigned,
    };

    internal void MarkSaved(int domIndex)
    {
        DomIndex = domIndex;
        Assigned = XLTextRunProperties.None;
    }
}
