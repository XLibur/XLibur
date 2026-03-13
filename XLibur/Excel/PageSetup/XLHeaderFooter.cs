using System;
using System.Collections.Generic;
using XLibur.Extensions;

namespace XLibur.Excel;

using System.Linq;

internal sealed class XLHeaderFooter : IXLHeaderFooter
{
    public XLHeaderFooter(XLWorksheet worksheet)
    {
        Worksheet = worksheet;
        Left = new XLHFItem(this);
        Right = new XLHFItem(this);
        Center = new XLHFItem(this);
        SetAsInitial();
    }

    public XLHeaderFooter(XLHeaderFooter defaultHF, XLWorksheet worksheet)
    {
        Worksheet = worksheet;
        defaultHF.innerTexts.ForEach(kp => innerTexts.Add(kp.Key, kp.Value));
        Left = new XLHFItem((XLHFItem)defaultHF.Left, this);
        Center = new XLHFItem((XLHFItem)defaultHF.Center, this);
        Right = new XLHFItem((XLHFItem)defaultHF.Right, this);
        SetAsInitial();
    }

    internal readonly IXLWorksheet Worksheet;

    public IXLHFItem Left { get; private set; }
    public IXLHFItem Center { get; private set; }
    public IXLHFItem Right { get; private set; }

    public string GetText(XLHFOccurrence occurrence)
    {
        //if (innerTexts.ContainsKey(occurrence)) return innerTexts[occurrence];

        var retVal = string.Empty;
        var leftText = Left.GetText(occurrence);
        var centerText = Center.GetText(occurrence);
        var rightText = Right.GetText(occurrence);
        retVal += leftText.Length > 0 ? "&L" + leftText : string.Empty;
        retVal += centerText.Length > 0 ? "&C" + centerText : string.Empty;
        retVal += rightText.Length > 0 ? "&R" + rightText : string.Empty;
        if (retVal.Length > 255)
            throw new ArgumentOutOfRangeException("Headers and Footers cannot be longer than 255 characters (including style markups)");
        return retVal;
    }

    private Dictionary<XLHFOccurrence, string> innerTexts = new();
    internal void SetInnerText(XLHFOccurrence occurrence, string text)
    {
        var parsedElements = ParseFormattedHeaderFooterText(text);

        if (parsedElements.Any(e => e.Position == 'L'))
            Left.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'L').Select(e => e.Text).ToArray()), occurrence);

        if (parsedElements.Any(e => e.Position == 'C'))
            Center.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'C').Select(e => e.Text).ToArray()), occurrence);

        if (parsedElements.Any(e => e.Position == 'R'))
            Right.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'R').Select(e => e.Text).ToArray()), occurrence);

        innerTexts[occurrence] = text;
    }

    private struct ParsedHeaderFooterElement
    {
        public char Position;
        public string Text;
    }

    private static List<ParsedHeaderFooterElement> ParseFormattedHeaderFooterText(string text)
    {
        Func<int, bool> IsAtPositionIndicator = i => i < text.Length - 1 && text[i] == '&' && SourceArray.Contains(text[i + 1]);

        var parsedElements = new List<ParsedHeaderFooterElement>();
        var currentPosition = 'L'; // default is LEFT
        var hfElement = "";

        for (int i = 0; i < text.Length; i++)
        {
            if (IsAtPositionIndicator(i))
            {
                if (hfElement.Length > 0) parsedElements.Add(new ParsedHeaderFooterElement()
                {
                    Position = currentPosition,
                    Text = hfElement
                });

                currentPosition = text[i + 1];
                i += 2;
                hfElement = "";
            }

            if (i < text.Length)
            {
                if (IsAtPositionIndicator(i))
                    i--;
                else
                    hfElement += text[i];
            }
        }

        if (hfElement.Length > 0)
            parsedElements.Add(new ParsedHeaderFooterElement()
            {
                Position = currentPosition,
                Text = hfElement
            });
        return parsedElements;
    }

    private Dictionary<XLHFOccurrence, string> _initialTexts = new();

    internal bool Changed
    {
        get { return field || _initialTexts.Any(it => GetText(it.Key) != it.Value); }
        set;
    }

    internal static readonly char[] SourceArray = ['L', 'C', 'R'];

    internal void SetAsInitial()
    {
        _initialTexts = new Dictionary<XLHFOccurrence, string>();
        foreach (var o in Enum.GetValues(typeof(XLHFOccurrence)).Cast<XLHFOccurrence>())
        {
            _initialTexts.Add(o, GetText(o));
        }
    }


    public IXLHeaderFooter Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
    {
        Left.Clear(occurrence);
        Right.Clear(occurrence);
        Center.Clear(occurrence);
        return this;
    }

    /// <summary>
    /// Returns true if any left/center/right item has images for any occurrence.
    /// </summary>
    internal bool HasImages =>
        ((XLHFItem)Left).HasImages ||
        ((XLHFItem)Center).HasImages ||
        ((XLHFItem)Right).HasImages;

    /// <summary>
    /// Collects all images across all occurrences, keyed by their VML position code
    /// (e.g. "LH", "CH", "RH" for header or "LF", "CF", "RF" for footer).
    /// </summary>
    /// <param name="suffix">"H" for header, "F" for footer.</param>
    internal List<XLHFImage> CollectImages(string suffix)
    {
        var result = new List<XLHFImage>();
        var items = new[] { ('L', (XLHFItem)Left), ('C', (XLHFItem)Center), ('R', (XLHFItem)Right) };
        var occurrences = new[] { XLHFOccurrence.OddPages, XLHFOccurrence.EvenPages, XLHFOccurrence.FirstPage };

        foreach (var (posChar, item) in items)
        {
            foreach (var occ in occurrences)
            {
                var image = item.GetImage(occ);
                if (image != null)
                {
                    image.PositionCode = $"{posChar}{suffix}";
                    if (!result.Contains(image))
                        result.Add(image);
                }
            }
        }

        return result;
    }
}
