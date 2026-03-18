using System;
using System.Collections.Generic;
using System.Text;
using XLibur.Excel.RichText;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLHFItem : IXLHFItem
{
    internal readonly XLHeaderFooter HeaderFooter;
    public XLHFItem(XLHeaderFooter headerFooter)
    {
        HeaderFooter = headerFooter;
    }
    public XLHFItem(XLHFItem defaultHFItem, XLHeaderFooter headerFooter)
        : this(headerFooter)
    {
        defaultHFItem.texts.ForEach(kp => texts.Add(kp.Key, kp.Value));
    }
    private readonly Dictionary<XLHFOccurrence, List<XLHFText>> texts = new Dictionary<XLHFOccurrence, List<XLHFText>>();

    /// <summary>
    /// Images keyed by occurrence. Each occurrence has at most one image per HFItem (left/center/right).
    /// </summary>
    private readonly Dictionary<XLHFOccurrence, XLHFImage> _images = new();

    public string GetText(XLHFOccurrence occurrence)
    {
        var sb = new StringBuilder();
        if (texts.TryGetValue(occurrence, out var hfTexts))
        {
            foreach (var hfText in hfTexts)
                sb.Append(hfText.GetHFText(sb.ToString()));
        }

        return sb.ToString();
    }

    public IXLRichString AddText(string text)
    {
        return AddText(text, XLHFOccurrence.AllPages);
    }
    public IXLRichString AddText(XLHFPredefinedText predefinedText)
    {
        return AddText(predefinedText, XLHFOccurrence.AllPages);
    }

    public IXLRichString AddText(string text, XLHFOccurrence occurrence)
    {
        var richText = new XLRichString(text, HeaderFooter.Worksheet.Style.Font, this, null);

        var hfText = new XLHFText(richText, this);
        if (occurrence == XLHFOccurrence.AllPages)
        {
            AddTextToOccurrence(hfText, XLHFOccurrence.EvenPages);
            AddTextToOccurrence(hfText, XLHFOccurrence.FirstPage);
            AddTextToOccurrence(hfText, XLHFOccurrence.OddPages);
        }
        else
        {
            AddTextToOccurrence(hfText, occurrence);
        }

        return richText;
    }

    public IXLRichString AddNewLine()
    {
        return AddText(Environment.NewLine);
    }

    public IXLRichString AddImage(string imagePath, XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
    {
        var image = XLHFImage.FromFile(imagePath, (XLWorkbook)HeaderFooter.Worksheet.Workbook);

        if (occurrence == XLHFOccurrence.AllPages)
        {
            AddImageToOccurrence(image, XLHFOccurrence.EvenPages);
            AddImageToOccurrence(image, XLHFOccurrence.FirstPage);
            AddImageToOccurrence(image, XLHFOccurrence.OddPages);
        }
        else
        {
            AddImageToOccurrence(image, occurrence);
        }

        // Insert the &G marker into the text stream at the current position
        return AddText("&G", occurrence);
    }

    private void AddImageToOccurrence(XLHFImage image, XLHFOccurrence occurrence)
    {
        _images[occurrence] = image;
        HeaderFooter.Changed = true;
    }

    /// <summary>
    /// Gets the image for a specific occurrence, or null if none.
    /// </summary>
    internal XLHFImage? GetImage(XLHFOccurrence occurrence)
    {
        return _images.TryGetValue(occurrence, out var image) ? image : null;
    }

    /// <summary>
    /// Returns true if any occurrence has an image.
    /// </summary>
    internal bool HasImages => _images.Count > 0;

    private void AddTextToOccurrence(XLHFText hfText, XLHFOccurrence occurrence)
    {
        if (texts.TryGetValue(occurrence, out var hfTexts))
            hfTexts.Add(hfText);
        else
            texts.Add(occurrence, [hfText]);

        HeaderFooter.Changed = true;
    }

    public IXLRichString AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence)
    {
        var hfText = predefinedText switch
        {
            XLHFPredefinedText.PageNumber => "&P",
            XLHFPredefinedText.NumberOfPages => "&N",
            XLHFPredefinedText.Date => "&D",
            XLHFPredefinedText.Time => "&T",
            XLHFPredefinedText.Path => "&Z",
            XLHFPredefinedText.File => "&F",
            XLHFPredefinedText.SheetName => "&A",
            XLHFPredefinedText.FullPath => "&Z&F",
            _ => throw new NotImplementedException(),
        };
        return AddText(hfText, occurrence);
    }

    public void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages)
    {
        if (occurrence == XLHFOccurrence.AllPages)
        {
            ClearOccurrence(XLHFOccurrence.EvenPages);
            ClearOccurrence(XLHFOccurrence.FirstPage);
            ClearOccurrence(XLHFOccurrence.OddPages);
        }
        else
        {
            ClearOccurrence(occurrence);
        }
    }

    private void ClearOccurrence(XLHFOccurrence occurrence)
    {
        texts.Remove(occurrence);
        _images.Remove(occurrence);
    }
}
