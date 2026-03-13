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
        throw new NotImplementedException();
    }

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
    }
}
