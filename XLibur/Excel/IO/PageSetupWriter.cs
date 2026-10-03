using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.XLWorkbook;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using Break = DocumentFormat.OpenXml.Spreadsheet.Break;

namespace XLibur.Excel.IO;

internal static class PageSetupWriter
{
    internal static void WriteHyperlinks(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        var relToRemove = worksheetPart.HyperlinkRelationships.ToList();
        relToRemove.ForEach(worksheetPart.DeleteReferenceRelationship);
        if (!xlWorksheet.Hyperlinks.Any())
        {
            worksheet.RemoveAllChildren<Hyperlinks>();
            cm.SetElement(XLWorksheetContents.Hyperlinks, null);
            return;
        }

        if (!worksheet.Elements<Hyperlinks>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.Hyperlinks);
            worksheet.InsertAfter(new Hyperlinks(), previousElement);
        }

        var hyperlinks = worksheet.Elements<Hyperlinks>().First();
        cm.SetElement(XLWorksheetContents.Hyperlinks, hyperlinks);
        hyperlinks.RemoveAllChildren<Hyperlink>();

        // IDs of the relationships the part still holds. A kept hyperlink ID is only reused when it
        // is free here: another hyperlink may already have claimed it (two cells loaded from one
        // range), or the hyperlink may have moved from another sheet whose part numbered its own.
        var usedRelIds = new HashSet<string>(StringComparer.Ordinal);
        usedRelIds.UnionWith(worksheetPart.Parts.Select(p => p.RelationshipId));
        usedRelIds.UnionWith(worksheetPart.ExternalRelationships.Select(r => r.Id));
        usedRelIds.UnionWith(worksheetPart.DataPartReferenceRelationships.Select(r => r.Id));
        foreach (var hl in xlWorksheet.Hyperlinks)
            hyperlinks.AppendChild(CreateHyperlink(hl, worksheetPart, context, usedRelIds));
    }

    private static Hyperlink CreateHyperlink(XLHyperlink hl, WorksheetPart worksheetPart, SaveContext context,
        HashSet<string> usedRelIds)
    {
        Hyperlink hyperlink;
        if (hl.IsExternal)
        {
            var rId = hl.RelId;
            if (rId is not null && usedRelIds.Add(rId))
                context.RelIdGenerator.Reserve(RelType.Workbook, rId);
            else
            {
                rId = context.RelIdGenerator.GetNext(RelType.Workbook);
                usedRelIds.Add(rId);
            }

            hyperlink = new Hyperlink { Reference = hl.Cell!.Address.ToString(), Id = rId };
            worksheetPart.AddHyperlinkRelationship(hl.ExternalAddress!, true, rId);
            hl.RelId = rId;
        }
        else
        {
            hyperlink = new Hyperlink
            {
                Reference = hl.Cell!.Address.ToString(),
                Location = hl.InternalAddress,
                Display = hl.Cell.GetFormattedString()
            };
        }

        if (!string.IsNullOrWhiteSpace(hl.Tooltip))
            hyperlink.Tooltip = hl.Tooltip;
        return hyperlink;
    }

    internal static void WritePrintOptions(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        var printOptions = worksheet.Elements<PrintOptions>().FirstOrDefault();
        var loadedEmpty = printOptions is not null && SchemaDefault.IsEmpty(printOptions);
        printOptions ??= new PrintOptions();

        var pageSetup = xlWorksheet.PageSetup;
        printOptions.HorizontalCentered = SchemaDefault.Bool(printOptions.HorizontalCentered, pageSetup.CenterHorizontally, false);
        printOptions.VerticalCentered = SchemaDefault.Bool(printOptions.VerticalCentered, pageSetup.CenterVertically, false);
        printOptions.Headings = SchemaDefault.Bool(printOptions.Headings, pageSetup.ShowRowAndColumnHeadings, false);
        printOptions.GridLines = SchemaDefault.Bool(printOptions.GridLines, pageSetup.ShowGridlines, false);

        SchemaDefault.Place(worksheet, cm, XLWorksheetContents.PrintOptions, printOptions, loadedEmpty);
    }

    internal static void WritePageMargins(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        if (!worksheet.Elements<PageMargins>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.PageMargins);
            worksheet.InsertAfter(new PageMargins(), previousElement);
        }

        var pageMargins = worksheet.Elements<PageMargins>().First();
        cm.SetElement(XLWorksheetContents.PageMargins, pageMargins);
        pageMargins.Left = xlWorksheet.PageSetup.Margins.Left;
        pageMargins.Right = xlWorksheet.PageSetup.Margins.Right;
        pageMargins.Top = xlWorksheet.PageSetup.Margins.Top;
        pageMargins.Bottom = xlWorksheet.PageSetup.Margins.Bottom;
        pageMargins.Header = xlWorksheet.PageSetup.Margins.Header;
        pageMargins.Footer = xlWorksheet.PageSetup.Margins.Footer;
    }

    internal static void WritePageSetup(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        var pageSetup = worksheet.Elements<PageSetup>().FirstOrDefault();
        var loadedEmpty = pageSetup is not null && SchemaDefault.IsEmpty(pageSetup);
        pageSetup ??= new PageSetup();

        SetPageSetupBasicProperties(pageSetup, xlWorksheet);
        SetPageSetupDpiAndScale(pageSetup, xlWorksheet);

        // For some reason some Excel files already contains pageSetup.Copies = 0
        // The validation fails for this
        // Let's remove the attribute of that's the case.
        if ((pageSetup.Copies ?? 0) <= 0)
            pageSetup.Copies = null;

        SchemaDefault.Place(worksheet, cm, XLWorksheetContents.PageSetup, pageSetup, loadedEmpty);
    }

    private static void SetPageSetupBasicProperties(PageSetup pageSetup, XLWorksheet xlWorksheet)
    {
        var model = xlWorksheet.PageSetup;
        pageSetup.Orientation = SchemaDefault.Enum(pageSetup.Orientation, model.PageOrientation.ToOpenXml(), OrientationValues.Default);
        pageSetup.PaperSize = SchemaDefault.UInt(pageSetup.PaperSize, (uint)model.PaperSize, 1);
        pageSetup.BlackAndWhite = SchemaDefault.Bool(pageSetup.BlackAndWhite, model.BlackAndWhite, false);
        pageSetup.Draft = SchemaDefault.Bool(pageSetup.Draft, model.DraftQuality, false);
        pageSetup.PageOrder = SchemaDefault.Enum(pageSetup.PageOrder, model.PageOrder.ToOpenXml(), PageOrderValues.DownThenOver);
        pageSetup.CellComments = SchemaDefault.Enum(pageSetup.CellComments, model.ShowComments.ToOpenXml(), CellCommentsValues.None);
        pageSetup.Errors = SchemaDefault.Enum(pageSetup.Errors, model.PrintErrorValue.ToOpenXml(), PrintErrorValues.Displayed);

        if (model.FirstPageNumber.HasValue)
        {
            // Negative first page numbers are written as uint, e.g. -1 is 4294967295.
            pageSetup.FirstPageNumber = UInt32Value.FromUInt32((uint)model.FirstPageNumber.Value);
            pageSetup.UseFirstPageNumber = true;
        }
        else if (pageSetup.UseFirstPageNumber is { HasValue: true, Value: true })
        {
            pageSetup.FirstPageNumber = null;
            pageSetup.UseFirstPageNumber = null;
        }

        // Otherwise the number is not used, and whatever the file had stays as it was.
    }

    private static void SetPageSetupDpiAndScale(PageSetup pageSetup, XLWorksheet xlWorksheet)
    {
        pageSetup.HorizontalDpi = xlWorksheet.PageSetup.HorizontalDpi > 0
            ? (uint)xlWorksheet.PageSetup.HorizontalDpi
            : null;

        pageSetup.VerticalDpi = xlWorksheet.PageSetup.VerticalDpi > 0
            ? (uint)xlWorksheet.PageSetup.VerticalDpi
            : null;

        if (xlWorksheet.PageSetup.Scale > 0)
        {
            pageSetup.Scale = SchemaDefault.UInt(pageSetup.Scale, (uint)xlWorksheet.PageSetup.Scale, 100);
            pageSetup.FitToWidth = null;
            pageSetup.FitToHeight = null;
        }
        else
        {
            pageSetup.Scale = null;

            if (xlWorksheet.PageSetup.PagesWide >= 0 && xlWorksheet.PageSetup.PagesWide != 1)
                pageSetup.FitToWidth = (uint)xlWorksheet.PageSetup.PagesWide;

            if (xlWorksheet.PageSetup.PagesTall >= 0 && xlWorksheet.PageSetup.PagesTall != 1)
                pageSetup.FitToHeight = (uint)xlWorksheet.PageSetup.PagesTall;
        }
    }

    internal static void WriteHeaderFooter(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        var headerFooter = worksheet.Elements<HeaderFooter>().FirstOrDefault();
        var loadedEmpty = headerFooter is not null && SchemaDefault.IsEmpty(headerFooter);
        if (headerFooter == null)
            headerFooter = new HeaderFooter();
        else
            worksheet.RemoveAllChildren<HeaderFooter>();

        var pageSetup = xlWorksheet.PageSetup;
        if (((XLHeaderFooter)pageSetup.Header).Changed || ((XLHeaderFooter)pageSetup.Footer).Changed)
        {
            headerFooter.RemoveAllChildren();

            headerFooter.ScaleWithDoc = SchemaDefault.Bool(headerFooter.ScaleWithDoc, pageSetup.ScaleHFWithDocument, true);
            headerFooter.AlignWithMargins = SchemaDefault.Bool(headerFooter.AlignWithMargins, pageSetup.AlignHFWithMargins, true);
            headerFooter.DifferentFirst = SchemaDefault.Bool(headerFooter.DifferentFirst, pageSetup.DifferentFirstPageOnHF, false);
            headerFooter.DifferentOddEven = SchemaDefault.Bool(headerFooter.DifferentOddEven, pageSetup.DifferentOddEvenPagesOnHF, false);

            // Excel writes a header or footer only when it has text.
            AppendText(headerFooter, pageSetup.Header.GetText(XLHFOccurrence.OddPages), t => new OddHeader(t));
            AppendText(headerFooter, pageSetup.Footer.GetText(XLHFOccurrence.OddPages), t => new OddFooter(t));
            AppendText(headerFooter, pageSetup.Header.GetText(XLHFOccurrence.EvenPages), t => new EvenHeader(t));
            AppendText(headerFooter, pageSetup.Footer.GetText(XLHFOccurrence.EvenPages), t => new EvenFooter(t));
            AppendText(headerFooter, pageSetup.Header.GetText(XLHFOccurrence.FirstPage), t => new FirstHeader(t));
            AppendText(headerFooter, pageSetup.Footer.GetText(XLHFOccurrence.FirstPage), t => new FirstFooter(t));
        }

        SchemaDefault.Place(worksheet, cm, XLWorksheetContents.HeaderFooter, headerFooter, loadedEmpty);
        return;

        static void AppendText(HeaderFooter headerFooter, string text, Func<string, OpenXmlElement> create)
        {
            if (!string.IsNullOrEmpty(text))
                headerFooter.AppendChild(create(text));
        }
    }

    internal static void WriteRowBreaks(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        var rowBreakCount = xlWorksheet.PageSetup.RowBreaks.Count;
        if (rowBreakCount > 0)
        {
            if (!worksheet.Elements<RowBreaks>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.RowBreaks);
                worksheet.InsertAfter(new RowBreaks(), previousElement);
            }

            var rowBreaks = worksheet.Elements<RowBreaks>().First();

            var existingBreaks = rowBreaks.ChildElements.OfType<Break>().ToArray();
            var rowBreaksToDelete = existingBreaks
                .Where(rb => rb.Id?.Value is null ||
                             !xlWorksheet.PageSetup.RowBreaks.Contains((int)rb.Id.Value))
                .ToList();

            foreach (var rb in rowBreaksToDelete)
            {
                rowBreaks.RemoveChild(rb);
            }

            var rowBreaksToAdd = xlWorksheet.PageSetup.RowBreaks
                .Where(xlRb => !existingBreaks.Any(rb => rb.Id?.HasValue == true && rb.Id.Value == xlRb));

            rowBreaks.Count = (uint)rowBreakCount;
            rowBreaks.ManualBreakCount = (uint)rowBreakCount;
            // brk@max on a row (horizontal) break is a COLUMN extent — how far the
            // break spans across columns — not a row index. Excel writes the full
            // sheet width (0-based XFD = 16383). Writing a row count here (e.g. the
            // 1048576-row sheet extent) makes Excel render a bogus scrollbar.
            // See ClosedXML issue #2842.
            const uint lastColumn = XLHelper.MaxColumnNumber - 1;
            foreach (var break1 in rowBreaksToAdd.Select(rb => new Break
            {
                Id = (uint)rb,
                Max = lastColumn,
                ManualPageBreak = true
            }))
                rowBreaks.AppendChild(break1);
            cm.SetElement(XLWorksheetContents.RowBreaks, rowBreaks);
        }
        else
        {
            worksheet.RemoveAllChildren<RowBreaks>();
            cm.SetElement(XLWorksheetContents.RowBreaks, null);
        }
    }

    internal static void WriteColumnBreaks(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        var columnBreakCount = xlWorksheet.PageSetup.ColumnBreaks.Count;
        if (columnBreakCount > 0)
        {
            if (!worksheet.Elements<ColumnBreaks>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.ColumnBreaks);
                worksheet.InsertAfter(new ColumnBreaks(), previousElement);
            }

            var columnBreaks = worksheet.Elements<ColumnBreaks>().First();

            var existingBreaks = columnBreaks.ChildElements.OfType<Break>().ToArray();
            var columnBreaksToDelete = existingBreaks
                .Where(cb => cb.Id?.Value is null ||
                             !xlWorksheet.PageSetup.ColumnBreaks.Contains((int)cb.Id.Value))
                .ToList();

            foreach (var rb in columnBreaksToDelete)
            {
                columnBreaks.RemoveChild(rb);
            }

            var columnBreaksToAdd = xlWorksheet.PageSetup.ColumnBreaks
                .Where(xlCb => !existingBreaks.Any(cb => cb.Id?.HasValue == true && cb.Id.Value == xlCb));

            columnBreaks.Count = (uint)columnBreakCount;
            columnBreaks.ManualBreakCount = (uint)columnBreakCount;
            // brk@max on a column (vertical) break is a ROW extent — how far the
            // break spans down rows — not a column index. Excel writes the full
            // sheet height (0-based row 1048575). See ClosedXML issue #2842.
            const uint lastRow = XLHelper.MaxRowNumber - 1;
            foreach (var break1 in columnBreaksToAdd.Select(cb => new Break
            {
                Id = (uint)cb,
                Max = lastRow,
                ManualPageBreak = true
            }))
                columnBreaks.AppendChild(break1);
            cm.SetElement(XLWorksheetContents.ColumnBreaks, columnBreaks);
        }
        else
        {
            worksheet.RemoveAllChildren<ColumnBreaks>();
            cm.SetElement(XLWorksheetContents.ColumnBreaks, null);
        }
    }
}
