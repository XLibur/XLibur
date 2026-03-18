using XLibur.Excel.ContentManagers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
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
        }
        else
        {
            if (!worksheet.Elements<Hyperlinks>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.Hyperlinks);
                worksheet.InsertAfter(new Hyperlinks(), previousElement);
            }

            var hyperlinks = worksheet.Elements<Hyperlinks>().First();
            cm.SetElement(XLWorksheetContents.Hyperlinks, hyperlinks);
            hyperlinks.RemoveAllChildren<Hyperlink>();
            foreach (var hl in xlWorksheet.Hyperlinks)
            {
                Hyperlink hyperlink;
                if (hl.IsExternal)
                {
                    var rId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    hyperlink = new Hyperlink { Reference = hl.Cell!.Address.ToString(), Id = rId };
                    worksheetPart.AddHyperlinkRelationship(hl.ExternalAddress!, true, rId);
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
                hyperlinks.AppendChild(hyperlink);
            }
        }
    }

    internal static void WritePrintOptions(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet)
    {
        if (!worksheet.Elements<PrintOptions>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.PrintOptions);
            worksheet.InsertAfter(new PrintOptions(), previousElement);
        }

        var printOptions = worksheet.Elements<PrintOptions>().First();
        cm.SetElement(XLWorksheetContents.PrintOptions, printOptions);

        printOptions.HorizontalCentered = xlWorksheet.PageSetup.CenterHorizontally;
        printOptions.VerticalCentered = xlWorksheet.PageSetup.CenterVertically;
        printOptions.Headings = xlWorksheet.PageSetup.ShowRowAndColumnHeadings;
        printOptions.GridLines = xlWorksheet.PageSetup.ShowGridlines;
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
        if (!worksheet.Elements<PageSetup>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.PageSetup);
            worksheet.InsertAfter(new PageSetup(), previousElement);
        }

        var pageSetup = worksheet.Elements<PageSetup>().First();
        cm.SetElement(XLWorksheetContents.PageSetup, pageSetup);

        SetPageSetupBasicProperties(pageSetup, xlWorksheet);
        SetPageSetupDpiAndScale(pageSetup, xlWorksheet);

        // For some reason some Excel files already contains pageSetup.Copies = 0
        // The validation fails for this
        // Let's remove the attribute of that's the case.
        if ((pageSetup.Copies ?? 0) <= 0)
            pageSetup.Copies = null;
    }

    private static void SetPageSetupBasicProperties(PageSetup pageSetup, XLWorksheet xlWorksheet)
    {
        pageSetup.Orientation = xlWorksheet.PageSetup.PageOrientation.ToOpenXml();
        pageSetup.PaperSize = (uint)xlWorksheet.PageSetup.PaperSize;
        pageSetup.BlackAndWhite = xlWorksheet.PageSetup.BlackAndWhite;
        pageSetup.Draft = xlWorksheet.PageSetup.DraftQuality;
        pageSetup.PageOrder = xlWorksheet.PageSetup.PageOrder.ToOpenXml();
        pageSetup.CellComments = xlWorksheet.PageSetup.ShowComments.ToOpenXml();
        pageSetup.Errors = xlWorksheet.PageSetup.PrintErrorValue.ToOpenXml();

        if (xlWorksheet.PageSetup.FirstPageNumber.HasValue)
        {
            // Negative first page numbers are written as uint, e.g. -1 is 4294967295.
            pageSetup.FirstPageNumber = UInt32Value.FromUInt32((uint)xlWorksheet.PageSetup.FirstPageNumber.Value);
            pageSetup.UseFirstPageNumber = true;
        }
        else
        {
            pageSetup.FirstPageNumber = null;
            pageSetup.UseFirstPageNumber = null;
        }
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
            pageSetup.Scale = (uint)xlWorksheet.PageSetup.Scale;
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
        if (headerFooter == null)
            headerFooter = new HeaderFooter();
        else
            worksheet.RemoveAllChildren<HeaderFooter>();

        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.HeaderFooter);
            worksheet.InsertAfter(headerFooter, previousElement);
            cm.SetElement(XLWorksheetContents.HeaderFooter, headerFooter);
        }
        if (((XLHeaderFooter)xlWorksheet.PageSetup.Header).Changed
            || ((XLHeaderFooter)xlWorksheet.PageSetup.Footer).Changed)
        {
            headerFooter.RemoveAllChildren();

            headerFooter.ScaleWithDoc = xlWorksheet.PageSetup.ScaleHFWithDocument;
            headerFooter.AlignWithMargins = xlWorksheet.PageSetup.AlignHFWithMargins;
            headerFooter.DifferentFirst = xlWorksheet.PageSetup.DifferentFirstPageOnHF;
            headerFooter.DifferentOddEven = xlWorksheet.PageSetup.DifferentOddEvenPagesOnHF;

            var oddHeader = new OddHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.OddPages));
            headerFooter.AppendChild(oddHeader);
            var oddFooter = new OddFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.OddPages));
            headerFooter.AppendChild(oddFooter);

            var evenHeader = new EvenHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.EvenPages));
            headerFooter.AppendChild(evenHeader);
            var evenFooter = new EvenFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.EvenPages));
            headerFooter.AppendChild(evenFooter);

            var firstHeader = new FirstHeader(xlWorksheet.PageSetup.Header.GetText(XLHFOccurrence.FirstPage));
            headerFooter.AppendChild(firstHeader);
            var firstFooter = new FirstFooter(xlWorksheet.PageSetup.Footer.GetText(XLHFOccurrence.FirstPage));
            headerFooter.AppendChild(firstFooter);
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
                             !xlWorksheet.PageSetup.RowBreaks.Contains((int)rb.Id!.Value))
                .ToList();

            foreach (var rb in rowBreaksToDelete)
            {
                rowBreaks.RemoveChild(rb);
            }

            var rowBreaksToAdd = xlWorksheet.PageSetup.RowBreaks
                .Where(xlRb => !existingBreaks.Any(rb => rb.Id?.HasValue == true && rb.Id.Value == xlRb));

            rowBreaks.Count = (uint)rowBreakCount;
            rowBreaks.ManualBreakCount = (uint)rowBreakCount;
            var lastRowNum = (uint)xlWorksheet.RangeAddress.LastAddress.RowNumber;
            foreach (var break1 in rowBreaksToAdd.Select(rb => new Break
            {
                Id = (uint)rb,
                Max = lastRowNum,
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
                             !xlWorksheet.PageSetup.ColumnBreaks.Contains((int)cb.Id!.Value))
                .ToList();

            foreach (var rb in columnBreaksToDelete)
            {
                columnBreaks.RemoveChild(rb);
            }

            var columnBreaksToAdd = xlWorksheet.PageSetup.ColumnBreaks
                .Where(xlCb => !existingBreaks.Any(cb => cb.Id?.HasValue == true && cb.Id.Value == xlCb));

            columnBreaks.Count = (uint)columnBreakCount;
            columnBreaks.ManualBreakCount = (uint)columnBreakCount;
            var maxColumnNumber = (uint)xlWorksheet.RangeAddress.LastAddress.ColumnNumber;
            foreach (var break1 in columnBreaksToAdd.Select(cb => new Break
            {
                Id = (uint)cb,
                Max = maxColumnNumber,
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
