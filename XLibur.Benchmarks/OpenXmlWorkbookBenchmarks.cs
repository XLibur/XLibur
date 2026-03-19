using System;
using System.Globalization;
using System.IO;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Benchmarks;

/// <summary>
/// Raw OpenXML SDK baseline — no XLibur/ClosedXML/EPPlus overhead.
/// Uses <see cref="OpenXmlWriter"/> streaming for sheet data and
/// manually builds the minimal stylesheet required for formatting.
/// This represents the practical floor for any library built on top of the SDK.
/// </summary>
[MemoryDiagnoser]
[Config(typeof(JoinSummaryConfig))]
public class OpenXmlWorkbookBenchmarks
{
    private const int RowCount = 50_000;
    private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private string[] _strings = null!;
    private double[] _numbers = null!;
    private DateTime[] _dates = null!;

    [GlobalSetup]
    public void Setup()
    {
        _strings = new string[RowCount];
        _numbers = new double[RowCount];
        _dates = new DateTime[RowCount];

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

        for (var i = 0; i < RowCount; i++)
        {
            _strings[i] = $"Item {i} - {random.Next(1000):D4}";
            _numbers[i] = Math.Round(random.NextDouble() * 10000, 2);
            _dates[i] = baseDate.AddDays(random.Next(0, 1500));
        }
    }

    [Benchmark]
    public void CreateAndSave()
    {
        using var stream = new MemoryStream();
        using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var relId = workbookPart.GetIdOfPart(worksheetPart);

        // Minimal stylesheet (default style only)
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateMinimalStylesheet();

        // Shared string table
        var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
        var sst = new SharedStringTable();

        // Stream sheet data
        using (var writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            // Header row
            WriteRow(writer, 1, [
                StringCell("Name", sst),
                StringCell("Amount", sst),
                StringCell("Date", sst)
            ]);

            // Data rows
            for (var i = 0; i < RowCount; i++)
            {
                var row = i + 2;
                WriteRow(writer, row, [
                    StringCell(_strings[i], sst),
                    NumberCell(_numbers[i]),
                    NumberCell(_dates[i].ToOADate())
                ]);
            }

            // Sum row
            var sumRow = RowCount + 2;
            WriteRow(writer, sumRow, [
                StringCell("Total", sst),
                FormulaCell($"SUM(B2:B{RowCount + 1})")
            ]);

            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
        }

        sstPart.SharedStringTable = sst;
        sst.Count = (uint)sst.ChildElements.Count;
        sst.UniqueCount = sst.Count;

        // Workbook
        workbookPart.Workbook = new Workbook(
            new Sheets(
                new Sheet { Name = "Data", SheetId = 1, Id = relId }));
    }

    [Benchmark]
    public void CreateFormattedAndSave()
    {
        using var stream = new MemoryStream();
        using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = doc.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var relId = workbookPart.GetIdOfPart(worksheetPart);

        // Stylesheet with formatting styles
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        var styles = CreateFormattedStylesheet();
        stylesPart.Stylesheet = styles;

        // Shared string table
        var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
        var sst = new SharedStringTable();

        // Stream sheet data
        using (var writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            // Header row
            WriteFormattedHeaderRow(writer, sst);

            // Data rows
            for (var i = 0; i < RowCount; i++)
            {
                var row = i + 2;
                var idx = i % _strings.Length;
                var applyFormatting = i % 2 == 0;

                WriteFormattedDataRow(writer, row, i, idx, applyFormatting, sst);
            }

            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
        }

        sstPart.SharedStringTable = sst;
        sst.Count = (uint)sst.ChildElements.Count;
        sst.UniqueCount = sst.Count;

        // Workbook
        workbookPart.Workbook = new Workbook(
            new Sheets(
                new Sheet { Name = "Formatted", SheetId = 1, Id = relId }));
    }

    #region Simple CreateAndSave helpers

    private static void WriteRow(OpenXmlWriter writer, int rowIndex, CellData[] cells)
    {
        writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
        var col = 'A';
        foreach (var cell in cells)
        {
            WriteCell(writer, $"{col}{rowIndex}", cell);
            col++;
        }
        writer.WriteEndElement(); // Row
    }

    private static void WriteCell(OpenXmlWriter writer, string reference, CellData data)
    {
        var cell = new Cell { CellReference = reference };
        if (data.StyleIndex > 0)
            cell.StyleIndex = data.StyleIndex;

        if (data.Formula is not null)
        {
            cell.DataType = CellValues.String;
            writer.WriteStartElement(cell);
            writer.WriteElement(new CellFormula(data.Formula));
            writer.WriteEndElement();
            return;
        }

        if (data.SharedStringId is not null)
        {
            cell.DataType = CellValues.SharedString;
            cell.CellValue = new CellValue(data.SharedStringId.Value);
        }
        else
        {
            cell.CellValue = new CellValue(data.NumericValue.ToString("G15", CultureInfo.InvariantCulture));
        }

        writer.WriteElement(cell);
    }

    private static CellData StringCell(string text, SharedStringTable sst)
    {
        var id = AddSharedString(sst, text);
        return new CellData { SharedStringId = id };
    }

    private static CellData NumberCell(double value) =>
        new() { NumericValue = value };

    private static CellData FormulaCell(string formula) =>
        new() { Formula = formula };

    private static int AddSharedString(SharedStringTable sst, string text)
    {
        var id = sst.ChildElements.Count;
        sst.AppendChild(new SharedStringItem(new Text(text)));
        return id;
    }

    #endregion

    #region Formatted CreateAndSave helpers

    private void WriteFormattedHeaderRow(OpenXmlWriter writer, SharedStringTable sst)
    {
        var headers = new[] { "Name", "Amount", "Date", "Quantity", "Price", "Total", "Status", "Category", "Region", "Notes" };
        writer.WriteStartElement(new Row { RowIndex = 1 });
        for (var c = 0; c < headers.Length; c++)
        {
            var colLetter = (char)('A' + c);
            var id = AddSharedString(sst, headers[c]);
            WriteCell(writer, $"{colLetter}1", new CellData { SharedStringId = id, StyleIndex = 1 }); // style 1 = header
        }
        writer.WriteEndElement();
    }

    private void WriteFormattedDataRow(OpenXmlWriter writer, int row, int i, int idx,
        bool applyFormatting, SharedStringTable sst)
    {
        writer.WriteStartElement(new Row { RowIndex = (uint)row });

        var status = (i % 3) switch { 0 => "Active", 1 => "Pending", _ => "Closed" };
        var region = (i % 5) switch { 0 => "North", 1 => "South", 2 => "East", 3 => "West", _ => "Central" };

        // For formatted rows, use style indices 2-11; for unformatted rows, use 0 (default).
        var s = applyFormatting ? 2u : 0u;

        WriteCellDirect(writer, row, 0, AddSharedString(sst, _strings[idx]), styleIndex: s);
        WriteCellNumberDirect(writer, row, 1, _numbers[idx], styleIndex: applyFormatting ? 3u : 0u);
        WriteCellNumberDirect(writer, row, 2, _dates[idx].ToOADate(), styleIndex: applyFormatting ? 4u : 0u);
        WriteCellNumberDirect(writer, row, 3, (i % 500) + 1, styleIndex: applyFormatting ? 5u : 0u);
        WriteCellNumberDirect(writer, row, 4, _numbers[idx] * 0.1, styleIndex: applyFormatting ? 6u : 0u);
        WriteCellNumberDirect(writer, row, 5, _numbers[idx] * ((i % 500) + 1) * 0.1, styleIndex: applyFormatting ? 7u : 0u);
        WriteCellDirect(writer, row, 6, AddSharedString(sst, status), styleIndex: applyFormatting ? 8u : 0u);
        WriteCellDirect(writer, row, 7, AddSharedString(sst, $"Cat-{(i % 12) + 1}"), styleIndex: applyFormatting ? 9u : 0u);
        WriteCellDirect(writer, row, 8, AddSharedString(sst, region), styleIndex: applyFormatting ? 10u : 0u);
        WriteCellDirect(writer, row, 9, AddSharedString(sst, $"Note for row {row}"), styleIndex: applyFormatting ? 11u : 0u);

        writer.WriteEndElement(); // Row
    }

    private static void WriteCellDirect(OpenXmlWriter writer, int row, int col, int sstId, uint styleIndex)
    {
        var colLetter = (char)('A' + col);
        var cell = new Cell
        {
            CellReference = $"{colLetter}{row}",
            DataType = CellValues.SharedString,
            CellValue = new CellValue(sstId),
            StyleIndex = styleIndex
        };
        writer.WriteElement(cell);
    }

    private static void WriteCellNumberDirect(OpenXmlWriter writer, int row, int col, double value, uint styleIndex)
    {
        var colLetter = (char)('A' + col);
        var cell = new Cell
        {
            CellReference = $"{colLetter}{row}",
            CellValue = new CellValue(value.ToString("G15", CultureInfo.InvariantCulture)),
            StyleIndex = styleIndex
        };
        writer.WriteElement(cell);
    }

    #endregion

    #region Stylesheet builders

    private static Stylesheet CreateMinimalStylesheet()
    {
        return new Stylesheet(
            new Fonts(new Font()),
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
                      new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
            new Borders(new Border()),
            new CellFormats(new CellFormat()));
    }

    /// <summary>
    /// Builds a stylesheet with 12 cell formats (0 = default, 1 = header, 2-11 = column styles).
    /// This mirrors the formatting applied by the XLibur/EPPlus benchmarks.
    /// </summary>
    private static Stylesheet CreateFormattedStylesheet()
    {
        var fonts = new Fonts(
            new Font(), // 0: default
            new Font(new Bold(), new Color { Rgb = "FFFFFFFF" }), // 1: header (white bold)
            new Font(new Bold()), // 2: bold
            new Font(new Color { Rgb = "FF8B0000" }), // 3: dark red
            new Font(new Italic()), // 4: italic
            new Font(), // 5: default (quantity)
            new Font(new FontName { Val = "Consolas" }, new FontSize { Val = 10 }), // 6: Consolas 10
            new Font(new Bold()), // 7: bold (total)
            new Font(), // 8: status (varies at runtime, use default here)
            new Font(new Underline(), new FontName { Val = "Calibri" }), // 9: underline
            new Font(new FontName { Val = "Georgia" }, new FontSize { Val = 12 }, new Color { Rgb = "FF008080" }), // 10: Georgia teal
            new Font(new Color { Rgb = "FF2F4F4F" }) // 11: dark slate gray
        );

        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }), // 0: required
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), // 1: required
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FF00008B" } }), // 2: dark blue (header)
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFADD8E6" } }), // 3: light blue
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FF90EE90" } }), // 4: light green
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFFFFFA0" } }), // 5: light yellow
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFF08080" } }), // 6: light coral
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFE6E6FA" } }), // 7: lavender
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFD3D3D3" } }), // 8: light gray
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFF5DEB3" } }), // 9: wheat
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFFFA07A" } })  // 10: light salmon
        );

        var borders = new Borders(
            new Border(), // 0: none
            new Border(new BottomBorder { Style = BorderStyleValues.Double, Color = new Color { Rgb = "FF000000" } }), // 1: header bottom
            new Border( // 2: thin gray all around
                new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Rgb = "FF808080" } },
                new RightBorder { Style = BorderStyleValues.Thin, Color = new Color { Rgb = "FF808080" } },
                new TopBorder { Style = BorderStyleValues.Thin, Color = new Color { Rgb = "FF808080" } },
                new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Rgb = "FF808080" } }),
            new Border( // 3: medium dark goldenrod
                new LeftBorder { Style = BorderStyleValues.Medium, Color = new Color { Rgb = "FFB8860B" } },
                new RightBorder { Style = BorderStyleValues.Medium, Color = new Color { Rgb = "FFB8860B" } },
                new TopBorder { Style = BorderStyleValues.Medium, Color = new Color { Rgb = "FFB8860B" } },
                new BottomBorder { Style = BorderStyleValues.Medium, Color = new Color { Rgb = "FFB8860B" } }),
            new Border(new BottomBorder { Style = BorderStyleValues.Dashed, Color = new Color { Rgb = "FF000080" } }), // 4: dashed navy bottom
            new Border(new RightBorder { Style = BorderStyleValues.Dotted, Color = new Color { Rgb = "FF800080" } }), // 5: dotted purple right
            new Border( // 6: thick teal
                new LeftBorder { Style = BorderStyleValues.Thick, Color = new Color { Rgb = "FF008080" } },
                new RightBorder { Style = BorderStyleValues.Thick, Color = new Color { Rgb = "FF008080" } },
                new TopBorder { Style = BorderStyleValues.Thick, Color = new Color { Rgb = "FF008080" } },
                new BottomBorder { Style = BorderStyleValues.Thick, Color = new Color { Rgb = "FF008080" } }),
            new Border( // 7: double chocolate
                new LeftBorder { Style = BorderStyleValues.Double, Color = new Color { Rgb = "FFD2691E" } },
                new RightBorder { Style = BorderStyleValues.Double, Color = new Color { Rgb = "FFD2691E" } },
                new TopBorder { Style = BorderStyleValues.Double, Color = new Color { Rgb = "FFD2691E" } },
                new BottomBorder { Style = BorderStyleValues.Double, Color = new Color { Rgb = "FFD2691E" } })
        );

        var numberFormats = new NumberingFormats(
            new NumberingFormat { NumberFormatId = 164, FormatCode = "#,##0.00" },
            new NumberingFormat { NumberFormatId = 165, FormatCode = "$ #,##0.00" },
            new NumberingFormat { NumberFormatId = 166, FormatCode = "_($* #,##0.00_)" }
        );

        var cellFormats = new CellFormats(
            new CellFormat(), // 0: default
            new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true,
                             Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, // 1: header
            new CellFormat { FontId = 2, FillId = 3, ApplyFont = true, ApplyFill = true }, // 2: name (bold, light blue)
            new CellFormat { FontId = 3, BorderId = 2, NumberFormatId = 164, ApplyFont = true, ApplyBorder = true, ApplyNumberFormat = true }, // 3: amount
            new CellFormat { FontId = 4, FillId = 4, NumberFormatId = 15, ApplyFont = true, ApplyFill = true, ApplyNumberFormat = true }, // 4: date
            new CellFormat { FillId = 5, BorderId = 3, ApplyFill = true, ApplyBorder = true,
                             Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, // 5: quantity
            new CellFormat { FontId = 6, FillId = 6, NumberFormatId = 165, ApplyFont = true, ApplyFill = true, ApplyNumberFormat = true }, // 6: price
            new CellFormat { FontId = 7, BorderId = 4, NumberFormatId = 166, ApplyFont = true, ApplyBorder = true, ApplyNumberFormat = true }, // 7: total
            new CellFormat { FillId = 7, ApplyFill = true }, // 8: status
            new CellFormat { FontId = 9, FillId = 8, BorderId = 5, ApplyFont = true, ApplyFill = true, ApplyBorder = true,
                             Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right } }, // 9: category
            new CellFormat { FontId = 10, FillId = 9, BorderId = 6, ApplyFont = true, ApplyFill = true, ApplyBorder = true }, // 10: region
            new CellFormat { FontId = 11, FillId = 10, BorderId = 7, ApplyFont = true, ApplyFill = true, ApplyBorder = true,
                             Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top } } // 11: notes
        );

        return new Stylesheet(numberFormats, fonts, fills, borders, cellFormats);
    }

    #endregion

#nullable enable
    private struct CellData
    {
        public int? SharedStringId;
        public double NumericValue;
        public string? Formula;
        public uint StyleIndex;
    }
#nullable restore
}
