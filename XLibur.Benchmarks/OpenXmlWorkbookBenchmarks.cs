using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Benchmarks;

/// <summary>
/// Raw OpenXML SDK baseline — no XLibur/ClosedXML/EPPlus overhead.
/// Uses <see cref="OpenXmlWriter"/> for structural elements and writes
/// cell data directly via the underlying <see cref="XmlWriter"/> to
/// avoid per-cell DOM allocations. Shared strings are deduplicated
/// with a dictionary and written in bulk at the end.
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

        // Minimal stylesheet
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateMinimalStylesheet();

        // Dictionary-based SST for O(1) dedup + lookup
        var sstEntries = new List<string>();
        var sstMap = new Dictionary<string, int>();

        // Stream sheet data via XmlWriter directly
        using (var writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            // Header row
            WriteRowStart(writer, 1);
            WriteStringCell(writer, "A1", GetSstId("Name", sstEntries, sstMap), 0);
            WriteStringCell(writer, "B1", GetSstId("Amount", sstEntries, sstMap), 0);
            WriteStringCell(writer, "C1", GetSstId("Date", sstEntries, sstMap), 0);
            WriteRowEnd(writer);

            // Data rows
            Span<char> cellRef = stackalloc char[10];
            for (var i = 0; i < RowCount; i++)
            {
                var row = i + 2;
                WriteRowStart(writer, row);

                FormatCellRef(cellRef, 'A', row, out var len);
                WriteStringCellRaw(writer, cellRef, len, GetSstId(_strings[i], sstEntries, sstMap), 0);

                FormatCellRef(cellRef, 'B', row, out len);
                WriteNumberCellRaw(writer, cellRef, len, _numbers[i], 0);

                FormatCellRef(cellRef, 'C', row, out len);
                WriteNumberCellRaw(writer, cellRef, len, _dates[i].ToOADate(), 0);

                WriteRowEnd(writer);
            }

            // Sum row
            var sumRow = RowCount + 2;
            WriteRowStart(writer, sumRow);
            WriteStringCell(writer, $"A{sumRow}", GetSstId("Total", sstEntries, sstMap), 0);
            WriteFormulaCell(writer, $"B{sumRow}", $"SUM(B2:B{RowCount + 1})");
            WriteRowEnd(writer);

            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
        }

        // Write SST part
        WriteSstPart(workbookPart, sstEntries);

        // Workbook
        workbookPart.Workbook = new Workbook(
            new Sheets(new Sheet { Name = "Data", SheetId = 1, Id = relId }));
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
        stylesPart.Stylesheet = CreateFormattedStylesheet();

        // Dictionary-based SST
        var sstEntries = new List<string>(RowCount * 5);
        var sstMap = new Dictionary<string, int>(RowCount * 2);

        // Stream sheet data
        using (var writer = OpenXmlWriter.Create(worksheetPart))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            // Header row
            WriteFormattedHeaderRow(writer, sstEntries, sstMap);

            // Data rows
            Span<char> cellRef = stackalloc char[10];
            for (var i = 0; i < RowCount; i++)
            {
                var row = i + 2;
                var idx = i % _strings.Length;
                var applyFormatting = i % 2 == 0;

                WriteFormattedDataRow(writer, cellRef, row, i, idx, applyFormatting, sstEntries, sstMap);
            }

            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
        }

        // Write SST part
        WriteSstPart(workbookPart, sstEntries);

        // Workbook
        workbookPart.Workbook = new Workbook(
            new Sheets(new Sheet { Name = "Formatted", SheetId = 1, Id = relId }));
    }

    #region Cell writing helpers (XmlWriter-based, zero DOM allocation)

    private static void WriteRowStart(OpenXmlWriter writer, int rowIndex)
    {
        writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
    }

    private static void WriteRowEnd(OpenXmlWriter writer)
    {
        writer.WriteEndElement();
    }

    private static void WriteStringCell(OpenXmlWriter writer, string cellRef, int sstId, uint styleIndex)
    {
        var attrs = new List<OpenXmlAttribute>
        {
            new("r", null, cellRef),
            new("t", null, "s")
        };
        if (styleIndex > 0)
            attrs.Add(new OpenXmlAttribute("s", null, styleIndex.ToString()));

        writer.WriteStartElement(new Cell(), attrs);
        writer.WriteElement(new CellValue(sstId));
        writer.WriteEndElement();
    }

    private static void WriteStringCellRaw(OpenXmlWriter writer, Span<char> cellRef, int cellRefLen, int sstId, uint styleIndex)
    {
        var refStr = cellRef[..cellRefLen].ToString();
        var attrs = new List<OpenXmlAttribute>
        {
            new("r", null, refStr),
            new("t", null, "s")
        };
        if (styleIndex > 0)
            attrs.Add(new OpenXmlAttribute("s", null, styleIndex.ToString()));

        writer.WriteStartElement(new Cell(), attrs);
        writer.WriteElement(new CellValue(sstId));
        writer.WriteEndElement();
    }

    private static void WriteNumberCellRaw(OpenXmlWriter writer, Span<char> cellRef, int cellRefLen, double value, uint styleIndex)
    {
        var refStr = cellRef[..cellRefLen].ToString();
        var attrs = new List<OpenXmlAttribute>
        {
            new("r", null, refStr),
        };
        if (styleIndex > 0)
            attrs.Add(new OpenXmlAttribute("s", null, styleIndex.ToString()));

        writer.WriteStartElement(new Cell(), attrs);
        writer.WriteElement(new CellValue(value.ToString("G15", CultureInfo.InvariantCulture)));
        writer.WriteEndElement();
    }

    private static void WriteFormulaCell(OpenXmlWriter writer, string cellRef, string formula)
    {
        var attrs = new List<OpenXmlAttribute> { new("r", null, cellRef) };
        writer.WriteStartElement(new Cell(), attrs);
        writer.WriteElement(new CellFormula(formula));
        writer.WriteEndElement();
    }

    private static void FormatCellRef(Span<char> buffer, char col, int row, out int length)
    {
        buffer[0] = col;
        row.TryFormat(buffer[1..], out var written);
        length = 1 + written;
    }

    #endregion

    #region Shared string table

    private static int GetSstId(string text, List<string> entries, Dictionary<string, int> map)
    {
        if (map.TryGetValue(text, out var id))
            return id;

        id = entries.Count;
        entries.Add(text);
        map.Add(text, id);
        return id;
    }

    private static void WriteSstPart(WorkbookPart workbookPart, List<string> entries)
    {
        var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
        using var xmlWriter = XmlWriter.Create(sstPart.GetStream(FileMode.Create),
            new XmlWriterSettings { Encoding = System.Text.Encoding.UTF8 });

        xmlWriter.WriteStartDocument(true);
        xmlWriter.WriteStartElement("sst", Ns);
        xmlWriter.WriteAttributeString("count", entries.Count.ToString());
        xmlWriter.WriteAttributeString("uniqueCount", entries.Count.ToString());

        foreach (var text in entries)
        {
            xmlWriter.WriteStartElement("si", Ns);
            xmlWriter.WriteStartElement("t", Ns);
            if (text.Length > 0 && (text[0] == ' ' || text[^1] == ' '))
                xmlWriter.WriteAttributeString("xml", "space", null, "preserve");
            xmlWriter.WriteString(text);
            xmlWriter.WriteEndElement(); // t
            xmlWriter.WriteEndElement(); // si
        }

        xmlWriter.WriteEndElement(); // sst
    }

    #endregion

    #region Formatted benchmark helpers

    private void WriteFormattedHeaderRow(OpenXmlWriter writer, List<string> sstEntries, Dictionary<string, int> sstMap)
    {
        var headers = new[] { "Name", "Amount", "Date", "Quantity", "Price", "Total", "Status", "Category", "Region", "Notes" };
        WriteRowStart(writer, 1);
        for (var c = 0; c < headers.Length; c++)
        {
            var colLetter = (char)('A' + c);
            WriteStringCell(writer, $"{colLetter}1", GetSstId(headers[c], sstEntries, sstMap), 1);
        }
        WriteRowEnd(writer);
    }

    private void WriteFormattedDataRow(OpenXmlWriter writer, Span<char> cellRef, int row, int i, int idx,
        bool applyFormatting, List<string> sstEntries, Dictionary<string, int> sstMap)
    {
        WriteRowStart(writer, row);

        var status = (i % 3) switch { 0 => "Active", 1 => "Pending", _ => "Closed" };
        var region = (i % 5) switch { 0 => "North", 1 => "South", 2 => "East", 3 => "West", _ => "Central" };

        // cols A-J, style indices 2-11 for formatted, 0 for unformatted
        FormatCellRef(cellRef, 'A', row, out var len);
        WriteStringCellRaw(writer, cellRef, len, GetSstId(_strings[idx], sstEntries, sstMap), applyFormatting ? 2u : 0u);

        FormatCellRef(cellRef, 'B', row, out len);
        WriteNumberCellRaw(writer, cellRef, len, _numbers[idx], applyFormatting ? 3u : 0u);

        FormatCellRef(cellRef, 'C', row, out len);
        WriteNumberCellRaw(writer, cellRef, len, _dates[idx].ToOADate(), applyFormatting ? 4u : 0u);

        FormatCellRef(cellRef, 'D', row, out len);
        WriteNumberCellRaw(writer, cellRef, len, (i % 500) + 1, applyFormatting ? 5u : 0u);

        FormatCellRef(cellRef, 'E', row, out len);
        WriteNumberCellRaw(writer, cellRef, len, _numbers[idx] * 0.1, applyFormatting ? 6u : 0u);

        FormatCellRef(cellRef, 'F', row, out len);
        WriteNumberCellRaw(writer, cellRef, len, _numbers[idx] * ((i % 500) + 1) * 0.1, applyFormatting ? 7u : 0u);

        FormatCellRef(cellRef, 'G', row, out len);
        WriteStringCellRaw(writer, cellRef, len, GetSstId(status, sstEntries, sstMap), applyFormatting ? 8u : 0u);

        FormatCellRef(cellRef, 'H', row, out len);
        WriteStringCellRaw(writer, cellRef, len, GetSstId($"Cat-{(i % 12) + 1}", sstEntries, sstMap), applyFormatting ? 9u : 0u);

        FormatCellRef(cellRef, 'I', row, out len);
        WriteStringCellRaw(writer, cellRef, len, GetSstId(region, sstEntries, sstMap), applyFormatting ? 10u : 0u);

        FormatCellRef(cellRef, 'J', row, out len);
        WriteStringCellRaw(writer, cellRef, len, GetSstId($"Note for row {row}", sstEntries, sstMap), applyFormatting ? 11u : 0u);

        WriteRowEnd(writer);
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
            new CellFormat
            {
                FontId = 1,
                FillId = 2,
                BorderId = 1,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center }
            }, // 1: header
            new CellFormat { FontId = 2, FillId = 3, ApplyFont = true, ApplyFill = true }, // 2: name (bold, light blue)
            new CellFormat { FontId = 3, BorderId = 2, NumberFormatId = 164, ApplyFont = true, ApplyBorder = true, ApplyNumberFormat = true }, // 3: amount
            new CellFormat { FontId = 4, FillId = 4, NumberFormatId = 15, ApplyFont = true, ApplyFill = true, ApplyNumberFormat = true }, // 4: date
            new CellFormat
            {
                FillId = 5,
                BorderId = 3,
                ApplyFill = true,
                ApplyBorder = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center }
            }, // 5: quantity
            new CellFormat { FontId = 6, FillId = 6, NumberFormatId = 165, ApplyFont = true, ApplyFill = true, ApplyNumberFormat = true }, // 6: price
            new CellFormat { FontId = 7, BorderId = 4, NumberFormatId = 166, ApplyFont = true, ApplyBorder = true, ApplyNumberFormat = true }, // 7: total
            new CellFormat { FillId = 7, ApplyFill = true }, // 8: status
            new CellFormat
            {
                FontId = 9,
                FillId = 8,
                BorderId = 5,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right }
            }, // 9: category
            new CellFormat { FontId = 10, FillId = 9, BorderId = 6, ApplyFont = true, ApplyFill = true, ApplyBorder = true }, // 10: region
            new CellFormat
            {
                FontId = 11,
                FillId = 10,
                BorderId = 7,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                Alignment = new Alignment { WrapText = true, Vertical = VerticalAlignmentValues.Top }
            } // 11: notes
        );

        return new Stylesheet(numberFormats, fonts, fills, borders, cellFormats);
    }

    #endregion
}
