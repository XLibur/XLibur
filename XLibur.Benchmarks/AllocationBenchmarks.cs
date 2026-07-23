using System.Drawing;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Excel.Cells;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Micro-benchmarks isolating the methods refactored on the perf/reduce-allocations
/// branch. Run the same benchmark against the parent commit for a before/after
/// allocation comparison (the method signatures are identical across both).
/// </summary>
[MemoryDiagnoser]
public class AllocationBenchmarks
{
    private const int Iterations = 1_000;

    private static readonly Color[] Colors =
    [
        Color.White, Color.DarkBlue, Color.Black, Color.LightGreen, Color.DarkRed,
        Color.Navy, Color.Teal, Color.Wheat, Color.LightSalmon, Color.DarkSlateGray,
    ];

    private static readonly string[] AddressStrings =
    [
        "A1", "B2", "AB123", "$C$5", "$D10", "E$20", "XFD1048576", "Z999", "AA100", "BC4567",
    ];

    private static readonly XLAddress[] Addresses =
    [
        new(1, 1, false, false),
        new(2, 2, false, false),
        new(123, 28, false, false),
        new(5, 3, true, true),
        new(10, 4, false, true),
        new(20, 5, true, false),
        new(1048576, 16384, false, false),
        new(999, 26, false, false),
        new(100, 27, true, true),
        new(4567, 55, false, false),
    ];

    private static readonly double[] Numbers =
    [
        1234.56, 0.1, 9999.99, 42, 1000000.5, 0.007, 88.88, 12345.678, 500, 3.14159,
    ];

    private static readonly string[] SheetNames =
    [
        "Sheet1", "Data", "My Sheet", "Report 2024", "Summary", "Q1'23", "Prices", "A1B2", "_hidden", "Notes",
    ];

    private static readonly string[] TextValues =
    [
        "plain text with no newline", "another simple value", "Item 123 - 0456",
        "multi\nline\nvalue", "trailing space ", "CODE-00042", "North", "Active",
        "a longer note for this row without any line breaks at all", "Cat-7",
    ];

    private static readonly string[] Formulas =
    [
        "A1+B2*C3", "SUM(B2:B100)+D5", "IF(A1>0,\"yes\",\"no\")", "VLOOKUP(A1,B1:C100,2,FALSE)",
        "E10-F20/G30", "AVERAGE(H1:H50)", "\"literal:A1\"&B2", "C5*D5+E5", "MAX(A1:Z1)", "MIN(B2:B200)",
    ];

    private XLWorkbook _workbook = null!;
    private XLWorksheet _worksheet = null!;
    private XLRange _shiftRange = null!;
    private XLCell[] _emptyCells = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _workbook = new XLWorkbook();
        _worksheet = (XLWorksheet)_workbook.AddWorksheet("Data");
        _shiftRange = (XLRange)_worksheet.Range("A1:Z1000");

        _emptyCells = new XLCell[Colors.Length];
        for (var i = 0; i < _emptyCells.Length; i++)
            _emptyCells[i] = (XLCell)_worksheet.Cell(i + 1, 40); // untouched, still empty
    }

    [GlobalCleanup]
    public void Cleanup() => _workbook.Dispose();

    // #5 ColorExtensions.ToHex
    [Benchmark]
    public int ToHex()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += Colors[i % Colors.Length].ToHex().Length;
        return sum;
    }

    // #6 XLAddress.ToString
    [Benchmark]
    public int AddressToString()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += Addresses[i % Addresses.Length].ToString().Length;
        return sum;
    }

    // #7 XLAddress.Create
    [Benchmark]
    public int AddressCreate()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += XLAddress.Create(null, AddressStrings[i % AddressStrings.Length]).ColumnNumber;
        return sum;
    }

    // #1 FormatExtensions.ToExcelFormat
    [Benchmark]
    public int ToExcelFormat()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += ((object)Numbers[i % Numbers.Length]).ToExcelFormat("#,##0.00", CultureInfo.InvariantCulture).Length;
        return sum;
    }

    // #10 StringExtensions.EscapeSheetName
    [Benchmark]
    public int EscapeSheetName()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += SheetNames[i % SheetNames.Length].EscapeSheetName().Length;
        return sum;
    }

    // #9 StringExtensions.FixNewLines
    [Benchmark]
    public int FixNewLines()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += TextValues[i % TextValues.Length].FixNewLines().Length;
        return sum;
    }

    // #8 StringExtensions.CharCount
    [Benchmark]
    public int CharCount()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += Formulas[i % Formulas.Length].CharCount('"');
        return sum;
    }

    // #3 XLCellFormulaShifter.ShiftFormulaRows (also exercises span quote counting)
    [Benchmark]
    public int ShiftFormulaRows()
    {
        var sum = 0;
        for (var i = 0; i < Iterations; i++)
            sum += XLCellFormulaShifter.ShiftFormulaRows(Formulas[i % Formulas.Length], _worksheet, _shiftRange, 5).Length;
        return sum;
    }

    // #2 XLCell.IsEmpty -> IsInsideConditionalFormat (SelectMany/Any removal)
    [Benchmark]
    public int IsEmptyConditionalFormats()
    {
        var count = 0;
        for (var i = 0; i < Iterations; i++)
        {
            if (_emptyCells[i % _emptyCells.Length].IsEmpty(XLCellsUsedOptions.All))
                count++;
        }
        return count;
    }
}
