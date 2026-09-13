using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Extensions;
using XLibur.Tests.Excel.IO;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The error values a cell can hold, and the text that stands for each one in a formula and in a file.
/// </summary>
public class XLErrorTests
{
    public static IEnumerable<XLError> AllErrors() => Enum.GetValues<XLError>();

    /// <summary>
    /// A member is numbered by the <c>errorType</c> that [MS-XLSX] 2.3.6.1.3 gives its error's
    /// <c>_error</c> rich value, which is one less than <c>ERROR.TYPE</c>. <c>#PYTHON!</c> shares
    /// <c>errorType</c> 18 with <c>#EXTERNAL!</c>, so it takes 20, which the spec does not use.
    /// </summary>
    [Test]
    [Arguments("#NULL!", XLError.NullValue, 0)]
    [Arguments("#DIV/0!", XLError.DivisionByZero, 1)]
    [Arguments("#VALUE!", XLError.IncompatibleValue, 2)]
    [Arguments("#REF!", XLError.CellReference, 3)]
    [Arguments("#NAME?", XLError.NameNotRecognized, 4)]
    [Arguments("#NUM!", XLError.NumberInvalid, 5)]
    [Arguments("#N/A", XLError.NoValueAvailable, 6)]
    [Arguments("#GETTING_DATA", XLError.GettingData, 7)]
    [Arguments("#SPILL!", XLError.SpillRange, 8)]
    [Arguments("#CONNECT!", XLError.Connect, 9)]
    [Arguments("#BLOCKED!", XLError.Blocked, 10)]
    [Arguments("#UNKNOWN!", XLError.Unknown, 11)]
    [Arguments("#FIELD!", XLError.Field, 12)]
    [Arguments("#CALC!", XLError.Calc, 13)]
    [Arguments("#BUSY!", XLError.Busy, 14)]
    [Arguments("#EXTERNAL!", XLError.External, 18)]
    [Arguments("#TIMEOUT!", XLError.Timeout, 19)]
    [Arguments("#PYTHON!", XLError.Python, 20)]
    public async Task Each_error_text_has_a_member_numbered_by_its_errorType(string text, XLError expected, int errorType)
    {
        await Assert.That(XLErrorParser.TryParseError(text, out var parsed)).IsTrue();
        await Assert.That(parsed).IsEqualTo(expected);
        await Assert.That((int)parsed).IsEqualTo(errorType);
    }

    /// <summary>Every member has a text, and that text reads back as the same member.</summary>
    [Test]
    [MethodDataSource(nameof(AllErrors))]
    public async Task Each_member_round_trips_through_its_text(XLError error)
    {
        await Assert.That(XLErrorParser.TryParseError(error.ToDisplayString(), out var parsed)).IsTrue();
        await Assert.That(parsed).IsEqualTo(error);
    }

    /// <summary>
    /// A cell error is written to the file as its text and read back from it. An error text the reader
    /// could not map used to load as a blank cell, with no warning.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(AllErrors))]
    public async Task A_cell_error_survives_a_save_and_load(XLError error)
    {
        using var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet().Cell("A1").Value = error;
            wb.SaveAs(package);
        }

        using var reloaded = new XLWorkbook(package);

        await Assert.That(reloaded.Worksheet(1).Cell("A1").Value).IsEqualTo(error);
    }

    /// <summary>
    /// A pivot cache stores an error in a field as its text, both among the field's shared items and
    /// in each record. An error text the reader could not map failed the whole load, and so did any
    /// error in a record, because the records were written with the element for a boolean.
    /// </summary>
    [Test]
    [Arguments(XLError.DivisionByZero)]
    [Arguments(XLError.Calc)]
    public async Task A_pivot_cache_holding_an_error_survives_a_save_and_load(XLError error)
    {
        var text = error.ToDisplayString();
        using var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = "Name";
            data.Cell("B1").Value = "Qty";
            data.Cell("A2").Value = error;
            data.Cell("B2").Value = 1;
            data.Cell("A3").Value = "Pie";
            data.Cell("B3").Value = 2;

            var pt = data.Range("A1:B3").CreatePivotTable(wb.AddWorksheet("Pivot").FirstCell(), "pt");
            pt.RowLabels.Add("Name");
            pt.Values.Add("Qty");
            wb.SaveAs(package);
        }

        // Guards the premise: the saved cache really holds the error.
        var partNames = package.PartNames();
        var cacheParts = partNames.Where(n => Regex.IsMatch(n, @"(^|/)pivotCacheDefinition\d*\.xml$")).ToArray();
        await Assert.That(cacheParts.Length).IsEqualTo(1).Because(string.Join(", ", partNames));
        await Assert.That(package.PartXml(cacheParts[0])).Contains($"\"{text}\"");

        using var reloaded = new XLWorkbook(package);

        var cache = (XLPivotCache)reloaded.PivotCaches.Single();
        await Assert.That(cache.TryGetFieldIndex("Name", out var name)).IsTrue();
        var names = cache.GetFieldSharedItems(name).GetCellValues();
        await Assert.That(names.Any(v => v.IsError && v.GetError() == error)).IsTrue();
    }

    /// <summary>A defined name can hold a newer error, and a formula that uses the name returns it.</summary>
    [Test]
    public async Task A_defined_name_can_hold_a_newer_error()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("Missing", "#FIELD!");
        ws.Cell("A1").FormulaA1 = "Missing";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.Field);
    }

    /// <summary>
    /// A cell value refuses a number that stands for no error. The numbering has gaps, because it
    /// follows [MS-XLSX], so a range check alone would let 15 to 17 through.
    /// </summary>
    [Test]
    [Arguments(-1)]
    [Arguments(15)]
    [Arguments(16)]
    [Arguments(17)]
    [Arguments(21)]
    public async Task A_cell_value_refuses_a_number_that_is_no_error(int value)
    {
        await Assert.That(() => (XLCellValue)(XLError)value).Throws<ArgumentOutOfRangeException>();
    }
}
