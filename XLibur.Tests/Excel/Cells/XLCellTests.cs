using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Cells;
// ReSharper disable once InconsistentNaming
public class XLCellTests
{
    [SuppressMessage("ReSharper", "RedundantCast")]
    private static readonly object[] AllNumberTypes =
    [
        (sbyte)1,
        (byte)2,
        (short)3,
        (ushort)4,
        (int)5,
        (uint)6,
        (long)7,
        (ulong)8,
        (float)9.5f,
        (double)10.75,
        (decimal)11.875m
    ];

    private static readonly string[] InsertDataStrings = ["a", "b", "c"];

    [Test]
    public async Task CellsUsed()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Cell(1, 1);
        ws.Cell(2, 2);
        var count = ws.Range("A1:B2").CellsUsed().Count();
        await Assert.That(count).IsEqualTo(0);
    }

    [Test]
    public async Task CellsUsedIncludeStyles1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Row(3).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Column(3).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Cell(2, 2).Value = "ASDF";
        var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
        await Assert.That(range).IsEqualTo("B2:C3");
    }

    [Test]
    public async Task CellsUsedIncludeStyles2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Row(2).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Column(2).Style.Fill.BackgroundColor = XLColor.Red;
        ws.Cell(3, 3).Value = "ASDF";
        var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
        await Assert.That(range).IsEqualTo("B2:C3");
    }

    [Test]
    public async Task CellsUsedIncludeStyles3()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var range = ws.RangeUsed(XLCellsUsedOptions.All);
        await Assert.That(range).IsNull();
    }

    [Test]
    public async Task CellUsedIncludesSparklines()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.Range("C3:E4").Value = 1;
        ws.SparklineGroups.Add("B2", "C3:E3");
        ws.SparklineGroups.Add("F5", "C4:E4");

        var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
        await Assert.That(range).IsEqualTo("B2:F5");
    }

    [Test]
    public async Task GetValue_Nullable()
    {
        var cell = new XLWorkbook().AddWorksheet().FirstCell();

        await Assert.That(cell.Clear().GetValue<double?>()).IsNull();
        await Assert.That(cell.SetValue(1.5).GetValue<double?>()).IsEqualTo(1.5);
        await Assert.That(cell.SetValue(2).GetValue<int?>()).IsEqualTo(2);
        await Assert.That(cell.SetValue(Blank.Value).GetValue<double?>()).IsNull();
        await Assert.That(() => cell.SetValue("text").GetValue<double?>()).Throws<InvalidCastException>();
    }

    [Test]
    public async Task InsertData1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var range = ws.Cell(2, 2).InsertData(InsertDataStrings);
        await Assert.That(range.ToString()).IsEqualTo("Sheet1!B2:B4");
    }

    [Test]
    public async Task InsertData_DoesntTransposeDataOnFalseFlag()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var range = ws.Cell(2, 2).InsertData(InsertDataStrings, false);
        await Assert.That(range.ToString()).IsEqualTo("Sheet1!B2:B4");
    }

    [Test]
    public async Task InsertData_TransposesDataOnTrueFlag()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var range = ws.Cell(2, 2).InsertData(InsertDataStrings, true);
        await Assert.That(range.ToString()).IsEqualTo("Sheet1!B2:D2");
    }

    [Test]
    public async Task InsertData_DifferentTypes()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        object[] values = ["Text", 45, DateTime.Today, true, "More text"];

        ws.FirstCell().InsertData(values);

        await Assert.That(ws.FirstCell().GetString()).IsEqualTo("Text");
        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(45);
        await Assert.That(ws.Cell("A3").GetDateTime()).IsEqualTo(DateTime.Today);
        await Assert.That(ws.Cell("A4").GetBoolean()).IsTrue();
        await Assert.That(ws.Cell("A5").GetString()).IsEqualTo("More text");
        await Assert.That(ws.Cell("A6").IsEmpty()).IsTrue();
    }

    [Test]
    public async Task InsertData_with_Guids()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.FirstCell().InsertData(Enumerable.Range(1, 20).Select(i => new { Guid = Guid.NewGuid() }));

        await Assert.That(ws.FirstCell().DataType).IsEqualTo(XLDataType.Text);
        await Assert.That(ws.FirstCell().GetText().Length).IsEqualTo(Guid.NewGuid().ToString().Length);
    }

    [Test]
    public async Task InsertData_with_Nulls()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");

        var table = new DataTable();
        table.TableName = "Patients";
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(21, "Combivent", DBNull.Value, new DateTime(2000, 1, 4, 0, 0, 0, DateTimeKind.Unspecified));
        table.Rows.Add(100, "Dilantin", "Melanie", DBNull.Value);

        ws.FirstCell().InsertData(table);

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(25);
        await Assert.That(ws.Cell("C4").Value.IsBlank).IsTrue();
        await Assert.That(ws.Cell("D5").Value.IsBlank).IsTrue();
    }

    [Test]
    public async Task InsertData_with_Nulls_IEnumerable()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");

        var dateTimeList = new List<DateTime?>
        {
            new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
            new DateTime(2000, 1, 2, 0, 0, 0, DateTimeKind.Unspecified),
            new DateTime(2000, 1, 3, 0, 0, 0, DateTimeKind.Unspecified),
            new DateTime(2000, 1, 4, 0, 0, 0, DateTimeKind.Unspecified),
            null
        };

        ws.FirstCell().InsertData(dateTimeList);

        await Assert.That(ws.Cell("A1").GetDateTime()).IsEqualTo(new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified));
        await Assert.That(ws.Cell("A5").Value).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task InsertData_AllNumberTypes_AreInsertedAsNumbers()
    {
        var ws = new XLWorkbook().Worksheets.Add();

        ws.FirstCell().InsertData(AllNumberTypes);

        for (var row = 1; row <= AllNumberTypes.Length; ++row)
        {
            var expectedValue = Convert.ChangeType(AllNumberTypes[row - 1], typeof(double));
            var actualValue = ws.Cell(row, 1).Value;
            await Assert.That(actualValue).IsEqualTo(ExpectedCellValue.From(expectedValue));
        }
    }

    [Test]
    public async Task InsertTable_AllNumberTypes_AreInsertedAsNumbers()
    {
        var ws = new XLWorkbook().Worksheets.Add();

        var table = new DataTable("Numbers");
        foreach (var number in AllNumberTypes)
        {
            var numberType = number.GetType();
            table.Columns.Add(numberType.Name, numberType);
        }

        table.Rows.Add(AllNumberTypes);

        ws.FirstCell().InsertTable(table);

        for (var column = 1; column <= AllNumberTypes.Length; ++column)
        {
            var expectedValue = Convert.ChangeType(AllNumberTypes[column - 1], typeof(double));
            var actualValue = ws.Cell(2, column).Value;
            await Assert.That(actualValue).IsEqualTo(ExpectedCellValue.From(expectedValue));
        }
    }

    [Test]
    public async Task IsEmpty1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        var actual = cell.IsEmpty();
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        var actual = cell.IsEmpty(XLCellsUsedOptions.All);
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty3()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        var actual = cell.IsEmpty();
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty4()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        var actual = cell.IsEmpty(XLCellsUsedOptions.AllContents);
        var expected = true;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty5()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Style.Fill.BackgroundColor = XLColor.Red;
        var actual = cell.IsEmpty(XLCellsUsedOptions.All);
        var expected = false;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task IsEmpty6()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Value = "X";
        var actual = cell.IsEmpty();
        var expected = false;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task NaN_is_not_a_number()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1");
        cell.Value = "NaN";

        await Assert.That(cell.DataType).IsNotEqualTo(XLDataType.Number);
    }

    [Test]
    public async Task Nan_is_not_a_number()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1");
        cell.Value = "Nan";

        await Assert.That(cell.DataType).IsNotEqualTo(XLDataType.Number);
    }

    [Test]
    public async Task TryGetValue_Boolean_Bad()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue("ABC");
        var success = cell.TryGetValue(out bool outValue);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public async Task TryGetValue_Boolean_False()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue(false);
        var success = cell.TryGetValue(out bool outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsFalse();
    }

    [Test]
    public async Task TryGetValue_Boolean_FalseText()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue("False");
        var success = cell.TryGetValue(out bool outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsFalse();
    }

    [Test]
    public async Task TryGetValue_Boolean_True()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue(true);
        var success = cell.TryGetValue(out bool outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsTrue();
    }

    [Test]
    public async Task TryGetValue_Boolean_TrueText()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue("True");
        var success = cell.TryGetValue(out bool outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsTrue();
    }

    [Test]
    public async Task TryGetValue_DateTime_Good2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var success = ws.Cell("A1").SetFormulaA1("=TODAY() + 10").TryGetValue(out DateTime outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(DateTime.Today.AddDays(10));
    }

    [Test]
    public async Task TryGetValue_DateTime_BadButFormulaGood()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var success = ws.Cell("A1").SetFormulaA1("=\"44\"&\"020\"").TryGetValue(out DateTime outValue);
        await Assert.That(success).IsFalse();

        ws.Cell("B1").SetFormulaA1("=A1+1");

        success = ws.Cell("B1").TryGetValue(out outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(new DateTime(2020, 07, 09, 0, 0, 0, DateTimeKind.Unspecified));
    }

    [Test]
    public async Task TryGetValue_DateTime_BadString()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var date = "ABC";
        var success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime outValue);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public async Task TryGetValue_DateTime_SerialDateTimeOutsideRange()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var serialDateTimeOutsideRange = 5545454;
        ws.FirstCell().SetValue(serialDateTimeOutsideRange);
        var success = ws.FirstCell().TryGetValue(out DateTime _);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public async Task TryGetValue_Enum_Good()
    {
        var ws = new XLWorkbook().AddWorksheet();
        await Assert.That(ws.FirstCell().SetValue(nameof(NumberStyles.AllowCurrencySymbol)).TryGetValue(out NumberStyles value)).IsTrue();
        await Assert.That(value).IsEqualTo(NumberStyles.AllowCurrencySymbol);

        // Nullable alternative
        await Assert.That(ws.FirstCell().SetValue(nameof(NumberStyles.AllowCurrencySymbol)).TryGetValue(out NumberStyles? value2)).IsTrue();
        await Assert.That(value2).IsEqualTo(NumberStyles.AllowCurrencySymbol);
    }

    [Test]
    public async Task TryGetValue_Enum_BadString()
    {
        var ws = new XLWorkbook().AddWorksheet();
        await Assert.That(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles value)).IsFalse();
        await Assert.That(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles? value2)).IsFalse();
    }

    [Test]
    public async Task TryGetValue_TimeSpan_BadString()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var timeSpan = "ABC";
        var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public async Task TryGetValue_TimeSpan_Good()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var timeSpan = new TimeSpan(1, 1, 1);
        var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(timeSpan);
    }

    [Test]
    public async Task TryGetValue_TimeSpan_Good2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var success = ws.Cell("A1").SetValue(0.0034722222222222199).TryGetValue(out TimeSpan outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(TimeSpan.FromMinutes(5));
    }

    [Test]
    public async Task TryGetValue_TimeSpan_Good_Large()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var timeSpan = TimeSpan.FromMilliseconds((double)int.MaxValue + 1);
        var success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(timeSpan);
    }

    [Test]
    [SetCulture("en-US")]
    public async Task TryGetValue_TimeSpan_Good_FromText()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var success = ws.Cell("A1").SetValue("300:14:50.453").TryGetValue(out TimeSpan outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(new TimeSpan(12, 12, 14, 50, 453));
    }

    [Test]
    public async Task TryGetValue_sbyte_Bad2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue("255");
        var success = cell.TryGetValue(out sbyte outValue);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public async Task TryGetValue_sbyte_Good()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell("A1").SetValue(5);
        var success = cell.TryGetValue(out sbyte outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo(ExpectedCellValue.From(5));
    }

    [Test]
    public async Task TryGetValue_Unicode_String()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");

        var success = ws.Cell("A1")
            .SetValue("Site_x0020_Column_x0020_Test")
            .TryGetValue(out string outValue);
        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo("Site Column Test");

        success = ws.Cell("A1")
            .SetValue("Site_x005F_x0020_Column_x005F_x0020_Test")
            .TryGetValue(out outValue);

        await Assert.That(success).IsTrue();
        await Assert.That(outValue).IsEqualTo("Site_x005F_x0020_Column_x005F_x0020_Test");
    }

    [Test]
    public async Task TryGetValue_Nullable()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Clear();
        ws.Cell("A2").SetValue(1.5);
        ws.Cell("A3").SetValue(2.5.ToString(CultureInfo.CurrentCulture));
        ws.Cell("A4").SetValue("text");

        await Assert.That(ws.Cell("A1").TryGetValue(out double? _)).IsTrue();
        await Assert.That(ws.Cell("A2").TryGetValue(out double? _)).IsTrue();
        await Assert.That(ws.Cell("A3").TryGetValue(out double? _)).IsTrue();
        await Assert.That(ws.Cell("A4").TryGetValue(out double? _)).IsFalse();
    }

    [Test]
    public async Task CopyRangeAtCellAddress()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");

        ws.Cell("A1").SetValue(2)
            .CellRight().SetValue(3)
            .CellRight().SetValue(5)
            .CellRight().SetValue(7);

        var range = ws.Range("1:1");

        ws.Cell("B2").CopyFrom(range);

        await Assert.That(ws.Cell("B2").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(3);
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(5);
        await Assert.That(ws.Cell("E2").Value).IsEqualTo(7);
    }

    [Test]
    public async Task ValueSetToEmptyString()
    {
        var expected = string.Empty;

        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.Value = new DateTime(2000, 1, 2, 0, 0, 0, DateTimeKind.Unspecified);
        cell.Value = string.Empty;
        await Assert.That(cell.GetText()).IsEqualTo(expected);
        await Assert.That(cell.Value).IsEqualTo(expected);

        cell.Value = new DateTime(2000, 1, 2, 0, 0, 0, DateTimeKind.Unspecified);
        cell.SetValue(string.Empty);
        await Assert.That(cell.GetText()).IsEqualTo(expected);
        await Assert.That(cell.Value).IsEqualTo(expected);
    }

    [Test]
    public async Task ValueSetDateWithShortUserDateFormat()
    {
        // For this test to make sense, user's local date format should be dd/MM/yy (note without the 2 century digits)
        // What happened previously was that the century digits got lost in .ToString() conversion and wrong century was sometimes returned.
        var ci = new CultureInfo(CultureInfo.InvariantCulture.LCID)
        {
            DateTimeFormat =
            {
                ShortDatePattern = "dd/MM/yy"
            }
        };
        Thread.CurrentThread.CurrentCulture = ci;
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        var expected = DateTime.Today.AddYears(20);
        cell.Value = expected;
        var actual = (DateTime)cell.Value;
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    public async Task SetStringValueTooLong()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.FirstCell().Value = new DateTime(2018, 5, 15, 0, 0, 0, DateTimeKind.Unspecified);

        ws.FirstCell().SetValue(new string('A', 32767));

        await Assert.That(() => ws.FirstCell().Value = new string('A', 32768)).Throws<ArgumentOutOfRangeException>();
        await Assert.That(() => ws.FirstCell().SetValue(new string('A', 32768))).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Load_InlineString_entities_decoded_text_is_at_limit()
    {
        // Inline string XML contains &#xA; entities making raw XML > 32767 chars,
        // but decoded text is exactly 32767 characters. Should load without error.
        // FixNewLines() on Windows converts \n to \r\n, making .Length larger,
        // but the Excel-visible length (not counting \r) must be 32767.
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\InlineStringEntitiesAtLimit.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var text = ws.Cell(1, 1).Value.GetText();
        var excelLength = text.Length - text.AsSpan().Count('\r');
        await Assert.That(excelLength).IsEqualTo(32767);
    }

    [Test]
    public async Task SetCellValueWipesFormulas()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.FirstCell().FormulaA1 = "=TODAY()";
        ws.FirstCell().Value = "hello world";
        await Assert.That(ws.FirstCell().HasFormula).IsFalse();

        ws.FirstCell().FormulaA1 = "=TODAY()";
        ws.FirstCell().SetValue("hello world");
        await Assert.That(ws.FirstCell().HasFormula).IsFalse();
    }

    [Test]
    public async Task CellValueLineWrapping()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.FirstCell().Value = "hello world";
        await Assert.That(ws.FirstCell().Style.Alignment.WrapText).IsFalse();

        ws.FirstCell().Value = "hello\r\nworld";
        await Assert.That(ws.FirstCell().Style.Alignment.WrapText).IsTrue();

        ws.FirstCell().Style.Alignment.WrapText = false;

        ws.FirstCell().SetValue("hello world");
        await Assert.That(ws.FirstCell().Style.Alignment.WrapText).IsFalse();

        ws.FirstCell().SetValue("hello\r\nworld");
        await Assert.That(ws.FirstCell().Style.Alignment.WrapText).IsTrue();
    }

    [Test]
    public async Task TestInvalidXmlCharacters()
    {
        byte[] data;

        using (var stream = new MemoryStream())
        {
            var wb = new XLWorkbook();
            wb.AddWorksheet("Sheet1").FirstCell().SetValue("\u0018");
            wb.SaveAs(stream);
            data = stream.ToArray();
        }

        using (var stream = new MemoryStream(data))
        {
            var wb = new XLWorkbook(stream);
            await Assert.That(wb.Worksheets.First().FirstCell().Value).IsEqualTo("\u0018");
        }
    }

    [Test]
    public async Task CanClearDateTimeCellValue()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell();
            c.SetValue(new DateTime(2017, 10, 08, 0, 0, 0, DateTimeKind.Unspecified));
            await Assert.That(c.DataType).IsEqualTo(XLDataType.DateTime);
            await Assert.That(c.Value).IsEqualTo(new DateTime(2017, 10, 08, 0, 0, 0, DateTimeKind.Unspecified));

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var c = ws.FirstCell();
            await Assert.That(c.DataType).IsEqualTo(XLDataType.DateTime);
            await Assert.That(c.Value).IsEqualTo(new DateTime(2017, 10, 08, 0, 0, 0, DateTimeKind.Unspecified));

            c.Clear();
            wb.Save();
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var c = ws.FirstCell();
            await Assert.That(c.DataType).IsEqualTo(XLDataType.Blank);
            await Assert.That(c.IsEmpty()).IsTrue();
        }
    }

    [Test]
    public async Task ClearCellRemovesSparkline()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        ws.SparklineGroups.Add("B1:B3", "C1:E3");

        ws.Cell("B1").Clear();
        ws.Cell("B2").Clear(XLClearOptions.Sparklines);

        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(1);
        await Assert.That(ws.Cell("B1").HasSparkline).IsFalse();
        await Assert.That(ws.Cell("B2").HasSparkline).IsFalse();
        await Assert.That(ws.Cell("B3").HasSparkline).IsTrue();
    }

    [Test]
    public async Task CurrentRegion()
    {
        // Partially based on sample in https://github.com/XLibur/XLibur/issues/120
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("B1").SetValue("x")
            .CellBelow().SetValue("x")
            .CellBelow().SetValue("x");

        ws.Cell("C1").SetValue("x")
            .CellBelow().SetValue("x")
            .CellBelow().SetValue("x");

        //Deliberately D2
        ws.Cell("D2").SetValue("x")
            .CellBelow().SetValue("x");

        ws.Cell("G1").SetValue("x")
            .CellBelow() // skip a cell
            .CellBelow().SetValue("x")
            .CellBelow().SetValue("x");

        // Deliberately H2
        ws.Cell("H2").SetValue("x")
            .CellBelow().SetValue("x")
            .CellBelow().SetValue("x");

        // A diagonal
        ws.Cell("E8").SetValue("x")
            .CellBelow().CellRight().SetValue("x")
            .CellBelow().CellRight().SetValue("x")
            .CellBelow().CellRight().SetValue("x")
            .CellBelow().CellRight().SetValue("x");

        await Assert.That(ws.Cell("A10").CurrentRegion.RangeAddress.ToString()).IsEqualTo("A10:A10");
        await Assert.That(ws.Cell("B5").CurrentRegion.RangeAddress.ToString()).IsEqualTo("B5:B5");
        await Assert.That(ws.Cell("P1").CurrentRegion.RangeAddress.ToString()).IsEqualTo("P1:P1");

        await Assert.That(ws.Cell("D3").CurrentRegion.RangeAddress.ToString()).IsEqualTo("B1:D3");
        await Assert.That(ws.Cell("D4").CurrentRegion.RangeAddress.ToString()).IsEqualTo("B1:D4");
        await Assert.That(ws.Cell("E4").CurrentRegion.RangeAddress.ToString()).IsEqualTo("B1:E4");

        foreach (var c in ws.Range("B1:D3").Cells())
        {
            await Assert.That(c.CurrentRegion.RangeAddress.ToString()).IsEqualTo("B1:D3");
        }

        foreach (var c in ws.Range("A1:A3").Cells())
        {
            await Assert.That(c.CurrentRegion.RangeAddress.ToString()).IsEqualTo("A1:D3");
        }

        await Assert.That(ws.Cell("A4").CurrentRegion.RangeAddress.ToString()).IsEqualTo("A1:D4");

        foreach (var c in ws.Range("E1:E3").Cells())
        {
            await Assert.That(c.CurrentRegion.RangeAddress.ToString()).IsEqualTo("B1:E3");
        }

        await Assert.That(ws.Cell("E4").CurrentRegion.RangeAddress.ToString()).IsEqualTo("B1:E4");

        //// SECOND REGION
        foreach (var c in ws.Range("F1:F4").Cells())
        {
            await Assert.That(c.CurrentRegion.RangeAddress.ToString()).IsEqualTo("F1:H4");
        }

        await Assert.That(ws.Cell("F5").CurrentRegion.RangeAddress.ToString()).IsEqualTo("F1:H5");

        //// DIAGONAL
        await Assert.That(ws.Cell("E8").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");
        await Assert.That(ws.Cell("F9").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");
        await Assert.That(ws.Cell("G10").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");
        await Assert.That(ws.Cell("H11").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");
        await Assert.That(ws.Cell("I12").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");

        await Assert.That(ws.Cell("G9").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");
        await Assert.That(ws.Cell("F10").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:I12");

        await Assert.That(ws.Cell("D7").CurrentRegion.RangeAddress.ToString()).IsEqualTo("D7:I12");
        await Assert.That(ws.Cell("J13").CurrentRegion.RangeAddress.ToString()).IsEqualTo("E8:J13");

        // Four corners of a sheet
        await Assert.That(ws.Cell(1, 1).CurrentRegion.RangeAddress.ToString()).IsEqualTo("A1:D3");
        await Assert.That(ws.Cell(1, XLHelper.MaxColumnNumber).CurrentRegion.RangeAddress.ToString()).IsEqualTo("XFD1:XFD1");
        await Assert.That(ws.Cell(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber).CurrentRegion.RangeAddress.ToString()).IsEqualTo("XFD1048576:XFD1048576");
        await Assert.That(ws.Cell(XLHelper.MaxRowNumber, 1).CurrentRegion.RangeAddress.ToString()).IsEqualTo("A1048576:A1048576");
    }

    // https://github.com/XLibur/XLibur/issues/630
    [Test]
    public async Task ConsiderEmptyValueAsNumericInSumFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").SetValue("Empty");
        ws.Cell("A2").SetValue("Numeric");
        ws.Cell("A3").SetValue("Copy of numeric");

        ws.Cell("B2").SetFormulaA1("=B1");
        ws.Cell("B3").SetFormulaA1("=B2");

        ws.Cell("C2").SetFormulaA1("=SUM(C1)");
        ws.Cell("C3").SetFormulaA1("=C2");

        var b1 = ws.Cell("B1").Value;
        var b2 = ws.Cell("B2").Value;
        var b3 = ws.Cell("B3").Value;

        await Assert.That(b1).IsEqualTo(Blank.Value);
        await Assert.That(b2).IsEqualTo(0);
        await Assert.That(b3).IsEqualTo(0);

        var c1 = ws.Cell("C1").Value;
        var c2 = ws.Cell("C2").Value;
        var c3 = ws.Cell("C3").Value;

        await Assert.That(c1).IsEqualTo(Blank.Value);
        await Assert.That(c2).IsEqualTo(0);
        await Assert.That(c3).IsEqualTo(0);
    }

    [Test]
    public async Task SetFormulaA1AffectsR1C1()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.FormulaR1C1 = "R[1]C";

        cell.FormulaA1 = "B2";

        await Assert.That(cell.FormulaR1C1).IsEqualTo("R[1]C[1]");
    }

    [Test]
    public async Task SetFormulaR1C1AffectsA1()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.FormulaA1 = "A2";

        cell.FormulaR1C1 = "R[1]C[1]";

        await Assert.That(cell.FormulaA1).IsEqualTo("B2");
    }

    [Test]
    [Arguments(" = 1 + SUM({ 1; 7})  - A8  ", "1 + SUM({ 1; 7})  - A8")]
    public async Task FormulaA1_setter_trims_and_removes_equal_if_present(string formula, string expectedResult)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = formula;
        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments(" =  1 +   R[1]C[7]  ", "1 +   R[1]C[7]")]
    public async Task FormulaR1C1_setter_trims_and_removes_equal_if_present(string formula, string expectedResult)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaR1C1 = formula;
        await Assert.That(ws.Cell("A1").FormulaR1C1).IsEqualTo(expectedResult);
    }

    [Test]
    public async Task FormulaWithCircularReferenceFails()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var A1 = ws.Cell("A1");
        var A2 = ws.Cell("A2");
        A1.FormulaA1 = "A2 + 1";
        A2.FormulaA1 = "A1 + 1";

        await Assert.That(() => _ = A1.Value).Throws<Exception>();
        await Assert.That(() => _ = A2.Value).Throws<Exception>();
    }

    [Test]
    public async Task InvalidFormulaShiftProducesREF()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").Value = 2;
            ws.Cell("B2").FormulaA1 = "=A1+B1";

            await Assert.That(ws.Cell("B2").Value).IsEqualTo(3);

            ws.Range("B2").CopyTo(ws.Range("A2"));
            var fA2 = ws.Cell("A2").FormulaA1;

            wb.SaveAs(ms);

            await Assert.That(fA2).IsEqualTo("#REF!+A1");
        }

        using (var wb2 = new XLWorkbook(ms))
        {
            var fA2 = wb2.Worksheets.First().Cell("A2").FormulaA1;
            await Assert.That(fA2).IsEqualTo("#REF!+A1");
        }
    }

    [Test]
    public async Task FormulaWithCircularReferenceFails2()
    {
        var cell = new XLWorkbook().Worksheets.Add("Sheet1").FirstCell();
        cell.FormulaA1 = "A1";
        await Assert.That(() =>
        {
            var _ = cell.Value;
        }).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task TryGetValueFormula_EvaluationFail_ReturnFalse()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var A1 = ws.Cell("A1");
        var A2 = ws.Cell("A2");
        var A3 = ws.Cell("A3");
        A1.FormulaA1 = "A2 + 1";
        A2.FormulaA1 = "A1 + 1";

        await Assert.That(A1.TryGetValue(out string _)).IsFalse();
        await Assert.That(A2.TryGetValue(out string _)).IsFalse();
        await Assert.That(A3.TryGetValue(out string _)).IsTrue();
    }

    [Test]
    public async Task ToStringNoFormatString()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var c = ws.FirstCell().CellBelow(2).CellRight(3);

        await Assert.That(c.ToString()).IsEqualTo("D3");
    }

    [Test]
    [Arguments("D3", "A")]
    [Arguments("YEAR(DATE(2018, 1, 1))", "F")]
    [Arguments("YEAR(DATE(2018, 1, 1))", "f")]
    [Arguments("0000.00", "NF")]
    [Arguments("0000.00", "nf")]
    [Arguments("FFFF0000", "fg")]
    [Arguments("Color Theme: Accent5, Tint: 0", "BG")]
    [Arguments("2018.00", "v")]
    public async Task ToStringFormatString(string expected, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var c = ws.FirstCell().CellBelow(2).CellRight(3);

        var formula = "YEAR(DATE(2018, 1, 1))";
        c.FormulaA1 = formula;

        var numberFormat = "0000.00";
        c.Style.NumberFormat.Format = numberFormat;

        c.Style.Font.FontColor = XLColor.Red;
        c.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);

        await Assert.That(c.ToString(format)).IsEqualTo(expected);

        await Assert.That(() => c.ToString("dummy")).Throws<FormatException>();
    }

    [Test]
    public async Task ToStringInvalidFormat()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var c = ws.FirstCell();

        await Assert.That(() => c.ToString("dummy")).Throws<FormatException>();
    }

    [Test]
    public async Task Property_Active_is_true_when_cell_has_same_address_as_active_cell_in_worksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.ActiveCell).IsNull();
        await Assert.That(ws.Cell(1, 1).Active).IsFalse();

        ws.ActiveCell = ws.Cell("C4");
        await Assert.That(ws.Cell("C4").Active).IsTrue();
        await Assert.That(ws.Cell("C5").Active).IsFalse();

        ws.ActiveCell = null;
        await Assert.That(ws.Cell("C4").Active).IsFalse();
    }

    [Test]
    public async Task Property_Active_deactivates_cell_only_when_the_cell_is_active()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.ActiveCell = ws.Cell("A2");

        ws.Cell("B2").Active = false;
        await Assert.That(ws.ActiveCell).IsEqualTo(ws.Cell("A2"));

        ws.Cell("A2").Active = false;
        await Assert.That(ws.ActiveCell).IsNull();
    }

    [Test]
    public async Task Property_Active_sets_cell_as_active_cell_of_worksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.ActiveCell).IsNull();

        ws.Cell("B2").Active = true;
        await Assert.That(ws.ActiveCell).IsEqualTo(ws.Cell("B2"));
    }

    [Test]
    [Arguments("PY(4)", "_xlfn._xlws.PY(4)")]
    [Arguments("5 + py(abs(4) )", "5 + _xlfn._xlws.PY(abs(4) )")]
    [Arguments("COT(COTH(A5 + 2 * SIN(B7)))", "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))")]
    [Arguments("_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))", "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))")]
    public async Task FormulaA1_adds_prefix_to_future_functions(string formula, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");
        cell.FormulaA1 = formula;

        await Assert.That(cell.FormulaA1).IsEqualTo(expected);
    }

    [Test]
    [Arguments("PY(4)", "_xlfn._xlws.PY(4)")]
    [Arguments("5 + py(abs(4) )", "5 + _xlfn._xlws.PY(abs(4) )")]
    [Arguments("COT(COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))", "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))")]
    [Arguments("_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))", "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))")]
    public async Task FormulaR1C1_adds_prefix_to_future_functions(string formula, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");
        cell.FormulaR1C1 = formula;

        await Assert.That(cell.FormulaR1C1).IsEqualTo(expected);
    }

    [Test]
    public async Task FormulaA1_adds_prefix_to_all_future_functions()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");
        foreach (var (simpleName, prefixedName) in XLConstants.FutureFunctionMap.Value)
        {
            cell.FormulaA1 = simpleName + "()";
            await Assert.That(cell.FormulaA1).IsEqualTo(prefixedName + "()");
        }
    }

    /// <summary>
    /// Cell wrappers are vended from a small direct-mapped cache, so two requests for the same
    /// address can hand back the same instance. That is only sound while the wrapper carries no
    /// cell state of its own — everything must round-trip through the slices.
    /// </summary>
    [Test]
    public async Task Repeated_cell_access_sees_writes_made_through_an_earlier_handle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var first = ws.Cell(3, 4);
        first.Value = "hello";
        first.Style.Font.Bold = true;

        var second = ws.Cell(3, 4);
        await Assert.That(second.GetString()).IsEqualTo("hello");
        await Assert.That(second.Style.Font.Bold).IsTrue();

        // ...and a write through the second handle is visible through the first.
        second.Value = 42;
        await Assert.That(first.GetDouble()).IsEqualTo(42);
    }

    /// <summary>
    /// Distinct addresses must never collide in the wrapper cache, including addresses that share
    /// a cache slot. The cache is 16-wide and keyed off the packed point, so stepping the column
    /// by 16 lands repeatedly on one slot.
    /// </summary>
    [Test]
    public async Task Cells_that_share_a_wrapper_cache_slot_stay_independent()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        for (var i = 0; i < 8; i++)
            ws.Cell(1, 1 + i * 16).Value = i;

        for (var i = 0; i < 8; i++)
        {
            var cell = ws.Cell(1, 1 + i * 16);
            await Assert.That(cell.Address.RowNumber).IsEqualTo(1);
            await Assert.That(cell.Address.ColumnNumber).IsEqualTo(1 + i * 16);
            await Assert.That(cell.GetDouble()).IsEqualTo(i).Because($"column {1 + i * 16}");
        }
    }

    /// <summary>
    /// A handle held across an eviction must keep pointing at its own address.
    /// </summary>
    [Test]
    public async Task A_held_cell_handle_survives_wrapper_cache_eviction()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var held = ws.Cell(5, 5);
        held.Value = "kept";

        // Touch far more cells than the cache can hold.
        for (var c = 1; c <= 200; c++)
            ws.Cell(9, c).Value = c;

        await Assert.That(held.Address.RowNumber).IsEqualTo(5);
        await Assert.That(held.Address.ColumnNumber).IsEqualTo(5);
        await Assert.That(held.GetString()).IsEqualTo("kept");

        held.Value = "still mine";
        await Assert.That(ws.Cell(5, 5).GetString()).IsEqualTo("still mine");
    }
}
