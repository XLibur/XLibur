using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Excel.Streaming;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// A save writes each number, date and time so that a load reads back the same <c>double</c>
/// (#741). The writer used <c>"G15"</c>, which drops the 16th and 17th significant digits a double
/// may need, and wrote a date by way of a <see cref="DateTime"/>, which moved the serial to the
/// nearest millisecond.
/// </summary>
public class CellNumberRoundTripTests
{
    /// <summary>
    /// Values from the files in the issue, and the extremes of the double range. Most need more than
    /// 15 significant digits; <c>double.MaxValue</c> in 15 digits rounds up past the range and no
    /// longer loads.
    /// </summary>
    public static double[] Numbers() =>
    [
        27.999999999999996, // =2.8/0.1
        1.3333333333333333, // =20/15
        -3.141592653589793, // =-PI()
        387.4800000000001,
        0.8333333333333335,
        0.1,
        1E-300,
        double.Epsilon,
        double.MaxValue,
        -double.MaxValue,
        123456789012345680,
    ];

    [Test]
    [MethodDataSource(nameof(Numbers))]
    public async Task A_number_round_trips(double number)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("S").Cell("A1").Value = number;
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet(1).Cell("A1").Value.GetNumber()).IsEqualTo(number);
    }

    [Test]
    [MethodDataSource(nameof(Numbers))]
    public async Task A_formula_cached_value_round_trips(double number)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = number;
            ws.Cell("B1").FormulaA1 = "A1";
            wb.RecalculateAllFormulas();
            wb.SaveAs(ms);
        }

        using var reloaded = new XLWorkbook(ms);
        await Assert.That(reloaded.Worksheet(1).Cell("B1").CachedValue.GetNumber()).IsEqualTo(number);
    }

    [Test]
    public async Task A_quotient_is_written_with_the_digits_it_needs()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("S").Cell("A1").FormulaA1 = "2.8/0.1";
            wb.RecalculateAllFormulas();
            wb.SaveAs(ms);
        }

        await Assert.That(Values(ms)).IsEquivalentTo(["27.999999999999996"]);
    }

    [Test]
    [Arguments(1d, "1")]
    [Arguments(-7d, "-7")]
    [Arguments(0.1, "0.1")]
    [Arguments(1234.5, "1234.5")]
    [Arguments(0.00001, "1E-05")]
    [Arguments(1E+20, "1E+20")]
    [Arguments(123456789012345d, "123456789012345")]
    public async Task A_common_number_is_spelled_as_before(double number, string text)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("S").Cell("A1").Value = number;
            wb.SaveAs(ms);
        }

        await Assert.That(Values(ms)).IsEquivalentTo([text]);
    }

    [Test]
    public async Task The_streaming_writer_writes_the_digits_a_number_needs()
    {
        using var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            wb.AddWorksheet("S").AppendRow(27.999999999999996, -3.141592653589793);
            wb.Finish();
        }

        await Assert.That(Values(ms)).IsEquivalentTo(["27.999999999999996", "-3.141592653589793"]);
    }

    /// <summary>
    /// A pivot cache record holds the number of its source cell, written as the cell is.
    /// </summary>
    [Test]
    public async Task A_pivot_cache_record_writes_the_digits_a_number_needs()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = "Name";
            data.Cell("B1").Value = "Amount";
            data.Cell("A2").Value = "a";
            data.Cell("B2").Value = 27.999999999999996;
            var pivots = wb.AddWorksheet("Pivots");
            var pt = pivots.PivotTables.Add("pvt", pivots.Cell("A1"), data.Range("A1:B2"));
            pt.RowLabels.Add("Name");
            pt.Values.Add("Amount");
            wb.SaveAs(ms);
        }

        await Assert.That(ms.ReadPart(ms.PartNames().Single(n => n.Contains("pivotCacheRecords", StringComparison.Ordinal)))).Contains("v=\"27.999999999999996\"");
    }

    /// <summary>
    /// A number formatted as a date or time loads as a <see cref="XLDataType.DateTime"/> or
    /// <see cref="XLDataType.TimeSpan"/>. Its serial must be written back as the file had it: neither
    /// cut to 15 digits (40982.56313888889) nor moved to the nearest millisecond (42536.5128889537).
    /// </summary>
    [Test]
    [Arguments("40982.56313888889", "yyyy-mm-dd hh:mm:ss", XLDataType.DateTime, false)]
    [Arguments("42536.5128889537", "yyyy-mm-dd hh:mm:ss", XLDataType.DateTime, false)]
    [Arguments("42536.5128889537", "yyyy-mm-dd hh:mm:ss", XLDataType.DateTime, true)]
    [Arguments("30.75", "yyyy-mm-dd hh:mm:ss", XLDataType.DateTime, true)]
    [Arguments("0.8333333333333335", "[h]:mm:ss", XLDataType.TimeSpan, false)]
    [Arguments("1.0000000000000002", "[h]:mm:ss", XLDataType.TimeSpan, false)]
    public async Task A_loaded_date_or_time_serial_is_written_back_unchanged(string serial, string format,
        XLDataType dataType, bool use1904DateSystem)
    {
        var number = double.Parse(serial, CultureInfo.InvariantCulture);
        using var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Use1904DateSystem = use1904DateSystem;
            var cell = wb.AddWorksheet("S").Cell("A1");
            cell.Value = number;
            cell.Style.NumberFormat.Format = format;
            wb.SaveAs(source);
        }

        await Assert.That(Values(source)).IsEquivalentTo([serial]);

        using var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            await Assert.That(wb.Worksheet(1).Cell("A1").DataType).IsEqualTo(dataType);
            wb.SaveAs(resaved, new SaveOptions { RewriteUnchangedSheets = true });
        }

        await Assert.That(Values(resaved)).IsEquivalentTo([serial]);
    }

    /// <summary>
    /// In the 1904 date system serial 0 is 1 January 1904. The 1900 system's phantom 29 February
    /// 1900 has no part in it, so the first two months of 1904 load and save on the same day as the
    /// rest.
    /// </summary>
    [Test]
    [Arguments(0d, 1904, 1, 1)]
    [Arguments(30d, 1904, 1, 31)]
    [Arguments(59d, 1904, 2, 29)]
    [Arguments(60d, 1904, 3, 1)]
    [Arguments(5000d, 1917, 9, 9)]
    public async Task A_1904_serial_is_the_day_Excel_shows(double serial, int year, int month, int day)
    {
        var date = new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Unspecified);
        using var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Use1904DateSystem = true;
            var cell = wb.AddWorksheet("S").Cell("A1");
            cell.Value = serial;
            cell.Style.NumberFormat.Format = "yyyy-mm-dd";
            wb.SaveAs(source);
        }

        using (var loaded = new XLWorkbook(source))
            await Assert.That(loaded.Worksheet(1).Cell("A1").GetDateTime()).IsEqualTo(date);

        using var written = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Use1904DateSystem = true;
            wb.AddWorksheet("S").Cell("A1").Value = date;
            wb.SaveAs(written);
        }

        await Assert.That(Values(written)).IsEquivalentTo([serial.ToString(CultureInfo.InvariantCulture)]);
    }

    /// <summary>
    /// A 1904 serial is a date only when the 1900 serial it moves to is one. Past the end of the
    /// range it stays a number, and is written back as the file had it.
    /// </summary>
    [Test]
    public async Task A_1904_serial_past_the_date_range_once_moved_stays_a_number()
    {
        using var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Use1904DateSystem = true;
            var cell = wb.AddWorksheet("S").Cell("A1");
            cell.Value = 2958000;
            cell.Style.NumberFormat.Format = "yyyy-mm-dd";
            wb.SaveAs(source);
        }

        using var resaved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            var cell = wb.Worksheet(1).Cell("A1");
            await Assert.That(cell.DataType).IsEqualTo(XLDataType.Number);
            await Assert.That(cell.Value.GetNumber()).IsEqualTo(2958000);
            wb.SaveAs(resaved, new SaveOptions { RewriteUnchangedSheets = true });
        }

        await Assert.That(Values(resaved)).IsEquivalentTo(["2958000"]);
    }

    /// <summary>
    /// The conversions between the model's 1900 serial and a 1904 file serial undo each other, also
    /// for dates before 1 March 1900, where the 1900 serial is one less than the OLE date.
    /// </summary>
    [Test]
    [Arguments(1d)]
    [Arguments(59d)]
    [Arguments(59.5)]
    [Arguments(61d)]
    [Arguments(61.25)]
    [Arguments(1462d)]
    [Arguments(44000.75)]
    public async Task The_1904_serial_conversions_undo_each_other(double serial)
    {
        var fileSerial = CellXmlWriter.ToFileSerialDateTime(serial, use1904DateSystem: true);
        await Assert.That(CellXmlWriter.FromFileSerialDateTime(fileSerial, use1904DateSystem: true)).IsEqualTo(serial);
        await Assert.That(CellXmlWriter.ToFileSerialDateTime(serial, use1904DateSystem: false)).IsEqualTo(serial);
    }

    /// <summary>
    /// The load reads the nearest double to the text, as <see cref="double.Parse(string)"/> does,
    /// for the 17 significant digits Excel writes. Its fast parser rounded a mantissa above 2^53
    /// twice and could land one bit off, so a load and save changed the value.
    /// </summary>
    [Test]
    [Arguments("9420.1999999999989")]
    [Arguments("9420.199999999999")]
    [Arguments("0.90878878027413209")]
    [Arguments("0.9463524662607499")]
    [Arguments("-0.95182005256239389")]
    [Arguments("35156.490000000005")]
    [Arguments("12345678901234567")]
    [Arguments("0.1")]
    public async Task A_loaded_number_is_the_nearest_double_to_its_text(string text)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("S").Cell("A1").Value = 1;
            wb.SaveAs(ms);
        }

        ms.RewriteSheet1(xml => Regex.Replace(xml, "(<(?:[A-Za-z_][\\w.-]*:)?v>)1(</)", "${1}" + text + "${2}"));

        using var loaded = new XLWorkbook(ms);
        await Assert.That(loaded.Worksheet(1).Cell("A1").Value.GetNumber())
            .IsEqualTo(double.Parse(text, CultureInfo.InvariantCulture));
    }

    /// <summary>The text of every <c>&lt;v&gt;</c> in the first sheet, prefixed or not.</summary>
    private static string[] Values(Stream package) =>
        Regex.Matches(package.Sheet1Xml(), "<(?:[A-Za-z_][\\w.-]*:)?v>([^<]*)</")
            .Select(m => m.Groups[1].Value)
            .ToArray();
}
