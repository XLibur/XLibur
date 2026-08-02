using System;
using XLibur.Excel;
using XLibur.Excel.Coordinates;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class XLRangeAddressTests
{
    [Test]
    public async Task ToStringTest()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var address = ws.Cell(1, 1).AsRange().RangeAddress;

        await Assert.That(address.ToString()).IsEqualTo("A1:A1");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1, true)).IsEqualTo("Sheet1!R1C1:R1C1");

        await Assert.That(address.ToStringRelative()).IsEqualTo("A1:A1");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("Sheet1!A1:A1");

        await Assert.That(address.ToStringFixed()).IsEqualTo("$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("R1C1:R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("Sheet1!$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("Sheet1!R1C1:R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$A$1:$A$1");
    }

    [Test]
    public async Task ToStringTestWithSpace()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        var address = ws.Cell(1, 1).AsRange().RangeAddress;

        await Assert.That(address.ToString()).IsEqualTo("A1:A1");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1, true)).IsEqualTo("'Sheet 1'!R1C1:R1C1");

        await Assert.That(address.ToStringRelative()).IsEqualTo("A1:A1");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("'Sheet 1'!A1:A1");

        await Assert.That(address.ToStringFixed()).IsEqualTo("$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("R1C1:R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!$A$1:$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("'Sheet 1'!R1C1:R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!$A$1:$A$1");
    }

    [Test]
    [Arguments("B2:E5", "B2:E5")]
    [Arguments("E5:B2", "B2:E5")]
    [Arguments("B5:E2", "B2:E5")]
    [Arguments("B2:E$5", "B2:E$5")]
    [Arguments("B2:$E$5", "B2:$E$5")]
    [Arguments("B$2:$E$5", "B$2:$E$5")]
    [Arguments("$B$2:$E$5", "$B$2:$E$5")]
    [Arguments("B5:E$2", "B$2:E5")]
    [Arguments("$B$5:E2", "$B2:E$5")]
    [Arguments("$B$5:E$2", "$B$2:E$5")]
    [Arguments("$B$5:$E$2", "$B$2:$E$5")]
    public async Task RangeAddressNormalizeTest(string inputAddress, string expectedAddress)
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1") as XLWorksheet;
        var rangeAddress = new XLRangeAddress(ws, inputAddress);

        var normalizedAddress = rangeAddress.Normalize();

        await Assert.That(rangeAddress.Worksheet).IsSameReferenceAs(ws);
        await Assert.That(normalizedAddress.ToString()).IsEqualTo(expectedAddress);
    }

    [Test]
    public async Task InvalidRangeAddressToStringTest()
    {
        var address = ProduceInvalidAddress();

        await Assert.That(address.ToString()).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1, true)).IsEqualTo("'Sheet 1'!#REF!");
    }

    [Test]
    public async Task InvalidRangeAddressToStringFixedTest()
    {
        var address = ProduceInvalidAddress();

        await Assert.That(address.ToStringFixed()).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("'Sheet 1'!#REF!");
    }

    [Test]
    public async Task InvalidRangeAddressToStringRelativeTest()
    {
        var address = ProduceInvalidAddress();

        await Assert.That(address.ToStringRelative()).IsEqualTo("#REF!");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("'Sheet 1'!#REF!");
    }

    [Test]
    public async Task RangeAddressOnDeletedWorksheetToStringTest()
    {
        var address = ProduceAddressOnDeletedWorksheet();

        await Assert.That(address.ToString()).IsEqualTo("#REF!A1:B2");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("#REF!A1:B2");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("#REF!A1:B2");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("#REF!R1C1:R2C2");
        await Assert.That(address.ToString(XLReferenceStyle.A1, true)).IsEqualTo("#REF!A1:B2");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("#REF!A1:B2");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1, true)).IsEqualTo("#REF!R1C1:R2C2");
    }

    [Test]
    public async Task RangeAddressOnDeletedWorksheetToStringFixedTest()
    {
        var address = ProduceAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringFixed()).IsEqualTo("#REF!$A$1:$B$2");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("#REF!$A$1:$B$2");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("#REF!$A$1:$B$2");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("#REF!R1C1:R2C2");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("#REF!$A$1:$B$2");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("#REF!$A$1:$B$2");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("#REF!R1C1:R2C2");
    }

    [Test]
    public async Task RangeAddressOnDeletedWorksheetToStringRelativeTest()
    {
        var address = ProduceAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringRelative()).IsEqualTo("#REF!A1:B2");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("#REF!A1:B2");
    }

    [Test]
    public async Task InvalidRangeAddressOnDeletedWorksheetToStringTest()
    {
        var address = ProduceInvalidAddressOnDeletedWorksheet();

        await Assert.That(address.ToString()).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.A1, true)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1, true)).IsEqualTo("#REF!#REF!");
    }

    [Test]
    public async Task InvalidRangeAddressOnDeletedWorksheetToStringFixedTest()
    {
        var address = ProduceInvalidAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringFixed()).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("#REF!#REF!");
    }

    [Test]
    public async Task InvalidRangeAddressOnDeletedWorksheetToStringRelativeTest()
    {
        var address = ProduceInvalidAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringRelative()).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("#REF!#REF!");
    }

    [Test]
    public async Task FullSpanAddressCannotChange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var wsRange = ws.AsRange();
        var row = ws.FirstRow().RowBelow(4).AsRange();
        var column = ws.FirstColumn().ColumnRight(4).AsRange();

        await Assert.That(wsRange.RangeAddress.ToString()).IsEqualTo($"1:{XLHelper.MaxRowNumber}");
        await Assert.That(row.RangeAddress.ToString()).IsEqualTo("5:5");
        await Assert.That(column.RangeAddress.ToString()).IsEqualTo("E:E");

        ws.Columns("Y:Z").Delete();
        ws.Rows("9:10").Delete();

        await Assert.That(wsRange.RangeAddress.ToString()).IsEqualTo($"1:{XLHelper.MaxRowNumber}");
        await Assert.That(row.RangeAddress.ToString()).IsEqualTo("5:5");
        await Assert.That(column.RangeAddress.ToString()).IsEqualTo("E:E");
    }

    [Test]
    public async Task RangeAddressIsNormalized()
    {
        var ws = new XLWorkbook().AddWorksheet();

        XLRangeAddress rangeAddress;

        rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("A1"), ws.Cell("C3")).RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsTrue();

        rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("C3"), ws.Cell("A1")).RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsFalse();

        rangeAddress = (XLRangeAddress)ws.Range("B2:B1").RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsFalse();

        rangeAddress = (XLRangeAddress)ws.Range("B2:B10").RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsTrue();

        rangeAddress = (XLRangeAddress)ws.Range("B:B").RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsTrue();

        rangeAddress = (XLRangeAddress)ws.Range("2:2").RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsTrue();

        rangeAddress = (XLRangeAddress)ws.RangeAddress;
        await Assert.That(rangeAddress.IsNormalized).IsTrue();
    }

    [Test]
    public async Task AsRangeTests()
    {
        XLRangeAddress rangeAddress;
        rangeAddress = new XLRangeAddress
        (
            new XLAddress(1, 1, false, false),
            new XLAddress(5, 5, false, false)
        );

        await Assert.That(rangeAddress.IsValid).IsTrue();
        await Assert.That(rangeAddress.IsNormalized).IsTrue();
        await Assert.That(() => rangeAddress.AsRange()).Throws<InvalidOperationException>();

        var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;
        rangeAddress = new XLRangeAddress
        (
            new XLAddress(ws, 1, 1, false, false),
            new XLAddress(ws, 5, 5, false, false)
        );

        await Assert.That(rangeAddress.IsValid).IsTrue();
        await Assert.That(rangeAddress.IsNormalized).IsTrue();
        await Assert.That(() => rangeAddress.AsRange()).ThrowsNothing();
    }

    [Test]
    public async Task RelativeRanges()
    {
        var ws = new XLWorkbook().AddWorksheet();

        IXLRangeAddress rangeAddress;

        rangeAddress = ws.Range("D4:E4").RangeAddress.Relative(ws.Range("A1:E4").RangeAddress, ws.Range("B10:F14").RangeAddress);
        await Assert.That(rangeAddress.IsValid).IsTrue();
        await Assert.That(rangeAddress.ToString()).IsEqualTo("E13:F13");

        rangeAddress = ws.Range("D4:E4").RangeAddress.Relative(ws.Range("B10:F14").RangeAddress, ws.Range("A1:E4").RangeAddress);
        await Assert.That(rangeAddress.IsValid).IsFalse();
        await Assert.That(rangeAddress.ToString()).IsEqualTo("#REF!");

        rangeAddress = ws.Range("C3").RangeAddress.Relative(ws.Range("A1:B2").RangeAddress, ws.Range("C3").RangeAddress);
        await Assert.That(rangeAddress.IsValid).IsTrue();
        await Assert.That(rangeAddress.ToString()).IsEqualTo("E5:E5");

        rangeAddress = ws.Range("B2").RangeAddress.Relative(ws.Range("A1").RangeAddress, ws.Range("C3").RangeAddress);
        await Assert.That(rangeAddress.IsValid).IsTrue();
        await Assert.That(rangeAddress.ToString()).IsEqualTo("D4:D4");

        rangeAddress = ws.Range("A1").RangeAddress.Relative(ws.Range("B2").RangeAddress, ws.Range("A1").RangeAddress);
        await Assert.That(rangeAddress.IsValid).IsFalse();
        await Assert.That(rangeAddress.ToString()).IsEqualTo("#REF!");
    }

    [Test]
    public async Task TestSpanProperties()
    {
        var ws = new XLWorkbook().AddWorksheet() as XLWorksheet;

        var range = ws.Range("B3:E5")!;
        var rangeAddress = range.RangeAddress as IXLRangeAddress;
        await Assert.That(rangeAddress.ColumnSpan).IsEqualTo(4);
        await Assert.That(rangeAddress.RowSpan).IsEqualTo(3);
        await Assert.That(rangeAddress.NumberOfCells).IsEqualTo(12);

        range = ws.Range("E5:B3");
        rangeAddress = range.RangeAddress;
        await Assert.That(rangeAddress.ColumnSpan).IsEqualTo(4);
        await Assert.That(rangeAddress.RowSpan).IsEqualTo(3);
        await Assert.That(rangeAddress.NumberOfCells).IsEqualTo(12);

        rangeAddress = ProduceAddressOnDeletedWorksheet();
        await Assert.That(rangeAddress.ColumnSpan).IsEqualTo(2);
        await Assert.That(rangeAddress.RowSpan).IsEqualTo(2);
        await Assert.That(rangeAddress.NumberOfCells).IsEqualTo(4);

        rangeAddress = ProduceInvalidAddress();
        await Assert.That(() => { var x = rangeAddress.ColumnSpan; }).Throws<InvalidOperationException>();
        await Assert.That(() => { var x = rangeAddress.RowSpan; }).Throws<InvalidOperationException>();
        await Assert.That(() => { var x = rangeAddress.NumberOfCells; }).Throws<InvalidOperationException>();
    }

    #region Private Methods

    private static IXLRangeAddress ProduceInvalidAddress()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        var range = ws.Range("A1:B2");

        ws.Rows(1, 5).Delete();
        return range.RangeAddress;
    }

    private static IXLRangeAddress ProduceAddressOnDeletedWorksheet()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        var address = ws.Range("A1:B2").RangeAddress;

        ws.Delete();
        return address;
    }

    private static IXLRangeAddress ProduceInvalidAddressOnDeletedWorksheet()
    {
        var address = ProduceInvalidAddress();
        address.Worksheet!.Delete();
        return address;
    }

    #endregion Private Methods
}
