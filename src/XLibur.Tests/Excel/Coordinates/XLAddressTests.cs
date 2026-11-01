using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests;

public class XLAddressTests
{
    [Test]
    public async Task ToStringTest()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var address = ws.Cell(1, 1).Address;

        await Assert.That(address.ToString()).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("R1C1");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!A1");

        await Assert.That(address.ToStringRelative()).IsEqualTo("A1");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("Sheet1!A1");

        await Assert.That(address.ToStringFixed()).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("Sheet1!$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("Sheet1!R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$A$1");
    }

    [Test]
    public async Task ToStringTestWithSpace()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        var address = ws.Cell(1, 1).Address;

        await Assert.That(address.ToString()).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("R1C1");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!A1");

        await Assert.That(address.ToStringRelative()).IsEqualTo("A1");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("'Sheet 1'!A1");

        await Assert.That(address.ToStringFixed()).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("'Sheet 1'!R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!$A$1");
    }

    [Test]
    public async Task InvalidAddressToStringTest()
    {
        var address = ProduceInvalidAddress();

        await Assert.That(address.ToString()).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!#REF!");
    }

    [Test]
    public async Task InvalidAddressToStringFixedTest()
    {
        var address = ProduceInvalidAddress();

        await Assert.That(address.ToStringFixed()).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("'Sheet 1'!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("'Sheet 1'!#REF!");
    }

    [Test]
    public async Task InvalidAddressToStringRelativeTest()
    {
        var address = ProduceInvalidAddress();

        await Assert.That(address.ToStringRelative()).IsEqualTo("#REF!");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("'Sheet 1'!#REF!");
    }

    [Test]
    public async Task AddressOnDeletedWorksheetToStringTest()
    {
        var address = ProduceAddressOnDeletedWorksheet();

        await Assert.That(address.ToString()).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("R1C1");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("A1");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("#REF!A1");
    }

    [Test]
    public async Task AddressOnDeletedWorksheetToStringFixedTest()
    {
        var address = ProduceAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringFixed()).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("#REF!$A$1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("#REF!R1C1");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("#REF!$A$1");
    }

    [Test]
    public async Task AddressOnDeletedWorksheetToStringRelativeTest()
    {
        var address = ProduceAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringRelative()).IsEqualTo("A1");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("#REF!A1");
    }

    [Test]
    public async Task InvalidAddressOnDeletedWorksheetToStringTest()
    {
        var address = await ProduceInvalidAddressOnDeletedWorksheet();

        await Assert.That(address.ToString()).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.A1)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.R1C1)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default)).IsEqualTo("#REF!");
        await Assert.That(address.ToString(XLReferenceStyle.Default, true)).IsEqualTo("#REF!#REF!");
    }

    [Test]
    public async Task InvalidAddressOnDeletedWorksheetToStringFixedTest()
    {
        var address = await ProduceInvalidAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringFixed()).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default)).IsEqualTo("#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.A1, true)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.R1C1, true)).IsEqualTo("#REF!#REF!");
        await Assert.That(address.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("#REF!#REF!");
    }

    [Test]
    public async Task InvalidAddressOnDeletedWorksheetToStringRelativeTest()
    {
        var address = await ProduceInvalidAddressOnDeletedWorksheet();

        await Assert.That(address.ToStringRelative()).IsEqualTo("#REF!");
        await Assert.That(address.ToStringRelative(true)).IsEqualTo("#REF!#REF!");
    }

    #region Private Methods

    private static IXLAddress ProduceInvalidAddress()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        var range = ws.Range("A1:B2");

        ws.Rows(1, 5).Delete();
        return range.RangeAddress.FirstAddress;
    }

    private static IXLAddress ProduceAddressOnDeletedWorksheet()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        var address = ws.Cell("A1").Address;

        ws.Delete();
        return address;
    }

    private static async Task<IXLAddress> ProduceInvalidAddressOnDeletedWorksheet()
    {
        var address = ProduceInvalidAddress();
        await Assert.That(address.Worksheet).IsNotNull();
        address.Worksheet.Delete();
        return address;
    }

    #endregion Private Methods
}
