using System.IO;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CustomProperties;

public class XLCustomPropertyTests
{
    [Test]
    public async Task NumericString_IsPreservedAsText()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("OrderId", "12345");

        var prop = wb.CustomProperty("OrderId");

        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Text);
        await Assert.That(prop.GetValue<string>()).IsEqualTo("12345");
    }

    [Test]
    public async Task NumericString_SurvivesRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            wb.CustomProperties.Add("OrderId", "12345");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var prop = wb.CustomProperty("OrderId");
            await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Text);
            await Assert.That(prop.GetValue<string>()).IsEqualTo("12345");
        }
    }

    [Test]
    public async Task Double_IsStoredAsNumber()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("Price", 99.99);

        var prop = wb.CustomProperty("Price");
        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Number);
    }

    [Test]
    public async Task Integer_IsStoredAsNumber()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("Count", 42);

        var prop = wb.CustomProperty("Count");
        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Number);
    }
}
