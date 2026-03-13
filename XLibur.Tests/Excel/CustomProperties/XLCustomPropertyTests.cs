using System;
using System.IO;
using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.CustomProperties;

[TestFixture]
public class XLCustomPropertyTests
{
    [Test]
    public void NumericString_IsPreservedAsText()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("OrderId", "12345");

        var prop = wb.CustomProperty("OrderId");

        Assert.That(prop.Type, Is.EqualTo(XLCustomPropertyType.Text));
        Assert.That(prop.GetValue<string>(), Is.EqualTo("12345"));
    }

    [Test]
    public void NumericString_SurvivesRoundTrip()
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
            Assert.That(prop.Type, Is.EqualTo(XLCustomPropertyType.Text));
            Assert.That(prop.GetValue<string>(), Is.EqualTo("12345"));
        }
    }

    [Test]
    public void Double_IsStoredAsNumber()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("Price", 99.99);

        var prop = wb.CustomProperty("Price");
        Assert.That(prop.Type, Is.EqualTo(XLCustomPropertyType.Number));
    }

    [Test]
    public void Integer_IsStoredAsNumber()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("Count", 42);

        var prop = wb.CustomProperty("Count");
        Assert.That(prop.Type, Is.EqualTo(XLCustomPropertyType.Number));
    }
}
