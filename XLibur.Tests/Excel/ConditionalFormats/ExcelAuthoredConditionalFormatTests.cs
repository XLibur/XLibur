using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// Conditional formats as Excel writes them, which XLibur's writer cannot produce: it always emits
/// the optional children the reader once assumed. Each input is made by saving a rule through
/// XLibur and then hand-editing the part to what Excel authors (spec 48): an Excel-authored
/// fixture would be binary, and these edits are the whole difference.
/// </summary>
public class ExcelAuthoredConditionalFormatTests
{
    private static MemoryStream Save(XLWorkbook wb)
    {
        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    private static MemoryStream SaveDataBar(XLColor positive, XLColor negative)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = -3;
        ws.Cell("A2").Value = 5;
        ws.Range("A1:A2").AddConditionalFormat().DataBar(positive, negative).LowestValue().HighestValue();
        return Save(wb);
    }

    private static string Edit(string xml, string pattern, string replacement)
    {
        var edited = Regex.Replace(xml, pattern, replacement);
        if (edited == xml)
            throw new InvalidOperationException($"The edit '{pattern}' matched nothing in: {xml}");
        return edited;
    }

    private static MemoryStream SaveAgain(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        return Save(wb);
    }

    #region Data bar without a negative fill (load crash)

    /// <summary>
    /// Excel leaves out <c>x14:negativeFillColor</c> when the negative bar is the colour of the
    /// positive one, and says so with <c>negativeBarColorSameAsPositive</c>.
    /// </summary>
    private static MemoryStream DataBarWithoutNegativeFill() =>
        SaveDataBar(XLColor.Red, XLColor.Blue)
            .RewriteSheet1(xml =>
            {
                xml = Edit(xml, @"<x14:negativeFillColor[^>]*/>", "");
                return Edit(xml, "<x14:dataBar ", "<x14:dataBar negativeBarColorSameAsPositive=\"1\" ");
            });

    [Test]
    public async Task DataBarWithoutNegativeFillLoads()
    {
        using var package = DataBarWithoutNegativeFill();

        using var wb = new XLWorkbook(package);
        var cf = wb.Worksheet("S").ConditionalFormats.Single();

        await Assert.That(cf.ConditionalFormatType).IsEqualTo(XLConditionalFormatType.DataBar);
        await Assert.That(cf.Colors.Count).IsEqualTo(1);
        await Assert.That(cf.Colors[1]).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task DataBarWithoutNegativeFillKeepsTheFlagAndStaysWithoutAFillAfterASave()
    {
        using var package = DataBarWithoutNegativeFill();
        using var saved = SaveAgain(package);

        var sheet = saved.Sheet1Xml();
        await Assert.That(sheet).Contains("negativeBarColorSameAsPositive=\"1\"");
        await Assert.That(sheet).DoesNotContain("negativeFillColor");

        using var wb = new XLWorkbook(saved);
        await Assert.That(wb.Worksheet("S").ConditionalFormats.Single().Colors.Count).IsEqualTo(1);
    }

    #endregion

    #region Icon set threshold without a value (save throws)

    private static MemoryStream IconSetWithoutAThresholdValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        ws.Cell("A1").Value = 1;
        ws.Range("A1:A1").AddConditionalFormat().IconSet(XLIconSetStyle.ThreeArrows)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 0, XLCFContentType.Percent)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 33, XLCFContentType.Percent)
            .AddValue(XLCFIconSetOperator.EqualOrGreaterThan, 67, XLCFContentType.Percent);
        // cfvo val is optional: Excel omits it for the first threshold of an icon set.
        return Save(wb).RewriteSheet1(xml => Edit(xml, "(<x:cfvo[^>]*?) val=\"0\"", "$1"));
    }

    [Test]
    public async Task IconSetThresholdWithoutAValueLoadsAndSaves()
    {
        using var package = IconSetWithoutAThresholdValue();

        using var wb = new XLWorkbook(package);
        var cf = wb.Worksheet("S").ConditionalFormats.Single();
        await Assert.That(cf.ConditionalFormatType).IsEqualTo(XLConditionalFormatType.IconSet);

        using var saved = Save(wb);
        using var reloaded = new XLWorkbook(saved);
        var again = reloaded.Worksheet("S").ConditionalFormats.Single();
        await Assert.That(again.ConditionalFormatType).IsEqualTo(XLConditionalFormatType.IconSet);
        await Assert.That(again.IconSetStyle).IsEqualTo(XLIconSetStyle.ThreeArrows);
        await Assert.That(again.ContentTypes.Count).IsEqualTo(3);
    }

    #endregion

    #region Data bar attributes (lost on save)

    private static MemoryStream BorderedDataBar() =>
        SaveDataBar(XLColor.Red, XLColor.Blue)
            .RewriteSheet1(xml =>
            {
                xml = Edit(xml, "<x14:dataBar [^>]*?minLength=\"0\" maxLength=\"100\"",
                    "<x14:dataBar minLength=\"25\" maxLength=\"80\" border=\"1\" " +
                    "negativeBarBorderColorSameAsPositive=\"0\"");
                return Edit(xml, "<x14:negativeFillColor",
                    "<x14:borderColor rgb=\"FF00FF00\"/><x14:negativeFillColor");
            })
            .RewriteSheet1(xml => Edit(xml, "<x14:axisColor",
                "<x14:negativeBorderColor rgb=\"FFFFFF00\"/><x14:axisColor"));

    [Test]
    public async Task DataBarBorderLengthsAndFlagsSurviveASave()
    {
        using var package = BorderedDataBar();
        using var saved = SaveAgain(package);

        var sheet = saved.Sheet1Xml();
        await Assert.That(sheet).Contains("minLength=\"25\"");
        await Assert.That(sheet).Contains("maxLength=\"80\"");
        await Assert.That(sheet).Contains("border=\"1\"");
        await Assert.That(sheet).Contains("negativeBarBorderColorSameAsPositive=\"0\"");
        await Assert.That(sheet).Contains("<x14:borderColor rgb=\"FF00FF00\"");
        await Assert.That(sheet).Contains("<x14:negativeBorderColor rgb=\"FFFFFF00\"");
        await Assert.That(sheet).Contains("<x14:negativeFillColor rgb=\"FF0000FF\"");
    }

    [Test]
    public async Task DataBarWrittenByXLiburKeepsItsFullLengthAndNoBorder()
    {
        using var package = SaveDataBar(XLColor.Red, XLColor.Blue);

        var sheet = package.Sheet1Xml();
        await Assert.That(sheet).Contains("minLength=\"0\"");
        await Assert.That(sheet).Contains("maxLength=\"100\"");
        await Assert.That(sheet).DoesNotContain("border=");
        await Assert.That(sheet).DoesNotContain("borderColor");
        await Assert.That(sheet).DoesNotContain("negativeBarColorSameAsPositive");
    }

    [Test]
    public async Task DataBarWithoutLengthsTakesTheSchemaDefaults()
    {
        using var package = SaveDataBar(XLColor.Red, XLColor.Blue)
            .RewriteSheet1(xml => Edit(xml, " minLength=\"0\" maxLength=\"100\"", ""));
        using var saved = SaveAgain(package);

        var sheet = saved.Sheet1Xml();
        await Assert.That(sheet).Contains("minLength=\"10\"");
        await Assert.That(sheet).Contains("maxLength=\"90\"");
    }

    #endregion
}
