using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Utils;
using NUnit.Framework;

namespace XLibur.Tests.Excel.Styles;

[TestFixture]
public class AlignmentTests
{
    [Test]
    public void TextRotationCanBeFromMinus90To90DegreesAnd255ForVerticalLayout()
    {
        TestHelper.CreateAndCompare(wb =>
        {
            var ws = wb.AddWorksheet();
            ws.ColumnWidth = 10;
            ws.Cell(1, 1)
                .SetValue("Vertical: 255")
                .Style.Alignment.SetTextRotation(255);

            for (var angle = -90; angle <= +90; angle += 10)
            {
                var column = (angle + 90) / 10 + 2;
                var cell = ws.Cell(1, column);
                cell.Value = $"Rotation: {angle}";
                cell.Style.Alignment.TextRotation = angle;
            }
        }, @"Other\Styles\Alignment\TextRotation.xlsx");
    }

    [Test]
    public void TextRotationIsConvertedOnLoadToMinus90To90Degrees()
    {
        TestHelper.LoadAndAssert(wb =>
        {
            var ws = wb.Worksheets.Single();
            Assert.AreEqual(255, ws.Cell(1, 1).Style.Alignment.TextRotation);
            for (var column = 2; column < 21; ++column)
            {
                var expectedAngle = (column - 2) * 10 - 90;
                Assert.AreEqual(expectedAngle, ws.Cell(1, column).Style.Alignment.TextRotation);
            }
        }, @"Other\Styles\Alignment\TextRotation.xlsx");
    }

    [TestCase(91)]
    [TestCase(-91)]
    [TestCase(254)]
    [TestCase(256)]
    public void TextRotationOutsideBoundsThrowsException(int textRotation)
    {
        Assert.Throws<ArgumentException>(() =>
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().Style.Alignment.TextRotation = textRotation;
        });
    }

    // Some third-party tools write spec-invalid upper-case alignment values
    // (e.g. horizontal="Center"). XLibur tolerates the casing rather than failing the load.
    [TestCase("center", XLAlignmentHorizontalValues.Center)]
    [TestCase("Center", XLAlignmentHorizontalValues.Center)]
    [TestCase("RIGHT", XLAlignmentHorizontalValues.Right)]
    public void HorizontalAlignmentToleratesInvalidCasing(string raw, XLAlignmentHorizontalValues expected)
    {
        var source = new EnumValue<HorizontalAlignmentValues> { InnerText = raw };
        Assert.AreEqual(expected, source.ToXLiburOrNull());
    }

    [TestCase("center", XLAlignmentVerticalValues.Center)]
    [TestCase("Center", XLAlignmentVerticalValues.Center)]
    [TestCase("TOP", XLAlignmentVerticalValues.Top)]
    public void VerticalAlignmentToleratesInvalidCasing(string raw, XLAlignmentVerticalValues expected)
    {
        var source = new EnumValue<VerticalAlignmentValues> { InnerText = raw };
        Assert.AreEqual(expected, source.ToXLiburOrNull());
    }

    [Test]
    public void UnrecognizedAlignmentValueIsDiscarded()
    {
        var horizontal = new EnumValue<HorizontalAlignmentValues> { InnerText = "not-an-alignment" };
        var vertical = new EnumValue<VerticalAlignmentValues> { InnerText = "not-an-alignment" };
        Assert.IsNull(horizontal.ToXLiburOrNull());
        Assert.IsNull(vertical.ToXLiburOrNull());
    }

    [Test]
    public void AlignmentToXLiburKeepsDefaultWhenValueUnrecognized()
    {
        var defaultKey = XLAlignmentValue.Default.Key;
        var alignment = new Alignment
        {
            Horizontal = new EnumValue<HorizontalAlignmentValues> { InnerText = "Center" },
            Vertical = new EnumValue<VerticalAlignmentValues> { InnerText = "bogus" },
        };

        var result = OpenXmlHelper.AlignmentToXLibur(alignment, defaultKey);

        // Bad casing is recovered; truly unknown values fall back to the default.
        Assert.AreEqual(XLAlignmentHorizontalValues.Center, result.Horizontal);
        Assert.AreEqual(defaultKey.Vertical, result.Vertical);
    }
}
