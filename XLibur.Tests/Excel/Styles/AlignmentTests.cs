using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Utils;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Styles;

public class AlignmentTests
{
    [Test]
    public async Task TextRotationCanBeFromMinus90To90DegreesAnd255ForVerticalLayout()
    {
        await TestHelper.CreateAndCompare(wb =>
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
    public async Task TextRotationIsConvertedOnLoadToMinus90To90Degrees()
    {
        await TestHelper.LoadAndAssert(async wb =>
        {
            var ws = wb.Worksheets.Single();
            await Assert.That(ws.Cell(1, 1).Style.Alignment.TextRotation).IsEqualTo(255);
            for (var column = 2; column < 21; ++column)
            {
                var expectedAngle = (column - 2) * 10 - 90;
                await Assert.That(ws.Cell(1, column).Style.Alignment.TextRotation).IsEqualTo(expectedAngle);
            }
        }, @"Other\Styles\Alignment\TextRotation.xlsx");
    }

    [Test]
    [Arguments(91)]
    [Arguments(-91)]
    [Arguments(254)]
    [Arguments(256)]
    public async Task TextRotationOutsideBoundsThrowsException(int textRotation)
    {
        await Assert.That(() =>
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.FirstCell().Style.Alignment.TextRotation = textRotation;
        }).Throws<ArgumentException>();
    }

    // Some third-party tools write spec-invalid upper-case alignment values
    // (e.g. horizontal="Center"). XLibur tolerates the casing rather than failing the load.
    [Test]
    [Arguments("center", XLAlignmentHorizontalValues.Center)]
    [Arguments("Center", XLAlignmentHorizontalValues.Center)]
    [Arguments("RIGHT", XLAlignmentHorizontalValues.Right)]
    public async Task HorizontalAlignmentToleratesInvalidCasing(string raw, XLAlignmentHorizontalValues expected)
    {
        var source = new EnumValue<HorizontalAlignmentValues> { InnerText = raw };
        await Assert.That(source.ToXLiburOrNull()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("center", XLAlignmentVerticalValues.Center)]
    [Arguments("Center", XLAlignmentVerticalValues.Center)]
    [Arguments("TOP", XLAlignmentVerticalValues.Top)]
    public async Task VerticalAlignmentToleratesInvalidCasing(string raw, XLAlignmentVerticalValues expected)
    {
        var source = new EnumValue<VerticalAlignmentValues> { InnerText = raw };
        await Assert.That(source.ToXLiburOrNull()).IsEqualTo(expected);
    }

    [Test]
    public async Task UnrecognizedAlignmentValueIsDiscarded()
    {
        var horizontal = new EnumValue<HorizontalAlignmentValues> { InnerText = "not-an-alignment" };
        var vertical = new EnumValue<VerticalAlignmentValues> { InnerText = "not-an-alignment" };
        await Assert.That(horizontal.ToXLiburOrNull()).IsNull();
        await Assert.That(vertical.ToXLiburOrNull()).IsNull();
    }

    [Test]
    public async Task AlignmentToXLiburKeepsDefaultWhenValueUnrecognized()
    {
        var defaultKey = XLAlignmentValue.Default.Key;
        var alignment = new Alignment
        {
            Horizontal = new EnumValue<HorizontalAlignmentValues> { InnerText = "Center" },
            Vertical = new EnumValue<VerticalAlignmentValues> { InnerText = "bogus" },
        };

        var result = OpenXmlHelper.AlignmentToXLibur(alignment, defaultKey);

        // Bad casing is recovered; truly unknown values fall back to the default.
        await Assert.That(result.Horizontal).IsEqualTo(XLAlignmentHorizontalValues.Center);
        await Assert.That(result.Vertical).IsEqualTo(defaultKey.Vertical);
    }
}
