using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Tests.Excel.Drawings.Shapes;

/// <summary>
/// Every shape the API can write, new or patched, passes the OpenXML SDK's schema validation (spec 15,
/// acceptance criterion 5). A save with <see cref="SaveOptions.ValidatePackage"/> throws on an error.
/// </summary>
public class ShapeValidationTests
{
    public static IEnumerable<XLShapeType> PresetTypes() =>
        Enum.GetValues<XLShapeType>().Where(type => type != XLShapeType.Custom);

    [Test]
    [MethodDataSource(nameof(PresetTypes))]
    public async Task A_new_shape_of_every_type_is_valid_plain_and_fully_formatted(XLShapeType type)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Shapes.Add(type, ws.Cell("A1"));
        FormatEverything(ws.Shapes.Add(type, 40, 50));

        await Assert.That(() => SaveValidated(wb)).ThrowsNothing();
    }

    [Test]
    [Arguments("moon")]
    [Arguments("heart")]
    [Arguments("flowChartMagneticDisk")]
    [Arguments("actionButtonHome")]
    public async Task A_new_shape_of_any_preset_is_valid(string preset)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        FormatEverything(ws.Shapes.AddPreset(preset, ws.Cell("B2")));

        await Assert.That(() => SaveValidated(wb)).ThrowsNothing();
    }

    [Test]
    public async Task Every_edit_of_a_loaded_shape_is_valid()
    {
        using var wb = TestHelper.LoadWorkbook(ShapeCreationTests.Fixture);
        foreach (var shape in wb.Worksheet(1).Shapes.ToList())
            FormatEverything(shape);

        ((IXLTextBox)wb.Worksheet(1).Shapes.Shape("Plain TextBox edited")).Text = "a\nb";
        wb.Worksheet(1).Shapes.Shape("Arc edited").NoLine = true;
        wb.Worksheet(1).Shapes.Shape("Left edited").NoFill = true;

        await Assert.That(() => SaveValidated(wb)).ThrowsNothing();
    }

    private static void FormatEverything(IXLShape shape)
    {
        shape.FillColor = XLColor.FromTheme(XLThemeColor.Accent2);
        shape.LineColor = XLColor.DarkRed;
        shape.LineWidthPt = 2.25;
        shape.LineDash = XLLineDashStyle.LongDashDot;
        shape.Name = shape.Name + " edited";
        shape.TextBody.VerticalAlignment = XLTextVerticalAlignment.Bottom;
        shape.TextBody.WrapText = false;
        var paragraph = shape.TextBody.AddParagraph();
        paragraph.Alignment = XLTextAlignment.Distributed;
        var run = paragraph.AddText("Every font property").SetBold().SetItalic().SetFontSize(20)
            .SetFontColor(XLColor.FromTheme(XLThemeColor.Accent5)).SetFontName("Consolas");
        run.Underline = XLShapeTextUnderline.Single;
        run.Strikethrough = false;
        shape.TextBody.AddParagraph();
        if (shape.TextBody.Count > 2)
            shape.TextBody[0].Alignment = XLTextAlignment.Justified;
    }

    private static void SaveValidated(XLWorkbook wb)
    {
        using var saved = new MemoryStream();
        wb.SaveAs(saved, new SaveOptions { ValidatePackage = true });
    }
}
