using System.Linq;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.AutoFilters;

[TestFixture]
public class ColorFilterTests
{
    [Test]
    public void Color_filter_shows_rows_matching_background_color()
    {
        new AutoFilterTester(f => f.ColorFilter(XLColor.Red))
            .Add("A", s => s.Fill.SetBackgroundColor(XLColor.Red), true)
            .Add("B", s => s.Fill.SetBackgroundColor(XLColor.Blue), false)
            .Add("C", s => s.Fill.SetBackgroundColor(XLColor.Red), true)
            .Add("D", false)
            .AssertVisibility();
    }

    [Test]
    public void Font_color_filter_shows_rows_matching_font_color()
    {
        new AutoFilterTester(f => f.FontColorFilter(XLColor.Green))
            .Add("A", s => s.Font.SetFontColor(XLColor.Green), true)
            .Add("B", s => s.Font.SetFontColor(XLColor.Red), false)
            .Add("C", s => s.Font.SetFontColor(XLColor.Green), true)
            .Add("D", false)
            .AssertVisibility();
    }

    [Test]
    public void Color_filter_replaces_previous_filter_type()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Data";
        ws.Cell("A2").Value = 1;
        ws.Cell("A2").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Cell("A3").Value = 2;
        ws.Cell("A3").Style.Fill.SetBackgroundColor(XLColor.Blue);

        var autoFilter = ws.Range("A1:A3").SetAutoFilter();

        // First, set a regular filter
        autoFilter.Column(1).AddFilter(1);
        Assert.That(autoFilter.Column(1).FilterType, Is.EqualTo(XLFilterType.Regular));

        // Then switch to the color filter — should replace
        autoFilter.Column(1).ColorFilter(XLColor.Red);
        Assert.That(autoFilter.Column(1).FilterType, Is.EqualTo(XLFilterType.Color));

        Assert.That(!ws.Row(2).IsHidden, "Row with red background should be visible");
        Assert.That(ws.Row(3).IsHidden, "Row with blue background should be hidden");
    }

    [Test]
    public void Color_filter_survives_save_and_load()
    {
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = "Data";
                ws.Cell("A2").Value = "Red";
                ws.Cell("A2").Style.Fill.SetBackgroundColor(XLColor.Red);
                ws.Cell("A3").Value = "Blue";
                ws.Cell("A3").Style.Fill.SetBackgroundColor(XLColor.Blue);
                ws.Cell("A4").Value = "Red again";
                ws.Cell("A4").Style.Fill.SetBackgroundColor(XLColor.Red);

                ws.Range("A1:A4").SetAutoFilter().Column(1).ColorFilter(XLColor.Red);
            },
            (_, ws) =>
            {
                var filterColumn = ws.AutoFilter.Column(1);
                Assert.That(filterColumn.FilterType, Is.EqualTo(XLFilterType.Color));
                Assert.That(filterColumn.FilterByCellColor, Is.True);

                // After load, reapply to check it works
                ws.AutoFilter.Reapply();
                var visibility = ws.Rows("2:4").Select(row => !row.IsHidden);
                Assert.That(visibility, Is.EqualTo([true, false, true]));
            });
    }

    [Test]
    public void Font_color_filter_survives_save_and_load()
    {
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = "Data";
                ws.Cell("A2").Value = "Green";
                ws.Cell("A2").Style.Font.SetFontColor(XLColor.Green);
                ws.Cell("A3").Value = "Red";
                ws.Cell("A3").Style.Font.SetFontColor(XLColor.Red);

                ws.Range("A1:A3").SetAutoFilter().Column(1).FontColorFilter(XLColor.Green);
            },
            (_, ws) =>
            {
                var filterColumn = ws.AutoFilter.Column(1);
                Assert.That(filterColumn.FilterType, Is.EqualTo(XLFilterType.Color));
                Assert.That(filterColumn.FilterByCellColor, Is.False);

                ws.AutoFilter.Reapply();
                var visibility = ws.Rows("2:3").Select(row => !row.IsHidden);
                Assert.That(visibility, Is.EqualTo([true, false]));
            });
    }

    [Test]
    public void Color_filter_with_theme_color_survives_save_and_load()
    {
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = "Data";
                ws.Cell("A2").Value = "Theme";
                ws.Cell("A2").Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent1));
                ws.Cell("A3").Value = "Other";
                ws.Cell("A3").Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent2));

                ws.Range("A1:A3").SetAutoFilter().Column(1)
                    .ColorFilter(XLColor.FromTheme(XLThemeColor.Accent1));
            },
            (_, ws) =>
            {
                var filterColumn = ws.AutoFilter.Column(1);
                Assert.That(filterColumn.FilterType, Is.EqualTo(XLFilterType.Color));
                Assert.That(filterColumn.FilterByCellColor, Is.True);

                ws.AutoFilter.Reapply();
                Assert.That(!ws.Row(2).IsHidden, "Theme color match should be visible");
                Assert.That(ws.Row(3).IsHidden, "Non-matching theme color should be hidden");
            });
    }

    [Test]
    public void Color_filter_works_with_multi_column_filter()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Color";
        ws.Cell("B1").Value = "Value";
        ws.Cell("A2").Value = "Red";
        ws.Cell("A2").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Cell("B2").Value = 10;
        ws.Cell("A3").Value = "Blue";
        ws.Cell("A3").Style.Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("B3").Value = 20;
        ws.Cell("A4").Value = "Red";
        ws.Cell("A4").Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.Cell("B4").Value = 30;

        var autoFilter = ws.Range("A1:B4").SetAutoFilter();
        // Color filter on column A + value filter on column B
        autoFilter.Column(1).ColorFilter(XLColor.Red, reapply: false);
        autoFilter.Column(2).AddFilter(30);

        // Row 2: Red + 10 → color matches, value doesn't → hidden
        Assert.That(ws.Row(2).IsHidden, "Row with Red/10 should be hidden (value doesn't match)");
        // Row 3: Blue + 20 → color doesn't match → hidden
        Assert.That(ws.Row(3).IsHidden, "Row with Blue/20 should be hidden (color doesn't match)");
        // Row 4: Red + 30 → both match → visible
        Assert.That(!ws.Row(4).IsHidden, "Row with Red/30 should be visible (both match)");
    }
}
