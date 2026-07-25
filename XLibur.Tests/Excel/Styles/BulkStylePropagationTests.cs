using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// Styling a range, row or column writes the style slice point by point instead of materialising an
/// <c>XLCell</c> per cell. These tests pin the semantics that rewrite has to preserve — which cells
/// are covered, that each cell is modified from its own prior style, and that the container's own
/// style is updated from the value it had before any cell was touched.
/// </summary>
[TestFixture]
public class BulkStylePropagationTests
{
    [Test]
    public void Range_style_covers_every_cell_in_the_rectangle_including_empty_ones()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;

        ws.Range("B2:D4").Style.Font.Bold = true;

        for (var row = 2; row <= 4; row++)
        {
            for (var column = 2; column <= 4; column++)
                Assert.IsTrue(ws.Cell(row, column).Style.Font.Bold, $"cell {row},{column}");
        }

        Assert.IsFalse(ws.Cell("A1").Style.Font.Bold);
        Assert.IsFalse(ws.Cell("E5").Style.Font.Bold);
        Assert.IsTrue(ws.Range("B2:D4").Style.Font.Bold);
    }

    /// <summary>
    /// Each cell must be modified from its <em>own</em> prior style. The fast path memoises the
    /// transform across runs of identical styles, so alternating styles are the case that would
    /// expose a memo smearing one cell's result onto its neighbour.
    /// </summary>
    [Test]
    public void Range_style_modifies_each_cell_from_its_own_prior_style()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        for (var column = 1; column <= 6; column++)
            ws.Cell(1, column).Style.Font.FontSize = column % 2 == 0 ? 8 : 20;

        ws.Range(1, 1, 1, 6).Style.Font.Bold = true;

        for (var column = 1; column <= 6; column++)
        {
            var style = ws.Cell(1, column).Style;
            Assert.IsTrue(style.Font.Bold, $"column {column}");
            Assert.AreEqual(column % 2 == 0 ? 8 : 20, style.Font.FontSize, $"column {column}");
        }
    }

    [Test]
    public void Row_style_covers_used_cells_and_the_row_itself()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(3, 1).Value = "a";
        ws.Cell(3, 5).Value = "b";

        ws.Row(3).Style.Font.Italic = true;

        Assert.IsTrue(ws.Row(3).Style.Font.Italic);
        Assert.IsTrue(ws.Cell(3, 1).Style.Font.Italic);
        Assert.IsTrue(ws.Cell(3, 5).Style.Font.Italic);

        // An unused cell in the row inherits from the row, so it reads as italic too...
        Assert.IsTrue(ws.Cell(3, 2).Style.Font.Italic);

        // ...but a different row is untouched.
        Assert.IsFalse(ws.Cell(4, 1).Style.Font.Italic);
    }

    [Test]
    public void Column_style_covers_used_cells_and_the_column_itself()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 3).Value = "a";
        ws.Cell(7, 3).Value = "b";

        ws.Column(3).Style.Font.Strikethrough = true;

        Assert.IsTrue(ws.Column(3).Style.Font.Strikethrough);
        Assert.IsTrue(ws.Cell(1, 3).Style.Font.Strikethrough);
        Assert.IsTrue(ws.Cell(7, 3).Style.Font.Strikethrough);
        Assert.IsTrue(ws.Cell(4, 3).Style.Font.Strikethrough);
        Assert.IsFalse(ws.Cell(1, 4).Style.Font.Strikethrough);
    }

    /// <summary>
    /// The container's own style must be modified from the value it had <em>before</em> the cells
    /// were touched. For a row this is load-bearing: unstyled cells inherit the row's style, so
    /// writing the row first would change what those cells resolve to mid-operation.
    /// </summary>
    [Test]
    public void Row_style_modification_does_not_leak_into_its_cells_mid_operation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Row(2).Style.Font.FontSize = 30;
        ws.Cell(2, 1).Value = "explicit";
        ws.Cell(2, 1).Style.Font.FontSize = 11;
        ws.Cell(2, 2).Value = "inherits";

        ws.Row(2).Style.Font.Bold = true;

        Assert.AreEqual(11, ws.Cell(2, 1).Style.Font.FontSize);
        Assert.IsTrue(ws.Cell(2, 1).Style.Font.Bold);

        Assert.AreEqual(30, ws.Cell(2, 2).Style.Font.FontSize);
        Assert.IsTrue(ws.Cell(2, 2).Style.Font.Bold);

        Assert.AreEqual(30, ws.Row(2).Style.Font.FontSize);
        Assert.IsTrue(ws.Row(2).Style.Font.Bold);
    }

    /// <summary>
    /// The worksheet opts out of the fast path — its children are rows and columns, not the whole
    /// address rectangle — so sheet-wide styling must still reach everything.
    /// </summary>
    [Test]
    public void Worksheet_style_still_propagates_to_rows_columns_and_cells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "a";
        ws.Cell(5, 5).Value = "b";
        ws.Row(3).Height = 30;
        ws.Column(4).Width = 20;

        ws.Style.Font.Bold = true;

        Assert.IsTrue(ws.Style.Font.Bold);
        Assert.IsTrue(ws.Cell(1, 1).Style.Font.Bold);
        Assert.IsTrue(ws.Cell(5, 5).Style.Font.Bold);
        Assert.IsTrue(ws.Row(3).Style.Font.Bold);
        Assert.IsTrue(ws.Column(4).Style.Font.Bold);
    }

    /// <summary>
    /// Assigning a whole style (rather than modifying one component) takes the same fast path with
    /// an absolute transform.
    /// </summary>
    [Test]
    public void Assigning_a_whole_style_to_a_range_overwrites_every_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Style.Font.FontSize = 30;
        ws.Cell("B1").Style.Font.Italic = true;

        var source = ws.Cell("Z100");
        source.Style.Font.Bold = true;
        source.Style.Font.FontSize = 14;

        ws.Range("A1:B2").Style = source.Style;

        foreach (var address in new[] { "A1", "A2", "B1", "B2" })
        {
            var style = ws.Cell(address).Style;
            Assert.IsTrue(style.Font.Bold, address);
            Assert.AreEqual(14, style.Font.FontSize, address);
            Assert.IsFalse(style.Font.Italic, address);
        }
    }

    [Test]
    public void Range_style_survives_a_save_and_reload()
    {
        using var ms = new System.IO.MemoryStream();

        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "x";
            ws.Range("A1:C3").Style.Fill.BackgroundColor = XLColor.Red;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var sheet = reloaded.Worksheet("Data");
        for (var row = 1; row <= 3; row++)
        {
            for (var column = 1; column <= 3; column++)
                Assert.AreEqual(XLColor.Red, sheet.Cell(row, column).Style.Fill.BackgroundColor, $"{row},{column}");
        }
    }
}
