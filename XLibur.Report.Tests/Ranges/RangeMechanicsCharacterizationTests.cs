using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// Pins the XLibur behaviour range expansion is built on.
/// </summary>
/// <remarks>
/// The expander repeats a template row by inserting rows and copying into them, which delegates
/// formula adjustment, merge tracking, conditional-format extension and defined-name shifting to
/// the core library rather than reimplementing them. These tests state what it is relying on, so
/// a change in the core surfaces here rather than as a mystifying report failure.
/// </remarks>
public class RangeMechanicsCharacterizationTests
{
    [Test]
    public async Task CopyingARangeDownAdjustsRelativeFormulas()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A1").Value = 2;
        sheet.Cell("B1").FormulaA1 = "A1*2";

        sheet.Range("A1:B1").CopyTo(sheet.Cell("A5"));

        await Assert.That(sheet.Cell("B5").FormulaA1).IsEqualTo("A5*2");
    }

    [Test]
    public async Task CopyingARangeCarriesStyleAndMerges()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A1").Value = "x";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Range("B1:C1").Merge();

        sheet.Range("A1:C1").CopyTo(sheet.Cell("A5"));

        await Assert.That(sheet.Cell("A5").Style.Font.Bold).IsTrue();
        await Assert.That(sheet.MergedRanges.Any(r => r.RangeAddress.ToStringRelative() == "B5:C5")).IsTrue();
    }

    [Test]
    public async Task InsertingRowsShiftsContentBelowDown()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A1").Value = "first";
        sheet.Cell("A2").Value = "second";

        sheet.Row(1).InsertRowsBelow(2);

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("first");
        await Assert.That(sheet.Cell("A4").Value.GetText()).IsEqualTo("second");
    }

    /// <summary>
    /// The expander re-reads a defined name's address before expanding it, which only works if
    /// inserting rows moves names that sit below the insertion point.
    /// </summary>
    [Test]
    public async Task InsertingRowsShiftsDefinedNamesBelow()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A5").Value = "target";
        workbook.DefinedNames.Add("Later", sheet.Range("A5:A6"));

        sheet.Row(1).InsertRowsBelow(3);

        var range = workbook.DefinedNames.Single(n => n.Name == "Later").Ranges.Single();
        await Assert.That(range.RangeAddress.ToStringRelative()).IsEqualTo("A8:A9");
    }

    [Test]
    public async Task DeletingRowsShrinksADefinedNameThatSpansThem()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        workbook.DefinedNames.Add("Block", sheet.Range("A2:A8"));

        sheet.Rows(5, 6).Delete();

        var range = workbook.DefinedNames.Single(n => n.Name == "Block").Ranges.Single();
        await Assert.That(range.RangeAddress.ToStringRelative()).IsEqualTo("A2:A6");
    }

    [Test]
    public async Task InsertedRowsCanBeGivenTheStyleOfTheRowAbove()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A1").Value = "x";
        sheet.Row(1).Style.Font.Bold = true;

        sheet.Row(1).InsertRowsBelow(1);

        await Assert.That(sheet.Row(2).Style.Font.Bold).IsTrue();
    }

    [Test]
    public async Task ConditionalFormatRangeGrowsWhenRowsAreInsertedInsideIt()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(10).Fill.SetBackgroundColor(XLColor.Red);

        sheet.Row(2).InsertRowsBelow(3);

        var format = sheet.ConditionalFormats.Single();
        await Assert.That(format.Ranges.Single().RangeAddress.ToStringRelative()).IsEqualTo("A1:A8");
    }

    [Test]
    public async Task RowHeightIsCarriedByCopy()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("S");
        sheet.Cell("A1").Value = "x";
        sheet.Row(1).Height = 33;

        sheet.Row(1).CopyTo(sheet.Row(4));

        await Assert.That(sheet.Row(4).Height).IsEqualTo(33d);
    }
}
