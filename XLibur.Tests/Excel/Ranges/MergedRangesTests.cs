using System.Linq;
using XLibur.Excel;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class MergedRangesTests
{
    [Test]
    public async Task LastCellFromMerge()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");
        ws.Range("B2:D4").Merge();

        var first = ws.FirstCellUsed(XLCellsUsedOptions.All).Address.ToStringRelative();
        var last = ws.LastCellUsed(XLCellsUsedOptions.All).Address.ToStringRelative();

        await Assert.That(first).IsEqualTo("B2");
        await Assert.That(last).IsEqualTo("D4");
    }

    [Test]
    [Arguments("A1:A2", "A1:A2")]
    [Arguments("A2:B2", "A2:B2")]
    [Arguments("A3:C3", "A3:E3")]
    [Arguments("B4:B6", "B4:B6")]
    [Arguments("C7:D7", "E7:F7")]
    public async Task MergedRangesShiftedOnColumnInsert(string originalRange, string expectedRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        var range = ws.Range(originalRange).Merge();

        ws.Column(2).InsertColumnsAfter(2);

        var mr = ws.MergedRanges.ToArray();
        await Assert.That(mr.Length).IsEqualTo(1);
        await Assert.That(mr.Single()).IsSameReferenceAs(range);
        await Assert.That(range.RangeAddress.ToString()).IsEqualTo(expectedRange);
    }

    [Test]
    [Arguments("A1:B1", "A1:B1")]
    [Arguments("B1:B2", "B1:B2")]
    [Arguments("C1:C3", "C1:C5")]
    [Arguments("D2:F2", "D2:F2")]
    [Arguments("G4:G5", "G6:G7")]
    public async Task MergedRangesShiftedOnRowInsert(string originalRange, string expectedRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        var range = ws.Range(originalRange).Merge();

        ws.Row(2).InsertRowsBelow(2);

        var mr = ws.MergedRanges.ToArray();
        await Assert.That(mr.Length).IsEqualTo(1);
        await Assert.That(mr.Single()).IsSameReferenceAs(range);
        await Assert.That(range.RangeAddress.ToString()).IsEqualTo(expectedRange);
    }

    [Test]
    [Arguments("A1:A2", true, "A1:A2")]
    [Arguments("A2:B2", true, "A2:A2")]
    [Arguments("A3:C3", true, "A3:B3")]
    [Arguments("B4:B6", false, "")]
    [Arguments("C7:D7", true, "B7:C7")]
    public async Task MergedRangesShiftedOnColumnDelete(string originalRange, bool expectedExist, string expectedRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        var range = ws.Range(originalRange).Merge();

        ws.Column(2).Delete();

        var mr = ws.MergedRanges.ToArray();
        if (expectedExist)
        {
            await Assert.That(mr.Length).IsEqualTo(1);
            await Assert.That(mr.Single()).IsSameReferenceAs(range);
            await Assert.That(range.RangeAddress.ToString()).IsEqualTo(expectedRange);
        }
        else
        {
            await Assert.That(mr.Length).IsEqualTo(0);
            await Assert.That(range.RangeAddress.IsValid).IsFalse();
        }
    }

    [Test]
    [Arguments("A1:B1", true, "A1:B1")]
    [Arguments("B1:B2", true, "B1:B1")]
    [Arguments("C1:C3", true, "C1:C2")]
    [Arguments("D2:F2", false, "")]
    [Arguments("G4:G5", true, "G3:G4")]
    public async Task MergedRangesShiftedOnRowDelete(string originalRange, bool expectedExist, string expectedRange)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        var range = ws.Range(originalRange).Merge();

        ws.Row(2).Delete();

        var mr = ws.MergedRanges.ToArray();
        if (expectedExist)
        {
            await Assert.That(mr.Length).IsEqualTo(1);
            await Assert.That(mr.Single()).IsSameReferenceAs(range);
            await Assert.That(range.RangeAddress.ToString()).IsEqualTo(expectedRange);
        }
        else
        {
            await Assert.That(mr.Length).IsEqualTo(0);
            await Assert.That(range.RangeAddress.IsValid).IsFalse();
        }
    }

    [Test]
    public async Task ShiftRangeRightBreaksMerges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        ws.Range("B2:C3").Merge();
        ws.Range("B4:C5").Merge();
        ws.Range("F2:G3").Merge(); // to be broken
        ws.Range("F4:G5").Merge(); // to be broken
        ws.Range("H1:I2").Merge();
        ws.Range("H5:I6").Merge();

        ws.Range("D3:E4").InsertColumnsAfter(2);

        var mr = ws.MergedRanges.ToArray();
        await Assert.That(mr.Length).IsEqualTo(4);
        await Assert.That(mr[0].RangeAddress.ToString()).IsEqualTo("H1:I2");
        await Assert.That(mr[1].RangeAddress.ToString()).IsEqualTo("B2:C3");
        await Assert.That(mr[2].RangeAddress.ToString()).IsEqualTo("B4:C5");
        await Assert.That(mr[3].RangeAddress.ToString()).IsEqualTo("H5:I6");
    }

    [Test]
    public async Task ShiftRangeLeftBreaksMerges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        ws.Range("B2:C3").Merge();
        ws.Range("B4:C5").Merge();
        ws.Range("F2:G3").Merge(); // to be broken
        ws.Range("F4:G5").Merge(); // to be broken
        ws.Range("H1:I2").Merge();
        ws.Range("H5:I6").Merge();

        ws.Range("D3:E4").Delete(XLShiftDeletedCells.ShiftCellsLeft);

        var mr = ws.MergedRanges.ToArray();
        await Assert.That(mr.Length).IsEqualTo(4);
        await Assert.That(mr[0].RangeAddress.ToString()).IsEqualTo("H1:I2");
        await Assert.That(mr[1].RangeAddress.ToString()).IsEqualTo("B2:C3");
        await Assert.That(mr[2].RangeAddress.ToString()).IsEqualTo("B4:C5");
        await Assert.That(mr[3].RangeAddress.ToString()).IsEqualTo("H5:I6");
    }

    [Test]
    public async Task RangeShiftDownBreaksMerges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        ws.Range("B2:C3").Merge();
        ws.Range("D2:E3").Merge();
        ws.Range("B6:C7").Merge(); // to be broken
        ws.Range("D6:E7").Merge(); // to be broken
        ws.Range("A8:B9").Merge();
        ws.Range("E8:F9").Merge();

        ws.Range("C4:D5").InsertRowsBelow(2);

        var mr = ws.MergedRanges.ToArray();
        await Assert.That(mr.Length).IsEqualTo(4);
        await Assert.That(mr[0].RangeAddress.ToString()).IsEqualTo("B2:C3");
        await Assert.That(mr[1].RangeAddress.ToString()).IsEqualTo("D2:E3");
        await Assert.That(mr[2].RangeAddress.ToString()).IsEqualTo("A8:B9");
        await Assert.That(mr[3].RangeAddress.ToString()).IsEqualTo("E8:F9");
    }

    [Test]
    public async Task RangeShiftUpBreaksMerges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MRShift");
        ws.Range("B2:C3").Merge();
        ws.Range("D2:E3").Merge();
        ws.Range("B6:C7").Merge(); // to be broken
        ws.Range("D6:E7").Merge(); // to be broken
        ws.Range("A8:B9").Merge();
        ws.Range("E8:F9").Merge();

        ws.Range("C4:D5").Delete(XLShiftDeletedCells.ShiftCellsUp);

        var mr = ws.MergedRanges.ToArray();
        await Assert.That(mr.Length).IsEqualTo(4);
        await Assert.That(mr[0].RangeAddress.ToString()).IsEqualTo("B2:C3");
        await Assert.That(mr[1].RangeAddress.ToString()).IsEqualTo("D2:E3");
        await Assert.That(mr[2].RangeAddress.ToString()).IsEqualTo("A8:B9");
        await Assert.That(mr[3].RangeAddress.ToString()).IsEqualTo("E8:F9");
    }

    [Test]
    public async Task MergedCellsAcquireFirstCellStyle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.Red;
        ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
        ws.Cell("A3").Style.Fill.BackgroundColor = XLColor.Green;
        ws.Range("A1:A3").Merge();

        await Assert.That(ws.Cell("A1").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("A2").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(ws.Cell("A3").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task MergedCellsLooseData()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("A1:A3").SetValue(100);
        ws.Range("A1:A3").Merge();

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(100);
        await Assert.That(ws.Cell("A2").Value).IsEqualTo(Blank.Value);
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(Blank.Value);
    }

    [Test]
    public async Task MergedCellsLooseConditionalFormats()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").AddConditionalFormat().WhenContains("1").Fill.BackgroundColor = XLColor.Red;
        ws.Cell("A2").AddConditionalFormat().WhenContains("2").Fill.BackgroundColor = XLColor.Yellow;

        ws.Range("A1:A2").Merge();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A1");
    }

    [Test]
    public async Task MergedCellsLooseDataValidation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").CreateDataValidation().WholeNumber.Between(1, 2);
        ws.Cell("A2").CreateDataValidation().Date.GreaterThan(new System.DateTime(2018, 1, 1));

        ws.Range("A1:A2").Merge();

        await Assert.That(ws.Cell("A1").HasDataValidation).IsTrue();
        await Assert.That(ws.Cell("A1").GetDataValidation().MinValue).IsEqualTo("1");
        await Assert.That(ws.Cell("A1").GetDataValidation().MaxValue).IsEqualTo("2");
        await Assert.That(ws.Cell("A2").HasDataValidation).IsFalse();
    }

    [Test]
    public async Task UnmergedCellsPreserveStyle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var range = ws.Range("B2:D4");
        range.Style.Fill.SetBackgroundColor(XLColor.Yellow);
        range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XLColor.DarkBlue)
            .Border.SetInsideBorder(XLBorderStyleValues.Thin)
            .Border.SetInsideBorderColor(XLColor.Pink);
        range.Cells().ForEach(c => c.Value = c.Address.ToString());

        var firstCell = ws.Cell("B2");
        firstCell.Style.Fill.SetBackgroundColor(XLColor.Red)
            .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Font.SetBold();

        range.Merge();
        range.Unmerge();

        await Assert.That(range.Cells().All(c => c.Style.Fill.BackgroundColor == XLColor.Red)).IsTrue();
        await Assert.That(range.Cells().Where(c => !c.Equals(firstCell)).All(c => c.Value.Equals(Blank.Value))).IsTrue();
        await Assert.That(firstCell.Value).IsEqualTo("B2");

        await Assert.That(ws.Cell("B2").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("B2").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B2").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);

        await Assert.That(ws.Cell("C2").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("C2").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C2").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);

        await Assert.That(ws.Cell("D2").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("D2").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("D2").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("D2").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);

        await Assert.That(ws.Cell("B3").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B3").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B3").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B3").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);

        await Assert.That(ws.Cell("C3").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C3").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C3").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C3").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);

        await Assert.That(ws.Cell("D3").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("D3").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("D3").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("D3").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);

        await Assert.That(ws.Cell("B4").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B4").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("B4").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("B4").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.Thick);

        await Assert.That(ws.Cell("C4").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C4").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("C4").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("C4").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);

        await Assert.That(ws.Cell("D4").Style.Border.TopBorder).IsEqualTo(XLBorderStyleValues.None);
        await Assert.That(ws.Cell("D4").Style.Border.RightBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("D4").Style.Border.BottomBorder).IsEqualTo(XLBorderStyleValues.Thick);
        await Assert.That(ws.Cell("D4").Style.Border.LeftBorder).IsEqualTo(XLBorderStyleValues.None);
    }

    [Test]
    public async Task MergedRangesCellValuesShouldNotBeSet()
    {
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").Value = 1;
            ws.Cell("A3").Value = 1;
            ws.Cell("A4").Value = 1;
            ws.Cell("B1").FormulaA1 = "SUM(A:A)";
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        }

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge().SetValue(1);
            ws.Cell("B1").FormulaA1 = "SUM(A:A)";
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        }
    }

    [Test]
    public async Task MergedRangesCellFormulasShouldNotBeSet()
    {
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").FormulaA1 = "=1";
            ws.Cell("A3").FormulaA1 = "=1";
            ws.Cell("A4").FormulaA1 = "=1";
            ws.Cell("B1").FormulaA1 = "SUM(A:A)";
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        }

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").SetFormulaA1("=1");
            ws.Cell("A3").SetFormulaA1("=1");
            ws.Cell("A4").SetFormulaA1("=1");
            ws.Cell("B1").SetFormulaA1("SUM(A:A)");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        }

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").FormulaR1C1 = "=1";
            ws.Cell("A3").FormulaR1C1 = "=1";
            ws.Cell("A4").FormulaR1C1 = "=1";
            ws.Cell("B1").FormulaR1C1 = "SUM(A:A)";
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        }

        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.AddWorksheet();
            ws.Range("A2:A4").Merge();
            ws.Cell("A2").SetFormulaR1C1("=1");
            ws.Cell("A3").SetFormulaR1C1("=1");
            ws.Cell("A4").SetFormulaR1C1("=1");
            ws.Cell("B1").SetFormulaR1C1("SUM(A:A)");
            await Assert.That(ws.Cell("B1").Value).IsEqualTo(1);
        }
    }

    [Test]
    public async Task FormulaReference_setter_silently_ignores_inferior_merged_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A3").Merge();

        // Setting formula on superior cell works
        ws.Cell("A1").FormulaA1 = "=1+2";
        ws.Cell("A1").FormulaReference = ws.Range("A1:A3").RangeAddress;

        // Setting FormulaReference on an inferior merged cell without a formula
        // should not throw (consistent with FormulaA1 setter behavior)
        await Assert.That(() => ws.Cell("A2").FormulaReference = ws.Range("A1:A3").RangeAddress).ThrowsNothing();
        await Assert.That(() => ws.Cell("A3").FormulaReference = ws.Range("A1:A3").RangeAddress).ThrowsNothing();
    }

    [Test]
    public async Task MergeSingleCellRangeDoesNothing()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range(1, 1, 1, 1);

        range.Merge();

        await Assert.That(range.IsMerged()).IsFalse();
        await Assert.That(ws.MergedRanges.Count).IsEqualTo(0);
    }

    /// <summary>
    /// <c>IsMerged</c> short-circuits on the merged-range count before consulting the range index,
    /// because it runs on every value and formula assignment. The short-circuit is only sound while
    /// the count tracks the index exactly, so pin the transitions: none → merged → unmerged →
    /// merged again.
    /// </summary>
    [Test]
    public async Task IsMerged_tracks_merge_and_unmerge_transitions()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(ws.MergedRanges.Count).IsEqualTo(0);
        await Assert.That(ws.Cell("A1").IsMerged()).IsFalse();
        await Assert.That(ws.Cell("A2").IsMerged()).IsFalse();

        ws.Range("A1:A3").Merge();
        await Assert.That(ws.MergedRanges.Count).IsEqualTo(1);
        await Assert.That(ws.Cell("A1").IsMerged()).IsTrue();
        await Assert.That(ws.Cell("A2").IsMerged()).IsTrue();
        await Assert.That(ws.Cell("B1").IsMerged()).IsFalse();

        ws.Range("A1:A3").Unmerge();
        await Assert.That(ws.MergedRanges.Count).IsEqualTo(0);
        await Assert.That(ws.Cell("A1").IsMerged()).IsFalse();
        await Assert.That(ws.Cell("A2").IsMerged()).IsFalse();

        ws.Range("A1:A3").Merge();
        await Assert.That(ws.MergedRanges.Count).IsEqualTo(1);
        await Assert.That(ws.Cell("A2").IsMerged()).IsTrue();
    }

    /// <summary>
    /// The guarded path must keep ignoring writes to inferior merged cells, and the unguarded path
    /// must keep accepting them once the merge is gone.
    /// </summary>
    [Test]
    public async Task Inferior_merged_cell_writes_are_ignored_only_while_merged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // No merges: every write lands.
        ws.Cell("A2").Value = 1;
        ws.Cell("A3").FormulaA1 = "=1";
        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(1);
        await Assert.That("=" + ws.Cell("A3").FormulaA1).IsEqualTo("=1");

        ws.Cell("A2").Clear();
        ws.Cell("A3").Clear();
        ws.Range("A1:A3").Merge();

        // Merged: writes to the inferior cells are silently dropped.
        ws.Cell("A2").Value = 42;
        ws.Cell("A3").FormulaA1 = "=99";
        await Assert.That(ws.Cell("A2").DataType).IsEqualTo(XLDataType.Blank);
        await Assert.That(ws.Cell("A3").FormulaA1).IsEqualTo(string.Empty);

        // The superior cell still accepts them.
        ws.Cell("A1").Value = 7;
        await Assert.That(ws.Cell("A1").GetDouble()).IsEqualTo(7);

        // Unmerged: the inferior cells accept writes again.
        ws.Range("A1:A3").Unmerge();
        ws.Cell("A2").Value = 42;
        await Assert.That(ws.Cell("A2").GetDouble()).IsEqualTo(42);
    }
}
