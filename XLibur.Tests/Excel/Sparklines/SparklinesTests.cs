using XLibur.Examples.Sparklines;
using XLibur.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Sparklines;

public class SparklinesTests
{
    #region Add sparklines

    [Test]
    public async Task CannotCreateSparklineGroupsWithoutWorksheet()
    {
        Action action = () => _ = new XLSparklineGroups(null!);
        await Assert.That(action).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task CannotCreateSparklineGroupWithoutWorksheet()
    {
        Action action = () => _ = new XLSparklineGroup(null!);
        await Assert.That(action).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task CannotCreateSparklineWithoutGroup()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        Action action = () => _ = new XLSparkline(null!, ws.Cell("A1"), ws.Range("A2:A5"));
        await Assert.That(action).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task CannotCreateSparklineWithoutLocation()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var group = new XLSparklineGroup(ws);
        Action action = () => _ = new XLSparkline(group, null!, ws.Range("A2:A5"));
        await Assert.That(action).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task CanCreateInvalidSparklineWithoutSourceData()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");
        var group = new XLSparklineGroup(ws);
        var sparkline = new XLSparkline(group, ws.FirstCell(), null);
        await Assert.That(sparkline.IsValid).IsFalse();
    }

    [Test]
    public async Task CanAddSparklineGroupForSingleCell()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add(new XLSparklineGroup(ws, "A1", "B1:E1"));
        ws.SparklineGroups.Add("A2", "B2:E2");
        ws.SparklineGroups.Add(ws.Cell("A3"), ws.Range("B3:E3"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(3);

        await Assert.That(ws.SparklineGroups.ElementAt(0).Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.ElementAt(1).Single().Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(ws.SparklineGroups.ElementAt(2).Single().Location.Address.ToString()).IsEqualTo("A3");

        await Assert.That(ws.SparklineGroups.ElementAt(0).Single().SourceData.RangeAddress.ToString()).IsEqualTo("B1:E1");
        await Assert.That(ws.SparklineGroups.ElementAt(1).Single().SourceData.RangeAddress.ToString()).IsEqualTo("B2:E2");
        await Assert.That(ws.SparklineGroups.ElementAt(2).Single().SourceData.RangeAddress.ToString()).IsEqualTo("B3:E3");

        await Assert.That(ws.SparklineGroups.All(g => g.Worksheet == ws)).IsTrue();
    }

    [Test]
    public async Task CanAddSparklineGroupForVerticalRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add(ws.Range("A1:A3"), ws.Range("B1:E3"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);

        await Assert.That(ws.SparklineGroups.Single().ElementAt(0).Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(1).Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(2).Location.Address.ToString()).IsEqualTo("A3");

        await Assert.That(ws.SparklineGroups.Single().ElementAt(0).SourceData.RangeAddress.ToString()).IsEqualTo("B1:E1");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(1).SourceData.RangeAddress.ToString()).IsEqualTo("B2:E2");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(2).SourceData.RangeAddress.ToString()).IsEqualTo("B3:E3");
    }

    [Test]
    public async Task CanAddSparklineGroupForHorizontalRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add(ws.Range("A1:C1"), ws.Range("A2:C4"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);

        await Assert.That(ws.SparklineGroups.Single().ElementAt(0).Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(1).Location.Address.ToString()).IsEqualTo("B1");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(2).Location.Address.ToString()).IsEqualTo("C1");

        await Assert.That(ws.SparklineGroups.Single().ElementAt(0).SourceData.RangeAddress.ToString()).IsEqualTo("A2:A4");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(1).SourceData.RangeAddress.ToString()).IsEqualTo("B2:B4");
        await Assert.That(ws.SparklineGroups.Single().ElementAt(2).SourceData.RangeAddress.ToString()).IsEqualTo("C2:C4");
    }

    [Test]
    public async Task CannotAddSparklineForNonLinearRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        Action action = () => ws.SparklineGroups.Add(ws.Range("A1:C2"), ws.Range("A3:C4"));

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("locationRange must have either a single row or a single column");
    }

    [Test]
    public async Task CannotAddSparklineWhenRangesHaveDifferentWidths()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        Action action = () => ws.SparklineGroups.Add(ws.Range("A1:C1"), ws.Range("A3:D4"));

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("locationRange and sourceDataRange must have the same width");
    }

    [Test]
    public async Task CannotAddSparklineWhenRangesHaveDifferentHeights()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        Action action = () => ws.SparklineGroups.Add(ws.Range("A1:A3"), ws.Range("B1:B4"));

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("locationRange and sourceDataRange must have the same height");
    }

    [Test]
    public async Task CannotAddSparklineForCellWhenDataRangeIsNotLinear()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        Action action = () => ws.SparklineGroups.Add(ws.Range("A1:A1"), ws.Range("B1:C4"));

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("SourceData range must have either a single row or a single column");
    }

    [Test]
    public async Task CanAddSparklineToExistingGroup()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        var group = new XLSparklineGroup(ws)
        {
            { "A2", "B2:E2" },
            { ws.Cell("A3"), ws.Range("B3:E3") }
        };

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(0);

        await Assert.That(group.ElementAt(0).Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(group.ElementAt(1).Location.Address.ToString()).IsEqualTo("A3");

        await Assert.That(group.ElementAt(0).SourceData.RangeAddress.ToString()).IsEqualTo("B2:E2");
        await Assert.That(group.ElementAt(1).SourceData.RangeAddress.ToString()).IsEqualTo("B3:E3");
    }

    [Test]
    public async Task CannotAddSparklineGroupFromDifferentWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws2 = wb.AddWorksheet("Sheet 2");

        var group = new XLSparklineGroup(ws1);

        Action action = () => ws2.SparklineGroups.Add(group);

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("The specified sparkline group belongs to the different worksheet");
    }

    [Test]
    public async Task CannotAddSparklineFromDifferentWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws2 = wb.AddWorksheet("Sheet 2");

        var group = new XLSparklineGroup(ws1);

        Action action = () => group.Add(ws2.Cell("A3"), ws1.Range("B3:E3"));

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("The specified sparkline belongs to the different worksheet");
    }

    [Test]
    public async Task AddSparklineToSameCellOverwritesItWhenSameGroup()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        var group = ws.SparklineGroups.Add("A1", "B1:E1");
        group.Add("A1", "B2:E2");

        await Assert.That(group.Count()).IsEqualTo(1);

        await Assert.That(group.Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(group.Single().SourceData.RangeAddress.ToString()).IsEqualTo("B2:E2");
    }

    [Test]
    public async Task AddSparklineToSameCellOverwritesItWhenDifferentGroup()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1", "B1:E1");
        ws.SparklineGroups.Add("A1", "B2:E2");

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.First().Any()).IsFalse();
        await Assert.That(ws.SparklineGroups.Last().Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Last().Single().SourceData.RangeAddress.ToString()).IsEqualTo("B2:E2");
    }

    [Test]
    public async Task CanAddSparklineReferringToDifferentWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws3 = wb.AddWorksheet("Sheet 3");

        var group = ws1.SparklineGroups.Add("A1", "'Sheet 3'!B1:F1");

        await Assert.That(group.Single().SourceData.Worksheet).IsSameReferenceAs(ws3);
    }

    #endregion Add sparklines

    #region Get sparklines

    [Test]
    [Arguments("A2", "B2:Z2")]
    [Arguments("A50", "B50:Z50")]
    [Arguments("A100", "B100:Z100")]
    [Arguments("B1", "B2:B100")]
    [Arguments("K1", "K2:K100")]
    [Arguments("Z1", "Z2:Z100")]
    public async Task CanGetSparklineForExistingCell(string cellAddress, string expectedSourceDataRange)
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A2:A100", "B2:Z100");
        ws.SparklineGroups.Add("B1:Z1", "B2:Z100");

        var sp = ws.SparklineGroups.GetSparkline(ws.Cell(cellAddress));
        await Assert.That(sp).IsNotNull();
        await Assert.That(sp.Location.Address.ToString()).IsEqualTo(cellAddress);
        await Assert.That(sp.SourceData.RangeAddress.ToString()).IsEqualTo(expectedSourceDataRange);
    }

    [Test]
    [Arguments("A1")]
    [Arguments("B2")]
    [Arguments("A101")]
    [Arguments("AA1")]
    public async Task CannotGetSparklineForNonExistingCell(string cellAddress)
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A2:A100", "B2:Z100");
        ws.SparklineGroups.Add("B1:Z1", "B2:Z100");

        var sp = ws.SparklineGroups.GetSparkline(ws.Cell(cellAddress));
        await Assert.That(sp).IsNull();
    }

    [Test]
    public async Task CanGetSparklinesForRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A2:A100", "B2:Z100");
        ws.SparklineGroups.Add("B1:Z1", "B2:Z100");

        var sparklines1 = ws.SparklineGroups.GetSparklines(ws.Range("A1:B2"));
        var sparklines2 = ws.SparklineGroups.GetSparklines(ws.Range("B2:E4"));
        var sparklines3 = ws.SparklineGroups.GetSparklines(ws.Range("A1:Z100"));
        var sparklines4 = ws.SparklineGroups.GetSparklines(ws.Range("A:A"));
        var sparklines5 = ws.SparklineGroups.GetSparklines(ws.Range("1:1"));

        await Assert.That(sparklines1.Count()).IsEqualTo(2);
        await Assert.That(sparklines2.Count()).IsEqualTo(0);
        await Assert.That(sparklines3.Count()).IsEqualTo(99 + 25);
        await Assert.That(sparklines4.Count()).IsEqualTo(99);
        await Assert.That(sparklines5.Count()).IsEqualTo(25);

        await Assert.That(sparklines1.First().Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(sparklines1.Last().Location.Address.ToString()).IsEqualTo("B1");
        await Assert.That(sparklines1.First().SourceData.RangeAddress.ToString()).IsEqualTo("B2:Z2");
        await Assert.That(sparklines1.Last().SourceData.RangeAddress.ToString()).IsEqualTo("B2:B100");
    }

    #endregion Get sparklines

    #region Remove sparklines

    [Test]
    public async Task CanRemoveSparklineFromCell()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A3", "B1:Z3");
        ws.SparklineGroups.Remove(ws.Cell("A2"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.Single().First().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().Last().Location.Address.ToString()).IsEqualTo("A3");
        await Assert.That(ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
        await Assert.That(ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString()).IsEqualTo("B3:Z3");
    }

    [Test]
    public async Task CanRemoveSparklineFromRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A5", "B1:Z5");
        ws.SparklineGroups.Remove(ws.Range("A2:D4"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.Single().First().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().Last().Location.Address.ToString()).IsEqualTo("A5");
        await Assert.That(ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
        await Assert.That(ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString()).IsEqualTo("B5:Z5");
    }

    [Test]
    public async Task RemoveSparklineFromEmptyCellDoesNothing()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A2", "B1:Z2");
        ws.SparklineGroups.Remove(ws.Cell("F2"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.Single().First().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().Last().Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
        await Assert.That(ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString()).IsEqualTo("B2:Z2");
    }

    #endregion Remove sparklines

    #region Change sparklines

    [Test]
    public async Task CanChangeSparklineLocationInsideWorksheet()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A2", "B1:Z2");
        ws.SparklineGroups.Single().Last().SetLocation(ws.Cell("F2"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.Single().First().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().Last().Location.Address.ToString()).IsEqualTo("F2");
        await Assert.That(ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
        await Assert.That(ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString()).IsEqualTo("B2:Z2");
        await Assert.That(ws.Cell("A1").HasSparkline).IsTrue();
        await Assert.That(ws.Cell("A2").HasSparkline).IsFalse();
        await Assert.That(ws.Cell("F2").HasSparkline).IsTrue();
    }

    [Test]
    public async Task ChangeSparklineLocationOverwritesExistingSparklineSameGroup()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A2", "B1:Z2");
        ws.SparklineGroups.Single().Last().SetLocation(ws.Cell("A1"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().Single().SourceData.RangeAddress.ToString()).IsEqualTo("B2:Z2");
    }

    [Test]
    public async Task ChangeSparklineLocationOverwritesExistingSparklineDifferentGroups()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A2", "B1:Z2");
        ws.SparklineGroups.Add("A3", "B3:Z3");
        ws.SparklineGroups.Last().Single().SetLocation(ws.Cell("A2"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.First().Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.First().Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.First().Single().SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
        await Assert.That(ws.SparklineGroups.Last().Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Last().Single().Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(ws.SparklineGroups.Last().Single().SourceData.RangeAddress.ToString()).IsEqualTo("B3:Z3");
    }

    [Test]
    public async Task CannotChangeSparklineLocationToAnotherWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws2 = wb.AddWorksheet("Sheet 2");

        var group = ws1.SparklineGroups.Add("A1:A2", "B1:Z2");

        Action action = () => group.First().SetLocation(ws2.FirstCell());

        var message = (await Assert.That(action).Throws<InvalidOperationException>())!.Message;
        await Assert.That(message).IsEqualTo("Cannot move the sparkline to a different worksheet");
    }

    [Test]
    public async Task CanChangeSparklineSourceDataInsideWorksheet()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");

        ws.SparklineGroups.Add("A1:A2", "B1:Z2");
        ws.SparklineGroups.Single().Last().SetSourceData(ws.Range("D4:D50"));

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws.SparklineGroups.Single().Count()).IsEqualTo(2);
        await Assert.That(ws.SparklineGroups.Single().First().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(ws.SparklineGroups.Single().Last().Location.Address.ToString()).IsEqualTo("A2");
        await Assert.That(ws.SparklineGroups.Single().First().SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
        await Assert.That(ws.SparklineGroups.Single().Last().SourceData.RangeAddress.ToString()).IsEqualTo("D4:D50");
    }

    [Test]
    public async Task CannotChangeSparklineSourceDataToNonLinearRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1", "B1:Z1");
        var sparkline = group.Single();

        Action action = () => sparkline.SetSourceData(ws.Range("B1:Z2"));

        var message = (await Assert.That(action).Throws<ArgumentException>())!.Message;
        await Assert.That(message).IsEqualTo("SourceData range must have either a single row or a single column");
    }

    [Test]
    public async Task CanChangeSparklineStyle()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1", "B1:Z1");

        group.Style = XLSparklineTheme.Colorful1;

        await Assert.That(group.Style.SeriesColor).IsEqualTo(XLColor.FromHtml("FF5F5F5F"));
        await Assert.That(group.Style.NegativeColor).IsEqualTo(XLColor.FromHtml("FFFFB620"));
        await Assert.That(group.Style.MarkersColor).IsEqualTo(XLColor.FromHtml("FFD70077"));
        await Assert.That(group.Style.HighMarkerColor).IsEqualTo(XLColor.FromHtml("FF56BE79"));
        await Assert.That(group.Style.LowMarkerColor).IsEqualTo(XLColor.FromHtml("FFFF5055"));
        await Assert.That(group.Style.FirstMarkerColor).IsEqualTo(XLColor.FromHtml("FF5687C2"));
        await Assert.That(group.Style.LastMarkerColor).IsEqualTo(XLColor.FromHtml("FF359CEB"));
    }

    [Test]
    public async Task ChangeSparklineStyleDoesNotAffectOriginal()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1", "B1:Z1");
        group.Style = XLSparklineTheme.Colorful1;

        group.Style.NegativeColor = XLColor.Red;

        await Assert.That(group.Style.NegativeColor).IsEqualTo(XLColor.Red);
        await Assert.That(XLSparklineTheme.Colorful1.NegativeColor).IsNotEqualTo(XLColor.Red);
    }

    [Test]
    public async Task CannotSetSparklineStyleToNull()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1", "B1:Z1");

        Action action = () => group.Style = null!;

        await Assert.That(action).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task SparklinesShiftOnRowInsert()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group1 = ws.SparklineGroups.Add("B2", "D4:F4");
        var group2 = ws.SparklineGroups.Add("B3", "D4:D8");
        var group3 = ws.SparklineGroups.Add("B4", "E1:E8");

        ws.Row(2).InsertRowsBelow(3);

        await Assert.That(group1.First().Location.Address.ToString()).IsEqualTo("B2");
        await Assert.That(group1.First().SourceData.RangeAddress.ToString()).IsEqualTo("D7:F7");
        await Assert.That(group2.First().Location.Address.ToString()).IsEqualTo("B6");
        await Assert.That(group2.First().SourceData.RangeAddress.ToString()).IsEqualTo("D7:D11");
        await Assert.That(group3.First().Location.Address.ToString()).IsEqualTo("B7");
        await Assert.That(group3.First().SourceData.RangeAddress.ToString()).IsEqualTo("E1:E11");
    }

    [Test]
    public async Task SparklinesShiftOnRowDelete()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group1 = ws.SparklineGroups.Add("B2", "D7:F7");
        var group2 = ws.SparklineGroups.Add("B6", "D7:D11");
        var group3 = ws.SparklineGroups.Add("B7", "E1:E11");

        ws.Rows(3, 5).Delete();

        await Assert.That(group1.First().Location.Address.ToString()).IsEqualTo("B2");
        await Assert.That(group1.First().SourceData.RangeAddress.ToString()).IsEqualTo("D4:F4");
        await Assert.That(group2.First().Location.Address.ToString()).IsEqualTo("B3");
        await Assert.That(group2.First().SourceData.RangeAddress.ToString()).IsEqualTo("D4:D8");
        await Assert.That(group3.First().Location.Address.ToString()).IsEqualTo("B4");
        await Assert.That(group3.First().SourceData.RangeAddress.ToString()).IsEqualTo("E1:E8");
    }

    [Test]
    public async Task SparklinesShiftOnColumnInsert()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group1 = ws.SparklineGroups.Add("B2", "D4:F4");
        var group2 = ws.SparklineGroups.Add("C3", "D4:D8");
        var group3 = ws.SparklineGroups.Add("D4", "A4:E4");

        ws.Column(2).InsertColumnsAfter(3);

        await Assert.That(group1.First().Location.Address.ToString()).IsEqualTo("B2");
        await Assert.That(group1.First().SourceData.RangeAddress.ToString()).IsEqualTo("G4:I4");
        await Assert.That(group2.First().Location.Address.ToString()).IsEqualTo("F3");
        await Assert.That(group2.First().SourceData.RangeAddress.ToString()).IsEqualTo("G4:G8");
        await Assert.That(group3.First().Location.Address.ToString()).IsEqualTo("G4");
        await Assert.That(group3.First().SourceData.RangeAddress.ToString()).IsEqualTo("A4:H4");
    }

    [Test]
    public async Task SparklinesShiftOnColumnDelete()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group1 = ws.SparklineGroups.Add("B2", "G4:I4");
        var group2 = ws.SparklineGroups.Add("F3", "G4:G8");
        var group3 = ws.SparklineGroups.Add("G4", "A4:H4");

        ws.Columns(3, 5).Delete();

        await Assert.That(group1.First().Location.Address.ToString()).IsEqualTo("B2");
        await Assert.That(group1.First().SourceData.RangeAddress.ToString()).IsEqualTo("D4:F4");
        await Assert.That(group2.First().Location.Address.ToString()).IsEqualTo("C3");
        await Assert.That(group2.First().SourceData.RangeAddress.ToString()).IsEqualTo("D4:D8");
        await Assert.That(group3.First().Location.Address.ToString()).IsEqualTo("D4");
        await Assert.That(group3.First().SourceData.RangeAddress.ToString()).IsEqualTo("A4:E4");
    }

    [Test]
    public async Task SparklineRemovedWhenColumnDeleted()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1:B1", "C2:D6");

        ws.Column(2).Delete();

        await Assert.That(group.Count()).IsEqualTo(1);
        await Assert.That(group.Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(group.Single().SourceData.RangeAddress.ToString()).IsEqualTo("B2:B6");
    }

    [Test]
    public async Task SparklineRemovedWhenRowDeleted()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1:A2", "C3:F4");

        ws.Row(2).Delete();

        await Assert.That(group.Count()).IsEqualTo(1);
        await Assert.That(group.Single().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(group.Single().SourceData.RangeAddress.ToString()).IsEqualTo("C2:F2");
    }

    [Test]
    public async Task SparklineRemovedWhenShiftedTooFarRight()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("XFD1", "A1:Z1");

        ws.Column(1).InsertColumnsBefore(1);

        await Assert.That(group.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task SparklineRemovedWhenShiftedTooFarDown()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1048576", "A1:Z1");

        ws.Row(1).InsertRowsAbove(1);

        await Assert.That(group.Count()).IsEqualTo(0);
    }

    [Test]
    public async Task SparklineRangeInvalidatedWhenDeleted()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1:B1", "C2:D6");

        ws.Column(4).Delete();

        await Assert.That(group.Count()).IsEqualTo(2);
        await Assert.That(group.First().Location.Address.ToString()).IsEqualTo("A1");
        await Assert.That(group.First().SourceData.RangeAddress.ToString()).IsEqualTo("C2:C6");
        await Assert.That(group.Last().Location.Address.ToString()).IsEqualTo("B1");
        await Assert.That(group.Last().SourceData.RangeAddress.IsValid).IsFalse();
    }

    #endregion Change sparklines

    #region Load and save sparkline groups

    [Test]
    public async Task CanChangeSaveAndLoadSparklineGroup()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet 1");
            var originalGroup = ws.SparklineGroups.Add("A1:A3", "B1:Z3")
                .SetDateRange(ws.Range("B4:Z4"))
                .SetLineWeight(5.5)
                .SetDisplayHidden(true)
                .SetShowMarkers(XLSparklineMarkers.FirstPoint | XLSparklineMarkers.LastPoint |
                                XLSparklineMarkers.HighPoint | XLSparklineMarkers.LowPoint |
                                XLSparklineMarkers.NegativePoints | XLSparklineMarkers.Markers)
                .SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues.Zero)
                .SetType(XLSparklineType.Stacked);

            originalGroup.HorizontalAxis
                .SetColor(XLColor.AirForceBlue)
                .SetVisible(true)
                .SetRightToLeft(true);

            originalGroup.VerticalAxis
                .SetManualMax(6.6)
                .SetManualMin(1.2)
                .SetMaxAxisType(XLSparklineAxisMinMax.Custom)
                .SetMinAxisType(XLSparklineAxisMinMax.Custom);

            originalGroup.Style
                .SetFirstMarkerColor(XLColor.AliceBlue)
                .SetHighMarkerColor(XLColor.Alizarin)
                .SetLastMarkerColor(XLColor.Almond)
                .SetLowMarkerColor(XLColor.Amaranth)
                .SetMarkersColor(XLColor.Amber)
                .SetNegativeColor(XLColor.AmberSaeEce)
                .SetSeriesColor(XLColor.AmericanRose);

            await AssertGroupIsValid(originalGroup);
            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();

            await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
            await AssertGroupIsValid(ws.SparklineGroups.Single());
        }

        async Task AssertGroupIsValid(IXLSparklineGroup group)
        {
            await Assert.That(group.Count()).IsEqualTo(3);

            await Assert.That(group.ElementAt(0).Location.Address.ToString()).IsEqualTo("A1");
            await Assert.That(group.ElementAt(1).Location.Address.ToString()).IsEqualTo("A2");
            await Assert.That(group.ElementAt(2).Location.Address.ToString()).IsEqualTo("A3");

            await Assert.That(group.ElementAt(0).SourceData.RangeAddress.ToString()).IsEqualTo("B1:Z1");
            await Assert.That(group.ElementAt(1).SourceData.RangeAddress.ToString()).IsEqualTo("B2:Z2");
            await Assert.That(group.ElementAt(2).SourceData.RangeAddress.ToString()).IsEqualTo("B3:Z3");

            await Assert.That(group.DateRange!.RangeAddress.ToString()!).IsEqualTo("B4:Z4");

            await Assert.That(group.Style.FirstMarkerColor).IsEqualTo(XLColor.AliceBlue);
            await Assert.That(group.Style.HighMarkerColor).IsEqualTo(XLColor.Alizarin);
            await Assert.That(group.Style.LastMarkerColor).IsEqualTo(XLColor.Almond);
            await Assert.That(group.Style.LowMarkerColor).IsEqualTo(XLColor.Amaranth);
            await Assert.That(group.Style.MarkersColor).IsEqualTo(XLColor.Amber);
            await Assert.That(group.Style.NegativeColor).IsEqualTo(XLColor.AmberSaeEce);
            await Assert.That(group.Style.SeriesColor).IsEqualTo(XLColor.AmericanRose);
            await Assert.That(group.DisplayHidden).IsTrue();
            await Assert.That(group.LineWeight).IsEqualTo(5.5).Within(XLHelper.Epsilon);
            await Assert.That(group.DisplayEmptyCellsAs).IsEqualTo(XLDisplayBlanksAsValues.Zero);
            await Assert.That(group.Type).IsEqualTo(XLSparklineType.Stacked);

            await Assert.That(group.ShowMarkers.HasFlag(XLSparklineMarkers.FirstPoint)).IsTrue();
            await Assert.That(group.ShowMarkers.HasFlag(XLSparklineMarkers.LastPoint)).IsTrue();
            await Assert.That(group.ShowMarkers.HasFlag(XLSparklineMarkers.HighPoint)).IsTrue();
            await Assert.That(group.ShowMarkers.HasFlag(XLSparklineMarkers.LowPoint)).IsTrue();
            await Assert.That(group.ShowMarkers.HasFlag(XLSparklineMarkers.NegativePoints)).IsTrue();
            await Assert.That(group.ShowMarkers.HasFlag(XLSparklineMarkers.Markers)).IsTrue();

            await Assert.That(group.HorizontalAxis.Color).IsEqualTo(XLColor.AirForceBlue);
            await Assert.That(group.HorizontalAxis.IsVisible).IsTrue();
            await Assert.That(group.HorizontalAxis.RightToLeft).IsTrue();
            await Assert.That(group.HorizontalAxis.DateAxis).IsTrue();

            await Assert.That(group.VerticalAxis.ManualMax!.Value).IsEqualTo(6.6).Within(XLHelper.Epsilon);
            await Assert.That(group.VerticalAxis.ManualMin!.Value).IsEqualTo(1.2).Within(XLHelper.Epsilon);
            await Assert.That(group.VerticalAxis.MaxAxisType).IsEqualTo(XLSparklineAxisMinMax.Custom);
            await Assert.That(group.VerticalAxis.MinAxisType).IsEqualTo(XLSparklineAxisMinMax.Custom);
        }
    }

    [Test]
    public async Task CanLoadSparklines()
    {
        using var ms = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Sparklines\SparklineThemes\inputfile.xlsx"));
        using var wb = new XLWorkbook(ms);
        await Assert.That(wb.Worksheets.All(ws => ws.SparklineGroups.Count() == 6)).IsTrue();
    }

    [Test]
    [Arguments("Accent!B1", nameof(XLSparklineTheme.Accent1))]
    [Arguments("Accent!B2", nameof(XLSparklineTheme.Accent2))]
    [Arguments("Accent!B3", nameof(XLSparklineTheme.Accent3))]
    [Arguments("Accent!B4", nameof(XLSparklineTheme.Accent4))]
    [Arguments("Accent!B5", nameof(XLSparklineTheme.Accent5))]
    [Arguments("Accent!B6", nameof(XLSparklineTheme.Accent6))]
    [Arguments("'Accent Darker 25%'!B1", nameof(XLSparklineTheme.Accent1Darker25))]
    [Arguments("'Accent Darker 25%'!B2", nameof(XLSparklineTheme.Accent2Darker25))]
    [Arguments("'Accent Darker 25%'!B3", nameof(XLSparklineTheme.Accent3Darker25))]
    [Arguments("'Accent Darker 25%'!B4", nameof(XLSparklineTheme.Accent4Darker25))]
    [Arguments("'Accent Darker 25%'!B5", nameof(XLSparklineTheme.Accent5Darker25))]
    [Arguments("'Accent Darker 25%'!B6", nameof(XLSparklineTheme.Accent6Darker25))]
    [Arguments("'Accent Darker 50%'!B1", nameof(XLSparklineTheme.Accent1Darker50))]
    [Arguments("'Accent Darker 50%'!B2", nameof(XLSparklineTheme.Accent2Darker50))]
    [Arguments("'Accent Darker 50%'!B3", nameof(XLSparklineTheme.Accent3Darker50))]
    [Arguments("'Accent Darker 50%'!B4", nameof(XLSparklineTheme.Accent4Darker50))]
    [Arguments("'Accent Darker 50%'!B5", nameof(XLSparklineTheme.Accent5Darker50))]
    [Arguments("'Accent Darker 50%'!B6", nameof(XLSparklineTheme.Accent6Darker50))]
    [Arguments("'Accent Lighter 40%'!B1", nameof(XLSparklineTheme.Accent1Lighter40))]
    [Arguments("'Accent Lighter 40%'!B2", nameof(XLSparklineTheme.Accent2Lighter40))]
    [Arguments("'Accent Lighter 40%'!B3", nameof(XLSparklineTheme.Accent3Lighter40))]
    [Arguments("'Accent Lighter 40%'!B4", nameof(XLSparklineTheme.Accent4Lighter40))]
    [Arguments("'Accent Lighter 40%'!B5", nameof(XLSparklineTheme.Accent5Lighter40))]
    [Arguments("'Accent Lighter 40%'!B6", nameof(XLSparklineTheme.Accent6Lighter40))]
    [Arguments("Dark!B1", nameof(XLSparklineTheme.Dark1))]
    [Arguments("Dark!B2", nameof(XLSparklineTheme.Dark2))]
    [Arguments("Dark!B3", nameof(XLSparklineTheme.Dark3))]
    [Arguments("Dark!B4", nameof(XLSparklineTheme.Dark4))]
    [Arguments("Dark!B5", nameof(XLSparklineTheme.Dark5))]
    [Arguments("Dark!B6", nameof(XLSparklineTheme.Dark6))]
    [Arguments("Colorful!B1", nameof(XLSparklineTheme.Colorful1))]
    [Arguments("Colorful!B2", nameof(XLSparklineTheme.Colorful2))]
    [Arguments("Colorful!B3", nameof(XLSparklineTheme.Colorful3))]
    [Arguments("Colorful!B4", nameof(XLSparklineTheme.Colorful4))]
    [Arguments("Colorful!B5", nameof(XLSparklineTheme.Colorful5))]
    [Arguments("Colorful!B6", nameof(XLSparklineTheme.Colorful6))]
    public async Task SparklineThemesAreIdenticalToExcel(string cellAddress, string expectedThemeName)
    {
        using var ms = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Sparklines\SparklineThemes\inputfile.xlsx"));
        using var wb = new XLWorkbook(ms);
        var expectedStyle = GetThemeByName(expectedThemeName);
        var actualStyle = wb.Cell(cellAddress)!.Sparkline!.SparklineGroup.Style;

        await Assert.That(actualStyle).IsEqualTo(expectedStyle);
        return;

        IXLSparklineStyle GetThemeByName(string themeName)
        {
            var themes = typeof(XLSparklineTheme);
            var prop = themes.GetProperty(themeName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
            return prop.GetValue(null, null)! as IXLSparklineStyle;
        }
    }

    [Test]
    public async Task DeletedSparklinesRemovedFromFile()
    {
        using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Sparklines\SparklineThemes\inputfile.xlsx"));
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(input))
        {
            wb.Worksheet(1).SparklineGroups.RemoveAll();
            wb.Worksheet(2).SparklineGroups.Remove(wb.Worksheet(2).Cell("B1"));
            wb.Worksheet(3).SparklineGroups.Remove(wb.Worksheet(3).Range("B2:B6"));
            wb.Worksheet(4).SparklineGroups.Remove(wb.Worksheet(4).SparklineGroups.First());

            wb.SaveAs(output);
        }

        using (var wb = new XLWorkbook(output))
        {
            await Assert.That(wb.Worksheet(1).SparklineGroups.Count()).IsEqualTo(0);
            await Assert.That(wb.Worksheet(2).SparklineGroups.Count()).IsEqualTo(5);
            await Assert.That(wb.Worksheet(3).SparklineGroups.Count()).IsEqualTo(1);
            await Assert.That(wb.Worksheet(4).SparklineGroups.Count()).IsEqualTo(5);
            await Assert.That(wb.Worksheet(5).SparklineGroups.Count()).IsEqualTo(6);
            await Assert.That(wb.Worksheet(6).SparklineGroups.Count()).IsEqualTo(6);
        }
    }

    [Test]
    public async Task EmptySparklineGroupsSkippedOnSaving()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet 1");
            var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");

            group.RemoveAll();

            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            await Assert.That(wb.Worksheets.First().SparklineGroups.Count()).IsEqualTo(0);
        }
    }

    [Test]
    public async Task CanSaveAndLoadSparklineWithInvalidRange()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws1 = wb.AddWorksheet("Sheet 1");
            var ws2 = wb.AddWorksheet("Sheet 2");

            ws1.SparklineGroups.Add("A1:A3", "'Sheet 2'!B1:F3");
            ws1.SparklineGroups.Add("A4:A6", "B4:F6")
                .SetDateRange(ws2.Range("A1:E1"));

            ws2.Delete();
            wb.SaveAs(ms);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.Single();

            await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(2);
            await Assert.That(ws.Cell("A2").Sparkline!.IsValid).IsFalse();
            await Assert.That(ws.Cell("A5").Sparkline!.SourceData.RangeAddress.ToString()).IsEqualTo("B5:F5");
            await Assert.That(ws.Cell("A5").Sparkline!.SparklineGroup.DateRange).IsNull();
        }
    }

    #endregion Load and save sparkline groups

    #region Change sparkline groups

    [Test]
    public async Task SetManualMinChangesAxisTypeToCustom()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var axis = ws.SparklineGroups.Add("A1:A2", "B1:Z2")
            .VerticalAxis
            .SetMinAxisType(XLSparklineAxisMinMax.SameForAll);

        axis.ManualMin = 100;

        await Assert.That(axis.ManualMin.Value).IsEqualTo(100).Within(XLHelper.Epsilon);
        await Assert.That(axis.MinAxisType).IsEqualTo(XLSparklineAxisMinMax.Custom);
    }

    [Test]
    public async Task SetManualMaxChangesAxisTypeToCustom()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var axis = ws.SparklineGroups.Add("A1:A2", "B1:Z2")
            .VerticalAxis
            .SetMaxAxisType(XLSparklineAxisMinMax.SameForAll);

        axis.ManualMax = 100;

        await Assert.That(axis.ManualMax.Value).IsEqualTo(100).Within(XLHelper.Epsilon);
        await Assert.That(axis.MaxAxisType).IsEqualTo(XLSparklineAxisMinMax.Custom);
    }

    [Test]
    [Arguments(XLSparklineAxisMinMax.Custom, 100)]
    [Arguments(XLSparklineAxisMinMax.SameForAll, null)]
    [Arguments(XLSparklineAxisMinMax.Automatic, null)]
    public async Task SetAxisTypeToNonCustomSetsManualMinToNull(XLSparklineAxisMinMax axisType, double? expectedManualMin)
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var axis = ws.SparklineGroups.Add("A1", "B1:Z1")
            .VerticalAxis
            .SetManualMin(100);

        axis.MinAxisType = axisType;

        if (expectedManualMin.HasValue)
            await Assert.That(axis.ManualMin!.Value).IsEqualTo(expectedManualMin.Value).Within(XLHelper.Epsilon);
        else
            await Assert.That(axis.ManualMin).IsNull();
    }

    [Test]
    [Arguments(XLSparklineAxisMinMax.Custom, 100)]
    [Arguments(XLSparklineAxisMinMax.SameForAll, null)]
    [Arguments(XLSparklineAxisMinMax.Automatic, null)]
    public async Task SetAxisTypeToNonCustomSetsManualMaxToNull(XLSparklineAxisMinMax axisType, double? expectedManualMax)
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var axis = ws.SparklineGroups.Add("A1", "B1:Z1")
            .VerticalAxis
            .SetManualMax(100);

        axis.MaxAxisType = axisType;

        if (expectedManualMax.HasValue)
            await Assert.That(axis.ManualMax!.Value).IsEqualTo(expectedManualMax.Value).Within(XLHelper.Epsilon);
        else
            await Assert.That(axis.ManualMax).IsNull();
    }

    [Test]
    public async Task SetDateRangeChangesAxisType()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");

        group.DateRange = ws.Range("B3:Z3");

        await Assert.That(group.HorizontalAxis.DateAxis).IsTrue();
    }

    [Test]
    public async Task SetDateRangeToNullChangesAxisType()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");
        group.DateRange = ws.Range("B3:Z3");

        group.DateRange = null;

        await Assert.That(group.HorizontalAxis.DateAxis).IsFalse();
    }

    [Test]
    public async Task CannotSetNonLinearDateRange()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        var group = ws.SparklineGroups.Add("A1:A2", "B1:Z2");

        Action action = () => group.DateRange = ws.Range("B3:Z4");

        await Assert.That(action).Throws<ArgumentException>();
    }

    #endregion Change sparkline groups

    #region Copy sparkline groups

    [Test]
    public async Task CopyCellToSameWorksheetCopiesSparkline()
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet 1");
        ws.SparklineGroups.Add("A1:A3", "B1:F3");
        var target = ws.Cell("D4");

        ws.Cell("A2").CopyTo(target);

        await Assert.That(ws.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(target.HasSparkline).IsTrue();
        await Assert.That(target.Sparkline!.SparklineGroup).IsSameReferenceAs(ws.Cell("A2").Sparkline!.SparklineGroup);
        await Assert.That(target.Sparkline.SourceData.RangeAddress.ToString()).IsEqualTo("E4:I4");
    }

    [Test]
    public async Task CopyCellToDifferentWorksheetCopiesSparklineGroup()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws2 = wb.AddWorksheet("Sheet 2");
        var ws3 = wb.AddWorksheet("Sheet 3");
        ws1.SparklineGroups.Add("A1:A3", "B1:F3");
        ws1.SparklineGroups.Add("A4:A6", "'Sheet 3'!B4:F6");
        var target1 = ws2.Cell("D4");
        var target2 = ws2.Cell("D5");

        ws1.Cell("A2").CopyTo(target1);
        ws1.Cell("A5").CopyTo(target2);

        await Assert.That(ws1.SparklineGroups.Count()).IsEqualTo(2);
        await Assert.That(ws2.SparklineGroups.Count()).IsEqualTo(2);
        await Assert.That(target1.HasSparkline).IsTrue();
        await Assert.That(target2.HasSparkline).IsTrue();
        await Assert.That(target1.Sparkline!.SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 2'!E4:I4");
        await Assert.That(target2.Sparkline!.SourceData.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 3'!E5:I5");
    }

    [Test]
    public async Task CopySparklineIfDateRangeOnSameWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws2 = wb.AddWorksheet("Sheet 2");
        var group = ws1.SparklineGroups.Add("A1:A3", "B1:F3");
        group.SetDateRange(ws1.Range("A4:E4"));
        var target = ws2.Cell("D4");

        ws1.Cell("A2").CopyTo(target);

        await Assert.That(ws1.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws2.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(target.HasSparkline).IsTrue();
        await Assert.That(target.Sparkline!.SparklineGroup.DateRange!.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 2'!D6:H6");
    }

    [Test]
    public async Task CopySparklineIfDateRangeSourceOnDifferentWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet 1");
        var ws2 = wb.AddWorksheet("Sheet 2");
        var ws3 = wb.AddWorksheet("Sheet 3");
        var group = ws1.SparklineGroups.Add("A1:A3", "B1:F3");
        group.SetDateRange(ws3.Range("A4:E4"));
        var target = ws2.Cell("D4");

        ws1.Cell("A2").CopyTo(target);

        await Assert.That(ws1.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(ws2.SparklineGroups.Count()).IsEqualTo(1);
        await Assert.That(target.HasSparkline).IsTrue();
        await Assert.That(target.Sparkline!.SparklineGroup.DateRange!.RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("'Sheet 3'!D6:H6");
    }

    #endregion Copy sparkline groups

    #region Test Examples

    [Test]
    public async Task CreateSampleSparklines()
    {
        await TestHelper.RunTestExample<SampleSparklines>(@"Sparklines\SampleSparklines.xlsx");
    }

    #endregion Test Examples
}
