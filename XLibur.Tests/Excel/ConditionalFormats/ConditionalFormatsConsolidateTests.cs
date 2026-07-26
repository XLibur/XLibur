using System.Collections.Generic;
using XLibur.Excel;
using System.Linq;
using XLibur.Excel.ConditionalFormats;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.ConditionalFormats;

public class ConditionalFormatsConsolidateTests
{
    [Test]
    public async Task ConsecutivelyRowsConsolidateTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B2:C2").AddConditionalFormat());
        SetFormat1(ws.Range("B4:C4").AddConditionalFormat());
        SetFormat1(ws.Range("B3:C3").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        var format = ws.ConditionalFormats.First();
        await Assert.That(format.Range.RangeAddress.ToStringRelative()).IsEqualTo("B2:C4");
        await Assert.That(format.Values.Values.First().Value).IsEqualTo("F2");
    }

    [Test]
    public async Task ConsecutivelyColumnsConsolidateTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("D2:D3").AddConditionalFormat());
        SetFormat1(ws.Range("B2:B3").AddConditionalFormat());
        SetFormat1(ws.Range("C2:C3").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        var format = ws.ConditionalFormats.First();
        await Assert.That(format.Ranges.First().RangeAddress.ToStringRelative()).IsEqualTo("B2:D3");
        await Assert.That(format.Values.Values.First().Value).IsEqualTo("F2");
    }

    [Test]
    public async Task Contains1ConsolidateTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
        SetFormat1(ws.Range("C12:D12").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        var format = ws.ConditionalFormats.First();
        await Assert.That(format.Range.RangeAddress.ToStringRelative()).IsEqualTo("B11:D12");
        await Assert.That(format.Values.Values.First().Value).IsEqualTo("F11");
    }

    [Test]
    public async Task Contains2ConsolidateTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B14:C14").AddConditionalFormat());
        SetFormat1(ws.Range("B14:B14").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        var format = ws.ConditionalFormats.First();
        await Assert.That(format.Range.RangeAddress.ToStringRelative()).IsEqualTo("B14:C14");
        await Assert.That(format.Values.Values.First().Value).IsEqualTo("F14");
    }

    [Test]
    public async Task SuperimposedConsolidateTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B16:D18").AddConditionalFormat());
        SetFormat1(ws.Range("B18:D19").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        var format = ws.ConditionalFormats.First();
        await Assert.That(format.Range.RangeAddress.ToStringRelative()).IsEqualTo("B16:D19");
        await Assert.That(format.Values.Values.First().Value).IsEqualTo("F16");
    }

    [Test]
    public async Task DifferentFormatNoConsolidateTest()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("B11:D12").AddConditionalFormat());
        SetFormat2(ws.Range("C12:D12").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task ConsolidatePreservesPriorities()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("A1:A5").AddConditionalFormat());
        SetFormat2(ws.Range("A1:A5").AddConditionalFormat());
        SetFormat2(ws.Range("A6:A10").AddConditionalFormat());
        SetFormat1(ws.Range("A6:A10").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(3);
        await Assert.That((ws.ConditionalFormats.Last().Style as XLStyle).Value).IsEqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value);
        await Assert.That((ws.ConditionalFormats.ElementAt(1).Style as XLStyle).Value).IsNotEqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value);
    }

    [Test]
    public async Task ConsolidatePreservesPriorities2()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        SetFormat1(ws.Range("A1:A1").AddConditionalFormat());
        SetFormat2(ws.Range("A2:A3").AddConditionalFormat());
        SetFormat1(ws.Range("A2:A6").AddConditionalFormat());
        SetFormat1(ws.Range("A7:A8").AddConditionalFormat());

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(3);
        await Assert.That((ws.ConditionalFormats.Last().Style as XLStyle).Value).IsEqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value);
        await Assert.That((ws.ConditionalFormats.ElementAt(1).Style as XLStyle).Value).IsNotEqualTo((ws.ConditionalFormats.First().Style as XLStyle).Value);
        await Assert.That(ws.ConditionalFormats.All(cf => cf.Ranges.Count == 1)).IsTrue().Because("Number of ranges in consolidated conditional formats is expected to be 1");
        await Assert.That(ws.ConditionalFormats.ElementAt(0).Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(ws.ConditionalFormats.ElementAt(1).Ranges.Single().RangeAddress.ToString()).IsEqualTo("A2:A3");
        await Assert.That(ws.ConditionalFormats.ElementAt(2).Ranges.Single().RangeAddress.ToString()).IsEqualTo("A2:A8");
    }

    [Test]
    public async Task ConsolidateShiftsFormulaRelativelyToTopMostCell()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8").Cast<XLRange>();
        var cf = new XLConditionalFormat(ranges);
        cf.Values.Add(new XLFormula("=A3=$D3"));
        cf.Style.Fill.SetBackgroundColor(XLColor.Red);
        ws.ConditionalFormats.Add(cf);

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That((cf.Style as XLStyle).Value).IsEqualTo((ws.ConditionalFormats.Single().Style as XLStyle).Value);
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString()).IsEqualTo("A3:C8");
        await Assert.That(ws.ConditionalFormats.Single().Values.Single().Value.IsFormula).IsTrue();
        await Assert.That(ws.ConditionalFormats.Single().Values.Single().Value.Value).IsEqualTo("A3=$D3");
    }

    [Test]
    public async Task ColorScaleComparing()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        var ranges = ws.Ranges("B3:B8,C3:C4,A3:A4,C5:C8,A5:A8").Cast<XLRange>();
        var cf1 = new XLConditionalFormat(ranges);
        cf1.ColorScale()
            .LowestValue(XLColor.Red)
            .HighestValue(XLColor.Green);

        var cf2 = new XLConditionalFormat(ranges);
        cf2.ColorScale()
            .LowestValue(XLColor.Red)
            .HighestValue(XLColor.Green);
        await Assert.That(XLConditionalFormat.NoRangeComparer.Equals(cf1, cf2)).IsTrue();
    }

    [Test]
    public async Task EqualFormats_have_same_hash_via_NoRangeComparer()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        var cf1 = new XLConditionalFormat((XLRange)ws.Range("A1:B2"));
        cf1.ColorScale()
            .LowestValue(XLColor.Red)
            .HighestValue(XLColor.Green);

        var cf2 = new XLConditionalFormat((XLRange)ws.Range("C3:D4"));
        cf2.ColorScale()
            .LowestValue(XLColor.Red)
            .HighestValue(XLColor.Green);

        // Equal formats (ignoring range) must produce the same hash
        var comparer = XLConditionalFormat.NoRangeComparer;
        await Assert.That(comparer.GetHashCode(cf1)).IsEqualTo(comparer.GetHashCode(cf2));

        // HashSet dedup should treat them as one entry
        var set = new HashSet<IXLConditionalFormat>(comparer) { cf1, cf2 };
        await Assert.That(set).Count().IsEqualTo(1);
    }

    [Test]
    public async Task DifferentFormats_not_equal_via_FullComparer()
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        var cf1 = new XLConditionalFormat((XLRange)ws.Range("A1:B2"));
        cf1.WhenEquals(5).Fill.SetBackgroundColor(XLColor.Blue);

        var cf2 = new XLConditionalFormat((XLRange)ws.Range("A1:B2"));
        cf2.WhenEquals(10).Fill.SetBackgroundColor(XLColor.Red);

        var comparer = XLConditionalFormat.FullComparer;
        await Assert.That(comparer.Equals(cf1, cf2)).IsFalse();

        // HashSet uses comparer.Equals to resolve collisions, so both are kept because Equals returns false
        var set = new HashSet<IXLConditionalFormat>(comparer) { cf1, cf2 };
        await Assert.That(set).Count().IsEqualTo(2);
    }

    private static void SetFormat1(IXLConditionalFormat format)
    {
        format.WhenEquals("=" + format.Range.FirstCell().CellRight(4).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
    }

    private static void SetFormat2(IXLConditionalFormat format)
    {
        format.WhenEquals(5).Fill.SetBackgroundColor(XLColor.AliceBlue);
    }
}
