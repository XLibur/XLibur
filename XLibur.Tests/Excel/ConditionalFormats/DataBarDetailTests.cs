using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;

namespace XLibur.Tests.Excel.ConditionalFormats;

/// <summary>
/// The detail attributes of a data bar (spec 48): equality and hashing agree on them.
/// </summary>
public class DataBarDetailTests
{
    private static XLConditionalFormat NewDataBar(IXLWorksheet ws)
    {
        var cf = ws.Range("A1:A2").AddConditionalFormat();
        cf.DataBar(XLColor.Red, XLColor.Blue).LowestValue().HighestValue();
        return (XLConditionalFormat)cf;
    }

    [Test]
    public async Task DataBarFieldsTakePartInTheHash()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        var comparer = XLConditionalFormat.FullComparer;
        var baseline = comparer.GetHashCode(NewDataBar(ws));

        var changes = new System.Action<XLConditionalFormat>[]
        {
            cf => cf.BarBorder = true,
            cf => cf.BarBorderColor = XLColor.Green,
            cf => cf.NegativeBarBorderColor = XLColor.Green,
            cf => cf.NegativeBarColorSameAsPositive = true,
            cf => cf.NegativeBarBorderColorSameAsPositive = false,
            cf => cf.BarMinLength = 7,
            cf => cf.BarMaxLength = 77,
        };

        foreach (var change in changes)
        {
            var other = NewDataBar(ws);
            change(other);
            await Assert.That(comparer.GetHashCode(other)).IsNotEqualTo(baseline);
            await Assert.That(comparer.Equals(other, NewDataBar(ws))).IsFalse();
        }
    }
}
