using NUnit.Framework;
using System;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Which of a plot area's chart groups carries the primary axis pair, and what a loaded chart refuses
/// to do. XLibur's writer always emits the primary group first, but nothing in the schema requires it,
/// so the reader must not decide by document order alone.
/// </summary>
[TestFixture]
public class ChartPrimaryAxisReaderTests
{
    /// <summary>
    /// A plot area holding two <c>c:barChart</c> groups: the one on the secondary axis pair (4/5, the
    /// value axis crossing at the maximum) is written <em>first</em>, ahead of the primary pair (1/2).
    /// </summary>
    private const string SecondaryGroupFirst = """
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="1"/>
            <c:order val="1"/>
            <c:tx><c:v>Right</c:v></c:tx>
            <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
            <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
          </c:ser>
          <c:axId val="4"/>
          <c:axId val="5"/>
        </c:barChart>
        <c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:ser>
            <c:idx val="0"/>
            <c:order val="0"/>
            <c:tx><c:v>Left</c:v></c:tx>
            <c:cat><c:strRef><c:f>Data!$A$1:$A$2</c:f></c:strRef></c:cat>
            <c:val><c:numRef><c:f>Data!$B$1:$B$2</c:f></c:numRef></c:val>
          </c:ser>
          <c:axId val="1"/>
          <c:axId val="2"/>
        </c:barChart>
        <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
        <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/><c:majorUnit val="10"/></c:valAx>
        <c:catAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="5"/></c:catAx>
        <c:valAx><c:axId val="5"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="4"/><c:crosses val="max"/><c:majorUnit val="25"/></c:valAx>
        """;

    [Test]
    public void SecondaryAxisBindingDoesNotDependOnGroupOrder()
    {
        var chart = ChartPartFixture.LoadChart(SecondaryGroupFirst, out var wb);
        using (wb)
        {
            var series = chart.Series.ToList();
            Assert.That(series.Select(s => s.Name), Is.EqualTo(new[] { "Right", "Left" }));

            // The group written first hangs off the axis that crosses at the maximum, so it is the
            // secondary one however early it appears.
            Assert.That(series[0].UseSecondaryAxis, Is.True);
            Assert.That(series[1].UseSecondaryAxis, Is.False);
        }
    }

    [Test]
    public void TheAxisModelsFollowTheSameBinding()
    {
        var chart = ChartPartFixture.LoadChart(SecondaryGroupFirst, out var wb);
        using (wb)
        {
            Assert.That(chart.ValueAxis.MajorUnit, Is.EqualTo(10),
                "The left-hand axis of the primary group is IXLChart.ValueAxis.");
            Assert.That(chart.SecondaryValueAxis.MajorUnit, Is.EqualTo(25));
        }
    }

    [Test]
    public void AddingASeriesToALoadedChartThrows()
    {
        var chart = ChartPartFixture.LoadChart(
            $"<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>{ChartPartFixture.CategoryAndValueSeries("Sales")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>{ChartPartFixture.CategoryAndValueAxes}",
            out var wb);
        using (wb)
        {
            // A new series has nowhere to go: the chart part is patched, never regenerated. Silently
            // dropping it on save would be worse than saying so.
            Assert.Throws<NotSupportedException>(
                () => chart.Series.Add("New", "Data!$B$1:$B$2", "Data!$A$1:$A$2"));
            Assert.Throws<NotSupportedException>(
                () => chart.SecondarySeries.Add("New", "Data!$B$1:$B$2", "Data!$A$1:$A$2"));
            Assert.That(chart.Series, Has.Exactly(1).Items);
        }
    }
}
