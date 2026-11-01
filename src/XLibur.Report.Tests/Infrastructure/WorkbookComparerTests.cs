using System;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Infrastructure;

public class WorkbookComparerTests
{
    private static XLWorkbook Sheet(Action<IXLWorksheet> build, string name = "Report")
    {
        var workbook = new XLWorkbook();
        build(workbook.AddWorksheet(name));
        return workbook;
    }

    [Test]
    public async Task IdenticalWorkbooksHaveNoDifferences()
    {
        using var expected = Sheet(ws => ws.Cell("A1").Value = "same");
        using var actual = Sheet(ws => ws.Cell("A1").Value = "same");

        await Assert.That(WorkbookComparer.Compare(expected, actual)).IsEmpty();
    }

    [Test]
    public async Task EmptyWorkbooksHaveNoDifferences()
    {
        using var expected = Sheet(_ => { });
        using var actual = Sheet(_ => { });

        await Assert.That(WorkbookComparer.Compare(expected, actual)).IsEmpty();
    }

    [Test]
    public async Task DifferentSheetNamesAreReported()
    {
        using var expected = Sheet(_ => { }, "Report");
        using var actual = Sheet(_ => { }, "Other");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Count).IsEqualTo(1);
        await Assert.That(differences[0]).Contains("Worksheets");
    }

    [Test]
    public async Task DifferentCellTextIsReported()
    {
        using var expected = Sheet(ws => ws.Cell("B2").Value = "left");
        using var actual = Sheet(ws => ws.Cell("B2").Value = "right");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("Report!B2") && d.Contains("value"))).IsTrue();
    }

    /// <summary>
    /// The comparison has to distinguish a number from its text form, because keeping expression
    /// results typed is the behaviour most worth protecting.
    /// </summary>
    [Test]
    public async Task NumberIsNotEqualToItsTextForm()
    {
        using var expected = Sheet(ws => ws.Cell("A1").Value = 9.5);
        using var actual = Sheet(ws => ws.Cell("A1").Value = "9.5");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("Number:9.5") && d.Contains("Text:9.5"))).IsTrue();
    }

    [Test]
    public async Task MissingCellIsReported()
    {
        using var expected = Sheet(ws => ws.Cell("A1").Value = "present");
        using var actual = Sheet(_ => { });

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("Report!A1") && d.Contains("Blank"))).IsTrue();
    }

    [Test]
    public async Task DifferentFormulaIsReported()
    {
        using var expected = Sheet(ws => ws.Cell("A1").FormulaA1 = "SUM(B1:B5)");
        using var actual = Sheet(ws => ws.Cell("A1").FormulaA1 = "SUM(B1:B9)");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("formula"))).IsTrue();
    }

    [Test]
    public async Task DifferentStyleIsReported()
    {
        using var expected = Sheet(ws => ws.Cell("A1").Style.Font.Bold = true);
        using var actual = Sheet(ws => ws.Cell("A1").Style.Font.Bold = false);

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("style differs"))).IsTrue();
    }

    [Test]
    public async Task StyleDifferenceCanBeIgnored()
    {
        using var expected = Sheet(ws => ws.Cell("A1").Style.Font.Bold = true);
        using var actual = Sheet(ws => ws.Cell("A1").Style.Font.Bold = false);

        var differences = WorkbookComparer.Compare(expected, actual, new WorkbookComparisonOptions { Styles = false });

        await Assert.That(differences).IsEmpty();
    }

    [Test]
    public async Task DifferentMergedRangeIsReported()
    {
        using var expected = Sheet(ws => ws.Range("A1:B1").Merge());
        using var actual = Sheet(ws => ws.Range("A1:C1").Merge());

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("merged ranges"))).IsTrue();
    }

    /// <summary>
    /// Rule count is asserted separately from rule ranges: duplicating a rule per generated cell
    /// leaves every range correct while multiplying the count, which is exactly the upstream
    /// behaviour (#216) this library does not reproduce.
    /// </summary>
    [Test]
    public async Task ExtraConditionalFormatIsReported()
    {
        using var expected = Sheet(ws => ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(10).Fill.SetBackgroundColor(XLColor.Red));
        using var actual = Sheet(ws =>
        {
            ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(10).Fill.SetBackgroundColor(XLColor.Red);
            ws.Range("A1:A5").AddConditionalFormat().WhenLessThan(2).Fill.SetBackgroundColor(XLColor.Blue);
        });

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("conditional format count"))).IsTrue();
    }

    [Test]
    public async Task DifferentConditionalFormatRangeIsReported()
    {
        using var expected = Sheet(ws => ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(10).Fill.SetBackgroundColor(XLColor.Red));
        using var actual = Sheet(ws => ws.Range("A1:A9").AddConditionalFormat().WhenGreaterThan(10).Fill.SetBackgroundColor(XLColor.Red));

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("conditional formats"))).IsTrue();
    }

    [Test]
    public async Task DifferentCommentIsReported()
    {
        using var expected = Sheet(ws => ws.Cell("A1").GetComment().AddText("hello"));
        using var actual = Sheet(ws => ws.Cell("A1").GetComment().AddText("goodbye"));

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("comment"))).IsTrue();
    }

    [Test]
    public async Task DifferentHyperlinkIsReported()
    {
        using var expected = Sheet(ws => ws.Cell("A1").SetHyperlink(new XLHyperlink("https://example.com/a")));
        using var actual = Sheet(ws => ws.Cell("A1").SetHyperlink(new XLHyperlink("https://example.com/b")));

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("hyperlink"))).IsTrue();
    }

    [Test]
    public async Task DifferentRowHeightIsReported()
    {
        using var expected = Sheet(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Row(1).Height = 30;
        });
        using var actual = Sheet(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Row(1).Height = 15;
        });

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("height"))).IsTrue();
    }

    [Test]
    public async Task DifferentOutlineLevelIsReported()
    {
        using var expected = Sheet(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Row(1).OutlineLevel = 1;
        });
        using var actual = Sheet(ws => ws.Cell("A1").Value = "x");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("outline level"))).IsTrue();
    }

    [Test]
    public async Task DifferentPageBreakIsReported()
    {
        using var expected = Sheet(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Row(1).AddHorizontalPageBreak();
        });
        using var actual = Sheet(ws => ws.Cell("A1").Value = "x");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("page breaks"))).IsTrue();
    }

    [Test]
    public async Task DifferentAutoFilterIsReported()
    {
        using var expected = Sheet(ws =>
        {
            ws.Cell("A1").Value = "x";
            ws.Range("A1:A1").SetAutoFilter();
        });
        using var actual = Sheet(ws => ws.Cell("A1").Value = "x");

        var differences = WorkbookComparer.Compare(expected, actual);

        await Assert.That(differences.Any(d => d.Contains("autofilter"))).IsTrue();
    }

    [Test]
    public async Task ComparisonStopsAfterTheDifferenceLimit()
    {
        using var expected = Sheet(ws =>
        {
            for (var row = 1; row <= 40; row++)
            {
                ws.Cell(row, 1).Value = "expected";
            }
        });
        using var actual = Sheet(ws =>
        {
            for (var row = 1; row <= 40; row++)
            {
                ws.Cell(row, 1).Value = "actual";
            }
        });

        var differences = WorkbookComparer.Compare(expected, actual, new WorkbookComparisonOptions { MaxDifferences = 5 });

        await Assert.That(differences.Count).IsLessThanOrEqualTo(6);
        await Assert.That(differences[^1]).Contains("stopped after");
    }

    [Test]
    public async Task WorkbooksThatRoundTripThroughAFileStillMatch()
    {
        using var expected = Sheet(ws =>
        {
            ws.Cell("A1").Value = "Widget";
            ws.Cell("A2").Value = 9.5;
            ws.Cell("A3").Value = new DateTime(2026, 3, 14);
            ws.Cell("A4").FormulaA1 = "A2*2";
            ws.Range("B1:C1").Merge();
            ws.Cell("A1").Style.Font.Bold = true;
        });

        using var stream = new System.IO.MemoryStream();
        expected.SaveAs(stream);
        stream.Position = 0;
        using var reloaded = new XLWorkbook(stream);

        await Assert.That(WorkbookComparer.Compare(expected, reloaded)).IsEmpty();
    }
}
