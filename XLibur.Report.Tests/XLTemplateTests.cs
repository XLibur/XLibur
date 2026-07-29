using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests;

public class XLTemplateTests
{
    private static IXLWorkbook WorkbookWith(params (string Address, string Text)[] cells)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        foreach (var (address, text) in cells)
        {
            sheet.Cell(address).Value = text;
        }

        return workbook;
    }

    private static SaleItem SampleItem() => new()
    {
        Product = "Widget",
        Quantity = 3,
        UnitPrice = 9.5m,
        SoldOn = new DateTime(2026, 3, 14),
    };

    [Test]
    public async Task WholeCellExpressionKeepsItsType()
    {
        using var workbook = WorkbookWith(("A1", "{{ item.UnitPrice }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("item", SampleItem());

        template.Generate();

        var cell = workbook.Worksheet("Report").Cell("A1");
        await Assert.That(cell.Value.IsNumber).IsTrue();
        await Assert.That(cell.Value.GetNumber()).IsEqualTo(9.5);
    }

    [Test]
    public async Task WholeCellDateExpressionStaysADate()
    {
        using var workbook = WorkbookWith(("A1", "{{ item.SoldOn }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("item", SampleItem());

        template.Generate();

        var cell = workbook.Worksheet("Report").Cell("A1");
        await Assert.That(cell.Value.IsDateTime).IsTrue();
        await Assert.That(cell.Value.GetDateTime()).IsEqualTo(new DateTime(2026, 3, 14));
    }

    [Test]
    public async Task MixedCellBecomesText()
    {
        using var workbook = WorkbookWith(("A1", "Sold {{ item.Quantity }} units"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("item", SampleItem());

        template.Generate();

        var cell = workbook.Worksheet("Report").Cell("A1");
        await Assert.That(cell.Value.IsText).IsTrue();
        await Assert.That(cell.Value.GetText()).IsEqualTo("Sold 3 units");
    }

    [Test]
    public async Task PlainCellsAreLeftAlone()
    {
        using var workbook = WorkbookWith(("A1", "Heading"));
        using var template = new XLTemplate(workbook);

        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetText()).IsEqualTo("Heading");
    }

    [Test]
    public async Task NumericCellsAreLeftAlone()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = 42;
        using var template = new XLTemplate(workbook);

        template.Generate();

        await Assert.That(sheet.Cell("A1").Value.GetNumber()).IsEqualTo(42d);
    }

    [Test]
    public async Task ExistingFormulasAreLeftAlone()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = 2;
        sheet.Cell("A2").FormulaA1 = "A1*2";
        using var template = new XLTemplate(workbook);

        template.Generate();

        await Assert.That(sheet.Cell("A2").FormulaA1).IsEqualTo("A1*2");
    }

    [Test]
    public async Task FormulaPrefixBuildsAFormula()
    {
        using var workbook = WorkbookWith(("A1", "&=SUM(B1:B{{ rows }})"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("rows", 10);

        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A1").FormulaA1).IsEqualTo("SUM(B1:B10)");
    }

    [Test]
    public async Task AddVariableReflectsPublicMembers()
    {
        using var workbook = WorkbookWith(("A1", "{{ Company }}"), ("A2", "{{ Year }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable(new ReportHeader { Company = "Contoso", RunOn = new DateTime(2026, 1, 1) });

        template.Generate();

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Contoso");
        await Assert.That(sheet.Cell("A2").Value.GetNumber()).IsEqualTo(2026d);
    }

    [Test]
    public async Task AddVariableExplodesADictionary()
    {
        using var workbook = WorkbookWith(("A1", "{{ City }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable(new Dictionary<string, object?> { ["City"] = "Perth" });

        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetText()).IsEqualTo("Perth");
    }

    [Test]
    public async Task HiddenSheetsAreNotGenerated()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Hidden");
        sheet.Cell("A1").Value = "{{ value }}";
        sheet.Visibility = XLWorksheetVisibility.Hidden;
        using var template = new XLTemplate(workbook);
        template.AddVariable("value", "generated");

        template.Generate();

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("{{ value }}");
    }

    /// <summary>
    /// One malformed expression must cost that cell and nothing else — upstream ClosedXML.Report
    /// issue #340 is the counter-example, where a single bad formula failed the whole report.
    /// </summary>
    [Test]
    public async Task BadExpressionIsRecordedWithoutAbortingGeneration()
    {
        using var workbook = WorkbookWith(("A1", "{{ item. }}"), ("A2", "{{ item.Product }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("item", SampleItem());

        var result = template.Generate();

        await Assert.That(result.HasErrors).IsTrue();
        await Assert.That(result.ParsingErrors.Count).IsEqualTo(1);
        await Assert.That(result.ParsingErrors[0].CellAddress).IsEqualTo("A1");
        await Assert.That(result.ParsingErrors[0].SheetName).IsEqualTo("Report");
        await Assert.That(workbook.Worksheet("Report").Cell("A2").Value.GetText()).IsEqualTo("Widget");
    }

    [Test]
    public async Task BadExpressionLeavesTheErrorVisibleInTheCell()
    {
        using var workbook = WorkbookWith(("A1", "{{ item. }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("item", SampleItem());

        template.Generate();

        var cell = workbook.Worksheet("Report").Cell("A1");
        await Assert.That(cell.Value.IsText).IsTrue();
        await Assert.That(cell.Value.GetText()).IsNotEqualTo("{{ item. }}");
        await Assert.That(cell.Style.Font.FontColor).IsEqualTo(XLColor.Red);
    }

    [Test]
    public async Task CleanRunReportsNoErrors()
    {
        using var workbook = WorkbookWith(("A1", "{{ item.Product }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("item", SampleItem());

        var result = template.Generate();

        await Assert.That(result.HasErrors).IsFalse();
        await Assert.That(result.ParsingErrors.Count).IsEqualTo(0);
    }

    [Test]
    public async Task GenerateTwiceIsRejected()
    {
        using var workbook = WorkbookWith(("A1", "{{ value }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("value", 1);
        template.Generate();

        await Assert.That(() => template.Generate()).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task IsGeneratedTracksGeneration()
    {
        using var workbook = WorkbookWith(("A1", "{{ value }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("value", 1);

        await Assert.That(template.IsGenerated).IsFalse();
        template.Generate();
        await Assert.That(template.IsGenerated).IsTrue();
    }

    [Test]
    public async Task GeneratedWorkbookRoundTripsThroughAStream()
    {
        using var stream = new MemoryStream();

        using (var workbook = WorkbookWith(("A1", "{{ item.Product }}"), ("A2", "{{ item.Total }}")))
        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("item", SampleItem());
            template.Generate();
            template.SaveAs(stream);
        }

        stream.Position = 0;
        using var reloaded = new XLWorkbook(stream);
        var sheet = reloaded.Worksheet("Report");

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A2").Value.IsNumber).IsTrue();
        await Assert.That(sheet.Cell("A2").Value.GetNumber()).IsEqualTo(28.5);
    }

    [Test]
    public async Task TemplateGivenAWorkbookDoesNotDisposeIt()
    {
        using var workbook = WorkbookWith(("A1", "{{ value }}"));

        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("value", "kept");
            template.Generate();
        }

        // Still usable after the template was disposed.
        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetText()).IsEqualTo("kept");
    }

    [Test]
    public async Task UsingADisposedTemplateIsRejected()
    {
        using var workbook = WorkbookWith(("A1", "{{ value }}"));
        var template = new XLTemplate(workbook);
        template.Dispose();

        await Assert.That(() => template.AddVariable("value", 1)).Throws<ObjectDisposedException>();
    }

    [Test]
    public async Task EveryVisibleSheetIsGenerated()
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("First").Cell("A1").Value = "{{ value }}";
        workbook.AddWorksheet("Second").Cell("A1").Value = "{{ value }}";
        using var template = new XLTemplate(workbook);
        template.AddVariable("value", "done");

        template.Generate();

        await Assert.That(workbook.Worksheet("First").Cell("A1").Value.GetText()).IsEqualTo("done");
        await Assert.That(workbook.Worksheet("Second").Cell("A1").Value.GetText()).IsEqualTo("done");
    }

    [Test]
    public async Task DataTableBindsAsItsRows()
    {
        using var table = new System.Data.DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("first");
        table.Rows.Add("second");

        using var workbook = WorkbookWith(("A1", "{{ rows.size }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("rows", table);

        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetNumber()).IsEqualTo(2d);
    }

    /// <summary>
    /// A DataRow exposes its columns through an indexer, not through properties, so binding rows
    /// as-is would leave <c>{{ row.Name }}</c> — the natural way to write it — resolving to
    /// nothing. Columns are bound by name instead.
    /// </summary>
    [Test]
    public async Task DataTableColumnsAreAddressableByName()
    {
        using var table = new System.Data.DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add("first", 12.5m);

        using var workbook = WorkbookWith(("A1", "{{ rows[0].Name }}"), ("A2", "{{ rows[0].Amount }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("rows", table);

        template.Generate();

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("first");
        await Assert.That(sheet.Cell("A2").Value.IsNumber).IsTrue();
        await Assert.That(sheet.Cell("A2").Value.GetNumber()).IsEqualTo(12.5);
    }

    [Test]
    public async Task DataTableNullsBecomeBlanks()
    {
        using var table = new System.Data.DataTable();
        table.Columns.Add("Notes", typeof(string));
        table.Rows.Add(DBNull.Value);

        using var workbook = WorkbookWith(("A1", "{{ rows[0].Notes }}"));
        using var template = new XLTemplate(workbook);
        template.AddVariable("rows", table);

        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.IsBlank).IsTrue();
    }
}
