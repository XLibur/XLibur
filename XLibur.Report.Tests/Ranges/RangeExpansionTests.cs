using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Ranges;

public class RangeExpansionTests
{
    private static List<SaleItem> Items(params (string Product, int Quantity, decimal UnitPrice)[] rows) =>
        rows.Select(r => new SaleItem
        {
            Product = r.Product,
            Quantity = r.Quantity,
            UnitPrice = r.UnitPrice,
            SoldOn = new DateTime(2026, 1, 1),
        }).ToList();

    private static List<SaleItem> ThreeItems() => Items(("Widget", 2, 5m), ("Gadget", 3, 10m), ("Doohickey", 1, 7.5m));

    /// <summary>
    /// Builds a sheet whose rows <paramref name="dataRows"/> are named <c>Items</c>, with the row
    /// after them acting as the options row.
    /// </summary>
    private static IXLWorkbook TemplateWithItemsRange(
        Action<IXLWorksheet> build,
        int firstRow = 1,
        int dataRows = 1,
        int columns = 3,
        string name = "Items")
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        build(sheet);
        workbook.DefinedNames.Add(name, sheet.Range(firstRow, 1, firstRow + dataRows, columns));
        return workbook;
    }

    private static XLGenerateResult Generate(IXLWorkbook workbook, string variableName, object? value)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable(variableName, value);
        return template.Generate();
    }

    [Test]
    public async Task EachItemGetsARow()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("B1").Value = "{{ item.Quantity }}";
        });

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A2").Value.GetText()).IsEqualTo("Gadget");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
    }

    [Test]
    public async Task GeneratedRowsKeepTheirValueTypes()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("B1").Value = "{{ item.UnitPrice }}";
            ws.Cell("C1").Value = "{{ item.SoldOn }}";
        });

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("B2").Value.IsNumber).IsTrue();
        await Assert.That(sheet.Cell("B2").Value.GetNumber()).IsEqualTo(10d);
        await Assert.That(sheet.Cell("C2").Value.IsDateTime).IsTrue();
    }

    [Test]
    public async Task EmptyOptionsRowIsRemoved()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ item.Product }}");

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(3);
    }

    [Test]
    public async Task OptionsRowWithContentIsKept()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("A2").Value = "Total";
        });

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(workbook.Worksheet("Report").Cell("A4").Value.GetText()).IsEqualTo("Total");
    }

    [Test]
    public async Task ContentBelowTheRangeMovesDown()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("A3").Value = "Footer";
        });

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(workbook.Worksheet("Report").Cell("A4").Value.GetText()).IsEqualTo("Footer");
    }

    [Test]
    public async Task EmptyCollectionRemovesTheTemplateRows()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("A3").Value = "Footer";
        });

        Generate(workbook, "Items", new List<SaleItem>());

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetText()).IsEqualTo("Footer");
    }

    [Test]
    public async Task SingleItemLeavesTheRangeItsOriginalSize()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("A3").Value = "Footer";
        });

        Generate(workbook, "Items", Items(("Only", 1, 1m)));

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Only");
        await Assert.That(sheet.Cell("A2").Value.GetText()).IsEqualTo("Footer");
    }

    [Test]
    public async Task MultiRowTemplateBlockRepeatsAsAWhole()
    {
        using var workbook = TemplateWithItemsRange(
            ws =>
            {
                ws.Cell("A1").Value = "{{ item.Product }}";
                ws.Cell("A2").Value = "qty {{ item.Quantity }}";
            },
            dataRows: 2);

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A2").Value.GetText()).IsEqualTo("qty 2");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Gadget");
        await Assert.That(sheet.Cell("A4").Value.GetText()).IsEqualTo("qty 3");
        await Assert.That(sheet.Cell("A5").Value.GetText()).IsEqualTo("Doohickey");
    }

    [Test]
    public async Task IndexIsAvailablePerRow()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ index }}");

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetNumber()).IsEqualTo(0d);
        await Assert.That(sheet.Cell("A3").Value.GetNumber()).IsEqualTo(2d);
    }

    [Test]
    public async Task TheWholeCollectionIsAvailablePerRow()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ items.size }}");

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetNumber()).IsEqualTo(3d);
    }

    [Test]
    public async Task GlobalVariablesStayVisibleInsideTheRange()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ Company }}: {{ item.Product }}");

        using var template = new XLTemplate(workbook);
        template.AddVariable("Company", "Contoso");
        template.AddVariable("Items", ThreeItems());
        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A2").Value.GetText()).IsEqualTo("Contoso: Gadget");
    }

    [Test]
    public async Task CellsOutsideTheRangeUseGlobalVariables()
    {
        using var workbook = TemplateWithItemsRange(
            ws =>
            {
                ws.Cell("A1").Value = "Report for {{ Company }}";
                ws.Cell("A2").Value = "{{ item.Product }}";
            },
            firstRow: 2);

        using var template = new XLTemplate(workbook);
        template.AddVariable("Company", "Contoso");
        template.AddVariable("Items", ThreeItems());
        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A1").Value.GetText()).IsEqualTo("Report for Contoso");
    }

    /// <summary>
    /// A template formula written against the first data row has to follow each generated row.
    /// Expansion gets this from the core library's copy semantics rather than rewriting formulas
    /// itself.
    /// </summary>
    [Test]
    public async Task RelativeFormulasFollowTheirRow()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Quantity }}";
            ws.Cell("B1").Value = "{{ item.UnitPrice }}";
            ws.Cell("C1").FormulaA1 = "A1*B1";
        });

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("C1").FormulaA1).IsEqualTo("A1*B1");
        await Assert.That(sheet.Cell("C2").FormulaA1).IsEqualTo("A2*B2");
        await Assert.That(sheet.Cell("C3").FormulaA1).IsEqualTo("A3*B3");
    }

    [Test]
    public async Task GeneratedFormulasCanBeBuiltPerRow()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Quantity }}";
            ws.Cell("B1").Value = "&=A{{ index + 1 }}*2";
        });

        Generate(workbook, "Items", ThreeItems());

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("B1").FormulaA1).IsEqualTo("A1*2");
        await Assert.That(sheet.Cell("B3").FormulaA1).IsEqualTo("A3*2");
    }

    [Test]
    public async Task StylesAreCarriedToGeneratedRows()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Cell("A1").Style.Font.Bold = true;
        });

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(workbook.Worksheet("Report").Cell("A3").Style.Font.Bold).IsTrue();
    }

    [Test]
    public async Task MergedCellsAreRepeatedPerRow()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Range("B1:C1").Merge();
        });

        Generate(workbook, "Items", ThreeItems());

        var merged = workbook.Worksheet("Report").MergedRanges.Select(r => r.RangeAddress.ToStringRelative()).ToList();
        await Assert.That(merged).Contains("B1:C1");
        await Assert.That(merged).Contains("B3:C3");
    }

    [Test]
    public async Task RowHeightIsCarriedToGeneratedRows()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Product }}";
            ws.Row(1).Height = 28;
        });

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(workbook.Worksheet("Report").Row(3).Height).IsEqualTo(28d);
    }

    /// <summary>
    /// ClosedXML.Report copies a conditional format onto every generated cell, so three rules over
    /// three rows become nine (its issue #216, where a user reports the duplication "kills the
    /// generation time"). Expanding by insert-and-copy stretches the original rule instead.
    /// </summary>
    [Test]
    public async Task ConditionalFormatIsStretchedNotDuplicated()
    {
        using var workbook = TemplateWithItemsRange(ws =>
        {
            ws.Cell("A1").Value = "{{ item.Quantity }}";
            ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(2).Fill.SetBackgroundColor(XLColor.Red);
        });

        Generate(workbook, "Items", ThreeItems());

        var formats = workbook.Worksheet("Report").ConditionalFormats.ToList();
        await Assert.That(formats.Count).IsEqualTo(1);
        await Assert.That(formats[0].Ranges.Single().RangeAddress.ToStringRelative()).IsEqualTo("A1:A3");
    }

    [Test]
    public async Task DefinedNameEndsUpCoveringTheGeneratedRows()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ item.Product }}");

        Generate(workbook, "Items", ThreeItems());

        var range = workbook.DefinedNames.Single(n => n.Name == "Items").Ranges.Single();
        await Assert.That(range.RangeAddress.ToStringRelative()).IsEqualTo("A1:C3");
    }

    [Test]
    public async Task TwoRangesOnOneSheetBothExpand()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.Cell("A4").Value = "{{ item.Product }}";
        workbook.DefinedNames.Add("First", sheet.Range("A1:C2"));
        workbook.DefinedNames.Add("Second", sheet.Range("A4:C5"));

        using var template = new XLTemplate(workbook);
        template.AddVariable("First", Items(("A1", 1, 1m), ("A2", 1, 1m)));
        template.AddVariable("Second", Items(("B1", 1, 1m), ("B2", 1, 1m), ("B3", 1, 1m)));
        template.Generate();

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("A1");
        await Assert.That(sheet.Cell("A2").Value.GetText()).IsEqualTo("A2");

        // The first range grew by one row, so the second started one row lower than it was authored.
        await Assert.That(sheet.Cell("A4").Value.GetText()).IsEqualTo("B1");
        await Assert.That(sheet.Cell("A6").Value.GetText()).IsEqualTo("B3");
    }

    [Test]
    public async Task SheetScopedDefinedNamesAreBound()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        sheet.DefinedNames.Add("Items", sheet.Range("A1:C2"));

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
    }

    [Test]
    public async Task UnderscoreNamesBindThroughAPropertyPath()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ item.Product }}", name: "Order_Lines");

        using var template = new XLTemplate(workbook);
        template.AddVariable("Order", new Order { Lines = ThreeItems() });
        template.Generate();

        await Assert.That(workbook.Worksheet("Report").Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
    }

    /// <summary>
    /// Where an underscore name stops being an Excel identifier. Its first segment picks a variable,
    /// so it matches the way Excel matches a name; the segments after it are read off the object as
    /// the C# members they are, and match exactly.
    /// </summary>
    [Test]
    [Arguments("ORDER_Lines", "Doohickey")]
    [Arguments("order_Lines", "Doohickey")]
    [Arguments("Order_lines", null)]
    [Arguments("Order_LINES", null)]
    public async Task OnlyAnUnderscoreNamesFirstSegmentIgnoresCase(string name, string? expected)
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ item.Product }}", name: name);

        using var template = new XLTemplate(workbook);
        template.AddVariable("Order", new Order { Lines = ThreeItems() });
        var result = template.Generate();

        await Assert.That(result.ParsingErrors).IsEmpty();

        var cell = workbook.Worksheet("Report").Cell("A3");
        if (expected is null)
        {
            await Assert.That(cell.IsEmpty()).IsTrue();
        }
        else
        {
            await Assert.That(cell.Value.GetText()).IsEqualTo(expected);
        }
    }

    [Test]
    public async Task DefinedNamesWithNoMatchingVariableAreLeftAlone()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "untouched";
        workbook.DefinedNames.Add("SomeRange", sheet.Range("A1:C2"));

        using var template = new XLTemplate(workbook);
        template.Generate();

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("untouched");
        await Assert.That(workbook.DefinedNames.Contains("SomeRange")).IsTrue();
    }

    [Test]
    public async Task DefinedNameBoundToASingleValueIsNotExpanded()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ Title }}";
        workbook.DefinedNames.Add("Title", sheet.Range("A1:A1"));

        Generate(workbook, "Title", "Annual Report");

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Annual Report");
    }

    [Test]
    public async Task SingleRowRangeHasNoOptionsRow()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "{{ item.Product }}";
        workbook.DefinedNames.Add("Items", sheet.Range("A1:C1"));

        Generate(workbook, "Items", ThreeItems());

        await Assert.That(sheet.Cell("A1").Value.GetText()).IsEqualTo("Widget");
        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
    }

    [Test]
    public async Task ManyItemsExpandCorrectly()
    {
        using var workbook = TemplateWithItemsRange(ws => ws.Cell("A1").Value = "{{ index }}");
        var many = Enumerable.Range(0, 500)
            .Select(i => new SaleItem { Product = "P" + i, Quantity = i, UnitPrice = 1m })
            .ToList();

        Generate(workbook, "Items", many);

        var sheet = workbook.Worksheet("Report");
        await Assert.That(sheet.Cell("A1").Value.GetNumber()).IsEqualTo(0d);
        await Assert.That(sheet.Cell("A500").Value.GetNumber()).IsEqualTo(499d);
        await Assert.That(sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber).IsEqualTo(500);
    }

    [Test]
    public async Task GeneratedWorkbookSurvivesARoundTrip()
    {
        using var stream = new System.IO.MemoryStream();

        using (var workbook = TemplateWithItemsRange(ws =>
               {
                   ws.Cell("A1").Value = "{{ item.Product }}";
                   ws.Cell("B1").Value = "{{ item.UnitPrice }}";
                   ws.Cell("C1").FormulaA1 = "B1*2";
               }))
        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("Items", ThreeItems());
            template.Generate();
            template.SaveAs(stream);
        }

        stream.Position = 0;
        using var reloaded = new XLWorkbook(stream);
        var sheet = reloaded.Worksheet("Report");

        await Assert.That(sheet.Cell("A3").Value.GetText()).IsEqualTo("Doohickey");
        await Assert.That(sheet.Cell("B3").Value.GetNumber()).IsEqualTo(7.5);
        await Assert.That(sheet.Cell("C3").FormulaA1).IsEqualTo("B3*2");
    }

    private sealed class Order
    {
        public List<SaleItem> Lines { get; set; } = new();
    }
}
