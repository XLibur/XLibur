using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Ranges;

/// <summary>
/// The template block is copied into the inserted slots by repeatedly doubling what has already been
/// written, rather than once per item. These tests pin that the two are indistinguishable in the output.
/// </summary>
/// <remarks>
/// <para>
/// The reason for doubling is that <c>CopyTo</c> costs more the larger the worksheet is, independently of
/// how much is being copied, so one call per item made generation super-linear in the row count — see
/// <c>ExpansionPhaseProbe</c> in <c>XLibur.Report.Benchmarks</c>. Doubling makes it ⌈log₂ n⌉ calls.
/// </para>
/// <para>
/// What that risks is the <em>boundaries</em>: the item counts where a doubling round is truncated
/// because fewer blocks are wanted than have been written. A power of two doubles exactly to the end;
/// one more or one fewer does not. So the counts below are chosen around those boundaries rather than
/// for roundness, and every one of them checks the first, last and a middle row — an off-by-one in the
/// rounding shows up as a duplicated or missing item, not as a formatting difference.
/// </para>
/// </remarks>
public class BlockCopyingTests
{
    private static List<SaleItem> Items(int count) => Enumerable.Range(1, count)
        .Select(i => new SaleItem
        {
            Product = "Product " + i,
            Quantity = i,
        })
        .ToList();

    /// <summary>Rows 2 (data) and 3 (options) bound to <c>Items</c>, with row 1 a heading.</summary>
    private static XLWorkbook Template(int rowsPerItem = 1)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        sheet.Cell("A1").Value = "Product";
        sheet.Cell("B1").Value = "Quantity";

        for (var row = 0; row < rowsPerItem; row++)
        {
            sheet.Cell(2 + row, 1).Value = "{{ item.Product }}";
            sheet.Cell(2 + row, 2).Value = "{{ item.Quantity }}";
        }

        workbook.DefinedNames.Add("Items", sheet.Range(2, 1, 2 + rowsPerItem, 2));

        return workbook;
    }

    private static XLGenerateResult Generate(XLWorkbook workbook, int itemCount)
    {
        using var template = new XLTemplate(workbook);
        template.AddVariable("Items", Items(itemCount));
        return template.Generate();
    }

    /// <summary>
    /// One row per item, at and around the powers of two where a doubling round is truncated.
    /// </summary>
    [Test]
    [Arguments(1)]
    [Arguments(2)]
    [Arguments(3)]
    [Arguments(4)]
    [Arguments(5)]
    [Arguments(7)]
    [Arguments(8)]
    [Arguments(9)]
    [Arguments(15)]
    [Arguments(16)]
    [Arguments(17)]
    [Arguments(31)]
    [Arguments(33)]
    [Arguments(100)]
    [Arguments(257)]
    public async Task EveryItemIsWrittenExactlyOnce(int itemCount)
    {
        using var workbook = Template();

        var result = Generate(workbook, itemCount);
        var sheet = workbook.Worksheet("Report");

        await Assert.That(result.HasErrors).IsFalse();

        // Every row holds its own item, in order, and there is nothing after the last one.
        for (var i = 1; i <= itemCount; i++)
        {
            await Assert.That(sheet.Cell(i + 1, 1).GetFormattedString()).IsEqualTo("Product " + i);
            await Assert.That(sheet.Cell(i + 1, 2).Value.GetNumber()).IsEqualTo(i);
        }

        await Assert.That(sheet.LastRowUsed()!.RowNumber()).IsEqualTo(itemCount + 1);
    }

    /// <summary>
    /// The same, with two rows per item: the doubling has to move whole blocks, and a block size the
    /// rounding does not account for would interleave the two rows of adjacent items.
    /// </summary>
    [Test]
    [Arguments(3)]
    [Arguments(5)]
    [Arguments(8)]
    [Arguments(11)]
    public async Task AMultiRowBlockIsCopiedWhole(int itemCount)
    {
        using var workbook = Template(rowsPerItem: 2);

        var result = Generate(workbook, itemCount);
        var sheet = workbook.Worksheet("Report");

        await Assert.That(result.HasErrors).IsFalse();

        for (var i = 1; i <= itemCount; i++)
        {
            var firstRow = 2 + ((i - 1) * 2);

            await Assert.That(sheet.Cell(firstRow, 1).GetFormattedString()).IsEqualTo("Product " + i);
            await Assert.That(sheet.Cell(firstRow + 1, 1).GetFormattedString()).IsEqualTo("Product " + i);
        }

        await Assert.That(sheet.LastRowUsed()!.RowNumber()).IsEqualTo(1 + (itemCount * 2));
    }

    /// <summary>
    /// A relative formula in the template row has to be re-pointed at the row it lands in — which the
    /// core library does on copy, and which doubling must not disturb. Copying a two-block region moves
    /// each of its rows by the same offset, so it holds, but it is the property most likely to break if
    /// the rounding were wrong.
    /// </summary>
    [Test]
    public async Task ARelativeFormulaFollowsItsOwnRow()
    {
        // Its own template: the formula has to be inside the bound range to be copied at all, so the
        // range runs to column C here rather than to B.
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");
        sheet.Cell("A1").Value = "Product";
        sheet.Cell("A2").Value = "{{ item.Product }}";
        sheet.Cell("B2").Value = "{{ item.Quantity }}";
        sheet.Cell("C2").FormulaA1 = "B2*2";
        workbook.DefinedNames.Add("Items", sheet.Range("A2:C3"));

        Generate(workbook, 10);

        for (var i = 1; i <= 10; i++)
        {
            await Assert.That(sheet.Cell(i + 1, 3).FormulaA1).IsEqualTo($"B{i + 1}*2");
        }
    }

    /// <summary>
    /// Styling declared on the template row reaches every generated row, whichever doubling round wrote
    /// it.
    /// </summary>
    [Test]
    public async Task StylingReachesEveryGeneratedRow()
    {
        using var workbook = Template();
        var sheet = workbook.Worksheet("Report");
        sheet.Cell("B2").Style.NumberFormat.Format = "#,##0.00";
        sheet.Cell("A2").Style.Font.SetBold();

        Generate(workbook, 20);

        for (var i = 1; i <= 20; i++)
        {
            await Assert.That(sheet.Cell(i + 1, 2).Style.NumberFormat.Format).IsEqualTo("#,##0.00");
            await Assert.That(sheet.Cell(i + 1, 1).Style.Font.Bold).IsTrue();
        }
    }
}
