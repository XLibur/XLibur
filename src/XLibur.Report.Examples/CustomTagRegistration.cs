using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Report.Tags;

namespace XLibur.Report.Examples;

/// <summary>
/// A tag of your own, registered into the template language.
/// </summary>
/// <remarks>
/// <para>
/// A tag has two moments it can act at, and may use both. <c>TransformItems</c> runs before a single
/// row exists, which is where anything that reorders or filters the data belongs.
/// <c>Execute</c> runs once the rows are there, which is where anything that refers to the generated
/// block belongs — a total, a border, a column width.
/// </para>
/// <para>
/// The two tags below take one moment each: <c>&lt;&lt;Top&gt;&gt;</c> keeps the largest few items
/// before anything is written, and <c>&lt;&lt;Banded&gt;&gt;</c> shades alternate rows once they exist.
/// Neither could be done at the other's moment, which is why there are two moments.
/// </para>
/// </remarks>
public sealed class CustomTagRegistration : ReportExample
{
    public override string Name => "CustomTag";

    public override string Summary => "Two tags of your own: one that filters the data, one that styles the result.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        // Registered once, for the process. Do this at start-up rather than per template.
        TagsRegister.Add<TopTag>("Top", priority: 15);
        TagsRegister.Add<BandedTag>("Banded", priority: 200);

        var sheet = workbook.AddWorksheet("Top lines");

        sheet.Cell("A1").Value = "The five largest lines";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

        Headings(sheet, 3, "Product", "Region", "Line total");

        sheet.Cell("A4").Value = "{{ item.Product }}";
        sheet.Cell("B4").Value = "{{ item.Region }}";
        sheet.Cell("C4").Value = "{{ item.Total }}";
        sheet.Cell("C4").Style.NumberFormat.Format = "#,##0.00";

        // <<Top>> runs at priority 15, after <<Sort>> at 10 — so the sort decides what "top" means and
        // the tag just takes the first few. Priority is how a tag says what it has to see first.
        sheet.Cell("C5").Value = "<<Desc>><<Top count=5>>";
        sheet.Cell("A5").Value = "<<Banded>>";

        workbook.DefinedNames.Add("Sales", sheet.Range("A4:C5"));

        sheet.Column(1).Width = 18;
        sheet.Columns(2, 3).Width = 14;
    }

    protected override void AddData(IXLTemplate template) =>
        template.AddVariable("Sales", SalesData.Sales());

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("Top lines");
        var lastRow = sheet.RangeUsed()!.RangeAddress.LastAddress.RowNumber;

        output.WriteLine($"  Bound     12 lines");
        output.WriteLine($"  Generated {lastRow - 3} rows   (<<Top count=5>> kept the largest five)");
        output.WriteLine($"  Largest   {sheet.Cell("A4").GetFormattedString()} at {sheet.Cell("C4").GetFormattedString()}"
            + "   (<<Desc>> ran first, so 'top' means what it should)");
        output.WriteLine();
    }

    /// <summary>
    /// Keeps the first <c>count</c> items and drops the rest. Written as
    /// <c>&lt;&lt;Top count=5&gt;&gt;</c>.
    /// </summary>
    /// <remarks>
    /// Acts before any row is written, so nothing downstream — the totals, the banding, a chart — has
    /// to know that anything was dropped. That is what <c>TransformItems</c> is for.
    /// </remarks>
    public sealed class TopTag : OptionTag
    {
        public override IReadOnlyList<object?> TransformItems(IReadOnlyList<object?> items, ProcessingContext context)
        {
            var count = (int)Token.Number("count", 10);

            return count >= items.Count ? items : items.Take(count).ToList();
        }
    }

    /// <summary>Shades alternate generated rows. Written as <c>&lt;&lt;Banded&gt;&gt;</c>.</summary>
    /// <remarks>
    /// Acts on the block, so it has to wait until the block exists — <c>Execute</c>. The context hands
    /// it the generated range, which is the only thing it needs to know.
    /// </remarks>
    public sealed class BandedTag : OptionTag
    {
        public override void Execute(ProcessingContext context)
        {
            var address = context.GeneratedRange.RangeAddress;
            var shade = XLColor.FromArgb(245, 245, 245);

            for (var row = address.FirstAddress.RowNumber + 1; row <= address.LastAddress.RowNumber; row += 2)
            {
                context.Worksheet
                    .Range(row, address.FirstAddress.ColumnNumber, row, address.LastAddress.ColumnNumber)
                    .Style.Fill.SetBackgroundColor(shade);
            }
        }
    }
}
