using System.IO;
using XLibur.Excel;

namespace XLibur.Report.Examples;

/// <summary>
/// What a mistake in a template does: not much, on purpose. And what turns out not to be a mistake.
/// </summary>
/// <remarks>
/// <para>
/// Generation never throws for a bad expression or a tag it cannot apply. The failure is recorded on
/// the result, the offending cell is left showing the message, and everything else is generated as if
/// nothing had happened.
/// </para>
/// <para>
/// That is deliberate, and it is the opposite of what a library normally does. A report of a hundred
/// pages with one bad cell is worth having; the same report as an exception is not, and a template
/// author who has to fix one thing at a time because each failure hides the next is being poorly
/// served. So <see cref="XLGenerateResult.HasErrors"/> is the thing to check, and every error carries
/// the sheet and cell it came from.
/// </para>
/// <para>
/// The second half of this example is the half worth reading: <b>a name that is not bound is not an
/// error</b>. The engine reads missing names and missing properties as blank, which is what makes a
/// template survive sparse data — an optional middle name, a discount that is only sometimes there —
/// without the template having to test for each. The cost is that a typo in a name is silent, so a
/// column that comes out empty is the first place to look for one.
/// </para>
/// <para>
/// This is the one example that ends with errors. Open its report: the broken cells say what went
/// wrong, their neighbours are fine, and the blanks are blank.
/// </para>
/// </remarks>
public sealed class ErrorsAreReportedNotThrown : ReportExample
{
    public override string Name => "ErrorHandling";

    public override string Summary =>
        "A bad expression costs its own cell and nothing else, and a missing name costs nothing at all. "
        + "Deliberately ends with errors.";

    protected override void BuildTemplate(IXLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Errors");

        sheet.Cell("A1").Value = "Mistakes, and a report that is generated anyway";
        sheet.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14);

        Reported(sheet);
        NotReported(sheet);
        PerRow(sheet);

        sheet.Column(1).Width = 40;
        sheet.Columns(2, 3).Width = 26;
    }

    /// <summary>The mistakes that are reported.</summary>
    private static void Reported(IXLWorksheet sheet)
    {
        sheet.Cell("A3").Value = "Reported as errors";
        sheet.Cell("A3").Style.Font.SetBold();

        Headings(sheet, 4, "What is wrong", "The cell");

        sheet.Cell("A5").Value = "Syntax the engine cannot parse";
        sheet.Cell("B5").Value = "{{ Company | | }}";

        sheet.Cell("A6").Value = "A function that does not exist";
        sheet.Cell("B6").Value = "{{ TOTALLYNOTAFUNCTION(1, 2) }}";

        sheet.Cell("A7").Value = "And this one is fine";
        sheet.Cell("B7").Value = "{{ Company }}";
    }

    /// <summary>
    /// The mistakes that are not reported, which is the part worth knowing.
    /// </summary>
    private static void NotReported(IXLWorksheet sheet)
    {
        sheet.Cell("A10").Value = "Not errors — read as blank";
        sheet.Cell("A10").Style.Font.SetBold();

        Headings(sheet, 11, "What it is", "The cell");

        sheet.Cell("A12").Value = "A variable nobody bound";
        sheet.Cell("B12").Value = "{{ NoSuchVariable }}";

        sheet.Cell("A13").Value = "A property the item has not got";
        sheet.Cell("B13").Value = "{{ Company.NoSuchProperty }}";

        sheet.Cell("A14").Value = "A tag outside a bound range is not read at all";
        sheet.Cell("B14").Value = "ordinary text <<Nonsense>>";

        sheet.Cell("A15").Value = "Relaxed access is what makes sparse data work: a template survives an";
        sheet.Cell("A16").Value = "optional field without testing for it, at the price of a silent typo.";
        sheet.Range("A15:A16").Style.Font.SetItalic().Font.SetFontColor(XLColor.Gray);
    }

    /// <summary>
    /// The same two kinds inside a bound range: one broken cell per row, and a tag the register has
    /// never heard of — which <em>is</em> reported, because inside a range the tags are read.
    /// </summary>
    private static void PerRow(IXLWorksheet sheet)
    {
        Headings(sheet, 17, "Product", "Line total", "Broken, once per row");

        sheet.Cell("A18").Value = "{{ item.Product }}";
        sheet.Cell("B18").Value = "{{ item.Total }}";
        sheet.Cell("B18").Style.NumberFormat.Format = "#,##0.00";
        sheet.Cell("C18").Value = "{{ item.Quantity + }}";

        // In the options row of a bound range, so this one is read — and reported, once.
        sheet.Cell("A19").Value = "<<Nonsense>>";

        sheet.Workbook.DefinedNames.Add("Sales", sheet.Range("A18:C19"));
    }

    protected override void AddData(IXLTemplate template)
    {
        template.AddVariable("Company", "Contoso Ltd");
        template.AddVariable("Sales", SalesData.Sales());
    }

    protected override void Describe(IXLWorkbook generated, TextWriter output)
    {
        var sheet = generated.Worksheet("Errors");

        output.WriteLine($"  B7  generated normally      '{sheet.Cell("B7").GetFormattedString()}'");
        output.WriteLine($"  B12 unbound name            '{sheet.Cell("B12").GetFormattedString()}'   (blank, not an error)");
        output.WriteLine($"  B13 missing property        '{sheet.Cell("B13").GetFormattedString()}'   (blank, not an error)");
        output.WriteLine($"  B14 tag outside a range     '{sheet.Cell("B14").GetFormattedString()}'   (left exactly as written)");
        output.WriteLine($"  A18 first generated row     '{sheet.Cell("A18").GetFormattedString()}'");
        output.WriteLine("  The errors below are the point of this example, not a failure of it.");
        output.WriteLine();
    }
}
