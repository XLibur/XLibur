using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Excel.IO;

namespace XLibur.Tests.Excel.NamedRanges;

/// <summary>
/// A defined name whose <c>refersTo</c> text the formula parser cannot read must not take the whole
/// workbook down with it.
/// </summary>
/// <remarks>
/// The reader handed the raw text straight to <c>IXLDefinedName.RefersTo</c>, whose setter parses it,
/// so a single empty or malformed <c>&lt;definedName&gt;</c> — which Excel opens without complaint and
/// third-party writers do emit — aborted the load with <c>ClosedXML.Parser.ParsingException</c>, an
/// exception type from a dependency that tells a caller nothing about which file is at fault. The
/// name is kept with its text intact instead, reported as invalid, and written back unchanged, which
/// is what the print-area reader beside it already does for a formula it cannot resolve.
/// </remarks>
public class DefinedNameLoadResilienceTests
{
    /// <summary>
    /// A one-sheet package whose workbook part carries a single defined name <c>x</c> with
    /// <paramref name="refersToText"/> as its literal text.
    /// </summary>
    private static MemoryStream BookWithRawDefinedName(string refersToText)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = 1;
            wb.SaveAs(package);
        }

        return package.RewriteWorkbook(xml =>
        {
            var definedNames = $"<x:definedNames><x:definedName name=\"x\">{refersToText}</x:definedName></x:definedNames>";
            var rewritten = xml.Replace("<x:definedNames />", definedNames);
            if (ReferenceEquals(rewritten, xml) || !rewritten.Contains("definedName name=\"x\"", StringComparison.Ordinal))
                throw new InvalidOperationException("The defined name was not spliced into the workbook part.");

            return rewritten;
        });
    }

    private static IXLDefinedName TheName(XLWorkbook wb) => wb.DefinedNames.Single(dn => dn.Name == "x");

    [Test]
    [Arguments("")]
    [Arguments(" ")]
    [Arguments("SUM(Sheet1!$A$1")]
    [Arguments("@@@")]
    public async Task A_defined_name_the_parser_rejects_does_not_abort_the_load(string refersToText)
    {
        using var package = BookWithRawDefinedName(refersToText);

        using var wb = new XLWorkbook(package);

        await Assert.That(TheName(wb).RefersTo).IsEqualTo(refersToText);
    }

    /// <summary>
    /// A name whose formula is a bare local reference is one Excel itself refuses to open, so setting
    /// one is an error. Refusing it is still not a reason to fail the whole load: the name is kept as
    /// found, reported invalid, and written back so the file is no worse for having been opened.
    /// </summary>
    [Test]
    public async Task A_defined_name_with_a_local_reference_does_not_abort_the_load()
    {
        using var package = BookWithRawDefinedName("$A$1");

        using var wb = new XLWorkbook(package);

        await Assert.That(TheName(wb).RefersTo).IsEqualTo("$A$1");
        await Assert.That(TheName(wb).IsValid).IsFalse();
    }

    [Test]
    public async Task A_defined_name_with_a_local_reference_survives_a_row_insert()
    {
        using var package = BookWithRawDefinedName("$A$1");
        using var wb = new XLWorkbook(package);

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(TheName(wb).RefersTo).IsEqualTo("$A$1");
    }

    [Test]
    public async Task A_defined_name_the_parser_rejects_is_reported_as_invalid()
    {
        using var package = BookWithRawDefinedName("SUM(Sheet1!$A$1");

        using var wb = new XLWorkbook(package);

        await Assert.That(TheName(wb).IsValid).IsFalse();
    }

    [Test]
    public async Task A_defined_name_the_parser_rejects_round_trips_unchanged()
    {
        using var package = BookWithRawDefinedName("SUM(Sheet1!$A$1");

        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(package))
            wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(TheName(reloaded).RefersTo).IsEqualTo("SUM(Sheet1!$A$1");
    }

    [Test]
    public async Task A_defined_name_the_parser_rejects_survives_a_row_insert()
    {
        using var package = BookWithRawDefinedName("SUM(Sheet1!$A$1");
        using var wb = new XLWorkbook(package);

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(TheName(wb).RefersTo).IsEqualTo("SUM(Sheet1!$A$1");
    }

    [Test]
    public async Task A_defined_name_the_parser_rejects_survives_a_sheet_rename()
    {
        using var package = BookWithRawDefinedName("SUM(Sheet1!$A$1");
        using var wb = new XLWorkbook(package);

        wb.Worksheet("Sheet1").Name = "Renamed";

        await Assert.That(TheName(wb).RefersTo).IsEqualTo("SUM(Sheet1!$A$1");
    }

    [Test]
    public async Task A_defined_name_the_parser_rejects_survives_a_sheet_delete()
    {
        using var package = BookWithRawDefinedName("SUM(Sheet1!$A$1");
        using var wb = new XLWorkbook(package);

        wb.Worksheet("Sheet1").Delete();

        await Assert.That(TheName(wb).RefersTo).IsEqualTo("SUM(Sheet1!$A$1");
    }

    /// <summary>
    /// Leniency is for the reader alone. Code that sets a formula gets told the formula is bad, and
    /// gets told it in XLibur's own exception type rather than the parser's.
    /// </summary>
    [Test]
    [Arguments("")]
    [Arguments("   ")]
    [Arguments("SUM(Sheet1!$A$1")]
    [Arguments("@@@")]
    public async Task Setting_RefersTo_to_a_formula_the_parser_rejects_throws_ArgumentException(string formula)
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        var definedName = wb.DefinedNames.Add("x", "Sheet1!$A$1");

        await Assert.That(() => definedName.RefersTo = formula).Throws<ArgumentException>();
    }
}
