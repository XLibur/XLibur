using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Issue #627. Conditional-format dxfs and the other dxfs (table fields, pivot formats) used to be
/// built by two near-copies that had drifted: the conditional-format copy dropped built-in number
/// formats and numbered custom ones by a different scheme, so the two could hand the same
/// <c>numFmtId</c> to different format codes.
/// </summary>
internal class DifferentialFormatNumberFormatTests
{
    [Test]
    public async Task A_conditional_format_with_a_built_in_number_format_round_trips()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5).NumberFormat
                .SetNumberFormatId(10);
            wb.SaveAs(ms);
        }

        var numFmts = ReadDxfNumberFormats(ms.ToArray());
        await Assert.That(numFmts.Count).IsEqualTo(1);
        await Assert.That(numFmts[0]).IsNotNull();
        await Assert.That(numFmts[0]!.NumberFormatId?.Value).IsEqualTo(10U);
        await Assert.That(numFmts[0]!.FormatCode?.Value).IsEqualTo("0.00%");

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var cf = reloaded.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf.Style.NumberFormat.NumberFormatId).IsEqualTo(10);
    }

    /// <summary>
    /// <c>CT_NumFmt</c> requires <c>formatCode</c>, and Excel will not open a file whose dxf
    /// <c>&lt;numFmt&gt;</c> lacks one. Accounting (44) is missing from
    /// <see cref="XLPredefinedFormat.FormatCodes"/>, so the writer must supply its code itself.
    /// </summary>
    [Test]
    public async Task A_built_in_accounting_format_is_written_with_its_format_code()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5).NumberFormat
                .SetNumberFormatId(44);
            wb.SaveAs(ms);
        }

        var numFmt = ReadDxfNumberFormats(ms.ToArray()).Single();
        await Assert.That(numFmt?.NumberFormatId?.Value).IsEqualTo(44U);
        await Assert.That(numFmt?.FormatCode?.Value)
            .IsEqualTo("_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)");

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var cf = reloaded.Worksheet("Sheet1").ConditionalFormats.Single();
        await Assert.That(cf.Style.NumberFormat.NumberFormatId).IsEqualTo(44);
    }

    /// <summary>
    /// A locale-specific built-in id has no code XLibur knows. Writing its <c>&lt;numFmt&gt;</c>
    /// without one would make the file unopenable, so the number format is left out.
    /// </summary>
    [Test]
    public async Task A_built_in_id_with_no_known_format_code_writes_no_number_format()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var cf = ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5);
            cf.NumberFormat.SetNumberFormatId(30);
            cf.Fill.SetBackgroundColor(XLColor.Red);
            wb.SaveAs(ms);
        }

        var numFmts = ReadDxfNumberFormats(ms.ToArray());
        await Assert.That(numFmts.Count).IsEqualTo(1);
        await Assert.That(numFmts[0]).IsNull();
    }

    [Test]
    public async Task Conditional_format_and_table_field_dxfs_never_share_a_number_format_id()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            // A conditional-format dxf with no number format, so that "all dxfs so far" and
            // "dxfs with a custom number format so far" disagree from here on.
            ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5).Fill
                .SetBackgroundColor(XLColor.Red);
            ws.Range("B1:B5").AddConditionalFormat().WhenLessThan(2).NumberFormat
                .SetFormat("0.000");

            // A table field whose data cells share a custom number format: written as a style dxf.
            ws.Cell("D1").Value = "Head";
            ws.Cell("D2").Value = 1;
            ws.Cell("D3").Value = 2;
            ws.Range("D2:D3").Style.NumberFormat.Format = "0.0000";
            ws.Range("D1:D3").CreateTable();

            wb.SaveAs(ms);
        }

        var numFmts = ReadDxfNumberFormats(ms.ToArray())
            .Where(nf => nf is not null)
            .Select(nf => (Id: nf!.NumberFormatId!.Value, Code: nf.FormatCode!.Value!))
            .ToList();

        await Assert.That(numFmts.Select(nf => nf.Code))
            .IsEquivalentTo(new[] { "0.000", "0.0000" });

        // Excel resolves a dxf numFmtId by id, so one id must never stand for two format codes.
        var codesPerId = numFmts.GroupBy(nf => nf.Id)
            .ToDictionary(g => g.Key, g => g.Select(nf => nf.Code).Distinct().Count());
        await Assert.That(codesPerId.Values.All(count => count == 1)).IsTrue()
            .Because($"ids: {string.Join(", ", numFmts.Select(nf => $"{nf.Id}={nf.Code}"))}");

        // Both formats reload onto the objects they came from.
        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var reloadedWs = reloaded.Worksheet("Sheet1");
        var cfFormats = reloadedWs.ConditionalFormats
            .Select(cf => cf.Style.NumberFormat.Format)
            .ToList();
        await Assert.That(cfFormats).Contains("0.000");
        await Assert.That(reloadedWs.Cell("D2").Style.NumberFormat.Format).IsEqualTo("0.0000");
    }

    /// <summary>The <c>&lt;numFmt&gt;</c> of every dxf in order, <c>null</c> for a dxf without one.</summary>
    private static List<NumberingFormat?> ReadDxfNumberFormats(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var doc = SpreadsheetDocument.Open(input, false);
        var dxfs = doc.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats;
        return dxfs is null
            ? []
            : dxfs.Elements<DifferentialFormat>()
                .Select(d => (NumberingFormat?)d.NumberingFormat?.CloneNode(true))
                .ToList();
    }
}
