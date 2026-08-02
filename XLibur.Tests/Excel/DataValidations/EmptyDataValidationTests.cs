using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.DataValidations;

/// <summary>
/// A data validation rule applies to a set of ranges, written as the <c>sqref</c> attribute.
/// The schema requires that attribute to be non-empty, so a rule that has ended up covering
/// nothing must never reach the file: Excel treats <c>sqref=""</c> as corruption and repairs
/// the workbook, dropping every validation on the sheet.
/// </summary>
public class EmptyDataValidationTests
{
    [Test]
    public async Task AddingValidationOverAnExistingOne_DropsTheRuleItFullyCovers()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");

        ws.Range("A1:A3").CreateDataValidation().WholeNumber.EqualOrGreaterThan(1);
        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);

        // Covers every cell the first rule applied to, so that rule now applies to nothing.
        ws.Range("A1:A5").CreateDataValidation().WholeNumber.EqualOrGreaterThan(2);

        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(ws.DataValidations.Single().Ranges.Single().RangeAddress.ToString())
            .IsEqualTo("A1:A5");
    }

    /// <summary>
    /// The counterpart: a rule only partly covered keeps the remainder and must survive.
    /// </summary>
    [Test]
    public async Task AddingValidationOverPartOfAnExistingOne_KeepsTheUncoveredRemainder()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");

        ws.Range("A1:A5").CreateDataValidation().WholeNumber.EqualOrGreaterThan(1);
        ws.Range("A2:A3").CreateDataValidation().WholeNumber.EqualOrGreaterThan(2);

        await Assert.That(ws.DataValidations.Count()).IsEqualTo(2);

        var remainder = ws.DataValidations
            .Single(dv => dv.Ranges.Count() > 1)
            .Ranges.Select(r => r.RangeAddress.ToString())
            .ToList();

        await Assert.That(remainder).IsEquivalentTo(new string?[] { "A1:A1", "A4:A5" });
    }

    [Test]
    public async Task ValidationCoveredByAnother_IsAbsentFromTheSavedFile()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Range("A1:A3").CreateDataValidation().WholeNumber.EqualOrGreaterThan(1);
        ws.Range("A1:A5").CreateDataValidation().WholeNumber.EqualOrGreaterThan(2);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        var written = ReadSqrefs(ms);
        await Assert.That(written).IsEquivalentTo(new string?[] { "A1:A5" });
    }

    /// <summary>
    /// <c>ClearRanges</c> is public and leaves the rule in the collection with no coverage, so
    /// it reaches the writer independently of any range splitting.
    /// </summary>
    [Test]
    public async Task ValidationWithClearedRanges_IsAbsentFromTheSavedFile()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var cleared = ws.Range("A1:A3").CreateDataValidation();
        cleared.WholeNumber.EqualOrGreaterThan(1);
        ws.Range("C1:C3").CreateDataValidation().WholeNumber.EqualOrGreaterThan(2);

        cleared.ClearRanges();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        var written = ReadSqrefs(ms);
        await Assert.That(written).IsEquivalentTo(new string?[] { "C1:C3" });
    }

    /// <summary>
    /// When the only rule on the sheet is emptied there is nothing left to write, and the
    /// <c>dataValidations</c> element must not be emitted holding a single broken child.
    /// </summary>
    [Test]
    public async Task SheetWhoseOnlyValidationWasCleared_WritesNoDataValidationsElement()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var only = ws.Range("A1:A3").CreateDataValidation();
        only.WholeNumber.EqualOrGreaterThan(1);
        only.ClearRanges();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.Single().Worksheet;

        await Assert.That(worksheet!.Elements<DocumentFormat.OpenXml.Spreadsheet.DataValidations>().Any())
            .IsFalse();
    }

    [Test]
    public async Task SavedValidations_RoundTripWithoutLosingCoverage()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Range("A1:A3").CreateDataValidation().WholeNumber.EqualOrGreaterThan(1);
            ws.Range("A1:A5").CreateDataValidation().WholeNumber.EqualOrGreaterThan(2);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var sheet = reloaded.Worksheet("Data");

        await Assert.That(sheet.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(sheet.DataValidations.Single().Ranges.Single().RangeAddress.ToString())
            .IsEqualTo("A1:A5");
    }

    /// <summary>
    /// Every <c>sqref</c> written to the sheet. Any empty entry here is the corruption this
    /// class exists to prevent.
    /// </summary>
    private static string[] ReadSqrefs(MemoryStream saved)
    {
        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var worksheet = doc.WorkbookPart!.WorksheetParts.Single().Worksheet;

        return worksheet
            .Elements<DocumentFormat.OpenXml.Spreadsheet.DataValidations>()
            .SelectMany(dvs => dvs.Elements<DataValidation>())
            .Select(dv => dv.SequenceOfReferences!.InnerText)
            .ToArray();
    }
}
