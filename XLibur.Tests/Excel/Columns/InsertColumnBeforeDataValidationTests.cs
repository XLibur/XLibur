using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Columns;

[TestFixture]
public class InsertColumnBeforeDataValidationTests
{
    [Test]
    public void InsertColumnBeforeCol1_PreservesMultiColumnDvSqrefsThroughSave()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            // Multi-column dv straddling the col-1 insert point.
            var rule1 = ws.Range("A2:B2").CreateDataValidation();
            rule1.InputTitle = "Rule1";
            rule1.InputMessage = "Rule1";

            // Multi-column dv well right of the insert point.
            var rule2 = ws.Range("D2:E2").CreateDataValidation();
            rule2.InputTitle = "Rule2";
            rule2.InputMessage = "Rule2";

            ws.Column(1).InsertColumnsBefore(1);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var sheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var written = sheetPart.Worksheet.Descendants<DataValidation>()
            .Select(dv => (title: dv.PromptTitle?.Value ?? string.Empty,
                           sqref: dv.SequenceOfReferences!.InnerText))
            .ToList();

        var rule1Out = written.SingleOrDefault(x => x.title == "Rule1");
        var rule2Out = written.SingleOrDefault(x => x.title == "Rule2");
        Assert.Multiple(() =>
        {
            Assert.AreEqual("B2:C2", rule1Out.sqref,
                "Multi-column dv straddling the insert point lost its sqref during save.");
            Assert.AreEqual("E2:F2", rule2Out.sqref,
                "Multi-column dv right of the insert point lost its sqref during save.");
        });
    }

    [Test]
    public void InsertColumnBeforeCol1_ManyAdjacentDvsAllSurviveWithCorrectSqrefs()
    {
        // Reproduces the production scenario where a workbook has many distinct dvs
        // packed sequentially across columns. Each row-2 dv is a single cell at
        // consecutive columns; without the index-reconciliation fix in
        // XLDataValidations.Consolidate, dvs whose post-shift address happens to match
        // a neighbour's pre-shift address get double-shifted (or wiped) at save time,
        // producing empty sqrefs that Excel rejects on open.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            // Multi-column dv straddling the insert point + multi-row+col dv.
            Add(ws.Range("A2:B2"), "Straddle");
            Add(ws.Range("J3:K100"), "DataRange");

            // Single-cell row-2 dvs at adjacent columns: shifting any one
            // produces an address that collides with another's pre-shift address.
            for (var col = 3; col <= 9; col++)
                Add(ws.Cell(2, col).AsRange(), $"Cell{col}");

            // A dv with formula1 to verify formula round-trips (allowed-value rules
            // exercise a different writer path than property-only rules).
            var withFormula = ws.Range("M3:M50").CreateDataValidation();
            withFormula.InputTitle = "WithFormula";
            withFormula.TextLength.GreaterThan(3);

            ws.Column(1).InsertColumnsBefore(1);
            wb.SaveAs(saved);

            static void Add(IXLRange range, string title)
            {
                var rule = range.CreateDataValidation();
                rule.InputTitle = title;
                rule.InputMessage = title;
            }
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var sheetPart = doc.WorkbookPart!.WorksheetParts.First();
        var dvs = sheetPart.Worksheet.Descendants<DataValidation>()
            .Select(dv => (title: dv.PromptTitle?.Value ?? string.Empty,
                           sqref: dv.SequenceOfReferences!.InnerText,
                           f1: dv.Formula1?.InnerText ?? string.Empty))
            .ToList();

        var empties = dvs.Where(x => string.IsNullOrEmpty(x.sqref)).ToList();
        Assert.IsEmpty(empties,
            $"Saved file has {empties.Count} dv(s) with empty sqref — Excel will reject "
            + "the file with a 'Removed Records: Data validation' recovery dialog.");

        // Verify every original dv survived AND landed at the expected shifted address.
        // Cell{col} originally at column `col` should shift to `col + 1`.
        Assert.Multiple(() =>
        {
            Assert.AreEqual("B2:C2", FindSqref(dvs, "Straddle"));
            Assert.AreEqual("K3:L100", FindSqref(dvs, "DataRange"));
            for (var col = 3; col <= 9; col++)
            {
                var expectedCol = XLHelper.GetColumnLetterFromNumber(col + 1);
                Assert.AreEqual($"{expectedCol}2:{expectedCol}2", FindSqref(dvs, $"Cell{col}"),
                    $"Cell{col} dv should have shifted from column {col} to {col + 1}.");
            }

            var withFormulaOut = dvs.Single(x => x.title == "WithFormula");
            Assert.AreEqual("N3:N50", withFormulaOut.sqref,
                "Formula-bearing dv lost or mis-shifted its sqref.");
            Assert.AreEqual("3", withFormulaOut.f1);
        });
    }

    private static string FindSqref(
        System.Collections.Generic.IEnumerable<(string title, string sqref, string f1)> dvs,
        string title) => dvs.Single(x => x.title == title).sqref;
}
