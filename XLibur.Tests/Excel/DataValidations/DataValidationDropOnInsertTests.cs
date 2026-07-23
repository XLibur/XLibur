using System.Linq;
using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.DataValidations;

// Regression coverage: inserting rows/columns must shift every data-validation rule by exactly
// the inserted amount without dropping any. The original defect wiped a rule whose extended
// neighbour transiently overlapped it during the shift (SplitExistingRanges against a
// not-yet-shifted rule). Related to the conditional-format #2850 double-shift class.
[TestFixture]
public class DataValidationDropOnInsertTests
{
    [Test]
    public void InsertRowsAbove_KeepsAllValidationsAndShiftsThem()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        // Distinct rules so nothing consolidates; 13+10 == 23 collides with an existing rule.
        ws.Range("K12:K12").CreateDataValidation().WholeNumber.GreaterThan(0);
        ws.Range("K13:K13").CreateDataValidation().WholeNumber.GreaterThan(1);
        ws.Range("K23:K23").CreateDataValidation().WholeNumber.GreaterThan(2);

        ws.Row(13).InsertRowsAbove(10);

        var actual = ws.DataValidations
            .SelectMany(dv => dv.Ranges.Select(r => r.RangeAddress.ToString()))
            .OrderBy(s => s)
            .ToList();

        // K12 spans the insertion -> K12:K22; K13 -> K23 (not dropped); K23 -> K33.
        Assert.That(actual, Is.EqualTo(new[] { "K12:K22", "K23:K23", "K33:K33" }));
    }

    [Test]
    public void InsertColumnsBefore_KeepsAllValidationsAndShiftsThem()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Range("B20:B20").CreateDataValidation().WholeNumber.GreaterThan(0);
        ws.Range("C20:C20").CreateDataValidation().WholeNumber.GreaterThan(1);
        ws.Range("M20:M20").CreateDataValidation().WholeNumber.GreaterThan(2);

        ws.Column(3).InsertColumnsBefore(10); // C+10 == M collides

        var actual = ws.DataValidations
            .SelectMany(dv => dv.Ranges.Select(r => r.RangeAddress.ToString()))
            .OrderBy(s => s, System.StringComparer.Ordinal)
            .ToList();

        // B spans the insertion boundary column -> B20:L20; C20 -> M20; M20 -> W20.
        Assert.That(actual, Is.EquivalentTo(new[] { "B20:L20", "M20:M20", "W20:W20" }));
    }
}
