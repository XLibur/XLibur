using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.NamedRanges;

/// <summary>
/// A defined name whose <c>RefersTo</c> is a formula rather than a bare range must survive a row or
/// column shift with the formula intact.
/// </summary>
/// <remarks>
/// The shift used to rebuild <c>RefersTo</c> from the name's reference list alone, so everything
/// around the references — the function calls, the operators, the literal arguments — was discarded.
/// A dynamic range built on <c>OFFSET</c>/<c>COUNTA</c>, the standard idiom for a growing list, came
/// back as a two-area union of its own arguments and silently stopped meaning anything.
/// </remarks>
public class DefinedNameFormulaShiftTests
{
    private static XLWorkbook BookWithName(string refersTo)
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        wb.DefinedNames.Add("x", refersTo);
        return wb;
    }

    private static string RefersTo(XLWorkbook wb) => wb.DefinedNames.First(dn => dn.Name == "x").RefersTo;

    [Test]
    public async Task A_function_around_the_reference_survives_a_row_insert()
    {
        using var wb = BookWithName("SUM(Sheet1!$A$1:$A$5)");

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(RefersTo(wb)).IsEqualTo("SUM(Sheet1!$A$2:$A$6)");
    }

    [Test]
    public async Task An_operator_around_the_reference_survives_a_row_insert()
    {
        using var wb = BookWithName("Sheet1!$A$1:$A$5*2");

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(RefersTo(wb)).IsEqualTo("Sheet1!$A$2:$A$6*2");
    }

    [Test]
    public async Task A_dynamic_range_keeps_its_OFFSET_when_rows_are_inserted()
    {
        using var wb = BookWithName("OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)");

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(RefersTo(wb)).IsEqualTo("OFFSET(Sheet1!$A$2,0,0,COUNTA(Sheet1!$A:$A),1)");
    }

    [Test]
    public async Task A_function_around_the_reference_survives_a_column_insert()
    {
        using var wb = BookWithName("SUM(Sheet1!$B$1:$C$5)");

        wb.Worksheet("Sheet1").Column(1).InsertColumnsBefore(1);

        await Assert.That(RefersTo(wb)).IsEqualTo("SUM(Sheet1!$C$1:$D$5)");
    }

    [Test]
    public async Task A_function_around_the_reference_survives_a_row_delete()
    {
        using var wb = BookWithName("SUM(Sheet1!$A$3:$A$5)");

        wb.Worksheet("Sheet1").Row(1).Delete();

        await Assert.That(RefersTo(wb)).IsEqualTo("SUM(Sheet1!$A$2:$A$4)");
    }

    [Test]
    public async Task A_bare_range_still_shifts()
    {
        using var wb = BookWithName("Sheet1!$A$1:$A$5");

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(RefersTo(wb)).IsEqualTo("Sheet1!$A$2:$A$6");
    }

    [Test]
    public async Task A_name_with_no_references_at_all_is_left_alone()
    {
        using var wb = BookWithName("42");

        wb.Worksheet("Sheet1").Row(1).InsertRowsAbove(1);

        await Assert.That(RefersTo(wb)).IsEqualTo("42");
    }
}
