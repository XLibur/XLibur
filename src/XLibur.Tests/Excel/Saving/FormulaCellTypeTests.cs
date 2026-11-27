using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using SaveOptions = XLibur.Excel.SaveOptions;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// #962. A formula cell's <c>t</c> attribute describes its <c>&lt;v&gt;</c>, so the two are written
/// together or not at all. A dirty formula is written with no cached value (ADR 0001), and so with no
/// type either.
/// </summary>
public class FormulaCellTypeTests
{
    /// <summary>XLibur refuses it, so a save that evaluates formulas leaves it dirty.</summary>
    private const string Unevaluable = "ABS(1,2)";

    [Test]
    [Arguments(XLDataType.Error, false)]
    [Arguments(XLDataType.Error, true)]
    [Arguments(XLDataType.Text, false)]
    [Arguments(XLDataType.Text, true)]
    [Arguments(XLDataType.Boolean, false)]
    [Arguments(XLDataType.Boolean, true)]
    public async Task A_dirty_formula_is_written_with_neither_a_type_nor_a_value(XLDataType cachedType,
        bool evaluateFormulas)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var a1 = (XLCell)ws.Cell("A1");
        a1.FormulaA1 = Unevaluable;
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.SetCellValue(a1.SheetPoint, cachedType switch
        {
            XLDataType.Error => XLError.NameNotRecognized,
            XLDataType.Text => "old",
            _ => true,
        });
        await Assert.That(a1.NeedsRecalculation).IsTrue();

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = evaluateFormulas });

        await Assert.That(CellXml(stream.Sheet1Xml(), "A1")).IsEqualTo("<c r=\"A1\"><f>ABS(1,2)</f></c>");
    }

    [Test]
    public async Task A_formula_that_recalculates_on_save_is_written_with_the_type_of_its_new_value()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var a1 = (XLCell)ws.Cell("A1");
        a1.FormulaA1 = "\"a\"&\"b\"";
        ((XLWorksheet)ws).Internals.CellsCollection.ValueSlice.SetCellValue(a1.SheetPoint, XLError.NameNotRecognized);

        using var stream = new MemoryStream();
        wb.SaveAs(stream, new SaveOptions { EvaluateFormulasBeforeSaving = true });

        await Assert.That(CellXml(stream.Sheet1Xml(), "A1"))
            .IsEqualTo("<c r=\"A1\" t=\"str\"><f>\"a\"&amp;\"b\"</f><v>ab</v></c>");
    }

    /// <summary>
    /// The issue's own measurement: sheet 2, C1 holds <c>averageif(A2:B2)</c> with a cached
    /// <c>#NAME?</c>. An edit anywhere marked it dirty, and it was written <c>t="e"</c> with no value.
    /// </summary>
    [Test]
    public async Task After_an_edit_an_unevaluable_loaded_formula_is_written_with_no_error_type()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\bug-fixes.xlsx"));
        using var stream = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            wb.Worksheet(2).Cell("E5").Value = 1;
            wb.SaveAs(stream);
        }

        var cell = stream.PartsUnder("xl/worksheets/sheet")
            .Select(stream.ReadPart)
            .SelectMany(xml => XDocument.Parse(xml).Descendants().Where(e => e.Name.LocalName == "c"))
            .Single(c => c.Elements().Any(f => f.Name.LocalName == "f" && f.Value == "averageif(A2:B2)"));
        await Assert.That((string?)cell.Attribute("t")).IsNull();
        await Assert.That(cell.Elements().Any(e => e.Name.LocalName == "v")).IsFalse();
    }

    /// <summary>The cell's XML with namespaces and the style index left out.</summary>
    private static string CellXml(string sheetXml, string address)
    {
        var cell = XDocument.Parse(sheetXml).Descendants()
            .First(e => e.Name.LocalName == "c" && (string?)e.Attribute("r") == address);
        var plain = Strip(cell);
        plain.Attribute("s")?.Remove();
        return plain.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static XElement Strip(XElement element) =>
        new(element.Name.LocalName,
            element.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => new XAttribute(a.Name.LocalName, a.Value)),
            element.Nodes().Select(n => n is XElement child ? Strip(child) : n));
}
