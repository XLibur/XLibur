using XLibur.Excel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Globalization;

public class GlobalizationTests
{
    [Test]
    [Arguments("A1*10", 1230d)]
    [Arguments("A1/10", 12.3)]
    [Arguments("A1&\" cells\"", "123 cells")]
    [Arguments("A1&\"000\"", "123000")]
    [Arguments("ISNUMBER(A1)", true)]
    [Arguments("ISBLANK(A1)", false)]
    [Arguments("DATE(2018,1,28)", 43128d)]
    public async Task LoadFormulaCachedValue(string formula, object expectedValue)
    {
        Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-RU");

        using var ms = new MemoryStream();
        using (var book1 = new XLWorkbook())
        {
            var sheet = book1.AddWorksheet("sheet1");
            sheet.Cell("A1").Value = 123;
            sheet.Cell("A2").FormulaA1 = formula;
            var options = new SaveOptions { EvaluateFormulasBeforeSaving = true };

            book1.SaveAs(ms, options);
        }
        ms.Position = 0;

        using (var book2 = new XLWorkbook(ms))
        {
            var ws = book2.Worksheet(1);
            var storedValueA2 = ws.Cell("A2").CachedValue;
            await Assert.That(ws.Cell("A2").NeedsRecalculation).IsFalse();
            await Assert.That(storedValueA2).IsEqualTo(ExpectedCellValue.From(expectedValue));
        }
    }
}
