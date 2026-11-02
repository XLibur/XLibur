using XLibur.Excel;
using XLibur.Operations.Paths;

[assembly: NotInParallel]

namespace XLibur.Operations.Tests;

public static class TestDefaults
{
    [Before(HookType.Assembly)]
    public static void GlobalSetup() => Fonts.PinataLabs.PinataLabsFontBootstrap.Register();
}

internal static class Fixtures
{
    public static string Sales(TempDirectory directory, string name = "sales.xlsx")
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Sales");
        ws.Cell("A1").Value = "Product";
        ws.Cell("B1").Value = "Qty";
        ws.Cell("C1").Value = "Price";
        ws.Cell("D1").Value = "Total";
        object[][] rows = [["Apple", 3, 1.5], ["Banana", 12, 0.25], ["Cherry", 40, 0.1]];
        for (var i = 0; i < rows.Length; i++)
        {
            var row = i + 2;
            ws.Cell(row, 1).Value = (string)rows[i][0];
            ws.Cell(row, 2).Value = (int)rows[i][1];
            ws.Cell(row, 3).Value = (double)rows[i][2];
            ws.Cell(row, 4).FormulaA1 = $"B{row}*C{row}";
        }

        ws.Cell("A5").Value = "Total";
        ws.Cell("B5").FormulaA1 = "SUM(B2:B4)";
        ws.Cell("D5").FormulaA1 = "SUM(D2:D4)";
        ws.Range("A1:D4").CreateTable("SalesTable");
        workbook.DefinedNames.Add("GrandTotal", "Sales!$D$5");
        return Save(workbook, directory, name);
    }

    public static string Build(TempDirectory directory, string name, Action<XLWorkbook> build)
    {
        using var workbook = new XLWorkbook();
        build(workbook);
        return Save(workbook, directory, name);
    }

    public static OperationContext Context(TempDirectory root, bool allowWrite = true, Limits? limits = null)
        => new(PathPolicy.Restricted([root.Path], allowWrite), limits);

    public static (OperationOutcome Outcome, string Text) Run(OperationContext context, IOperation operation, object request,
        ResultFormat format = ResultFormat.Text)
    {
        var outcome = OperationRunner.Run(operation, request, context, format, out var text);
        return (outcome, text);
    }

    private static string Save(XLWorkbook workbook, TempDirectory directory, string name)
    {
        var path = directory.Combine(name);
        workbook.SaveAs(path, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        return path;
    }
}
