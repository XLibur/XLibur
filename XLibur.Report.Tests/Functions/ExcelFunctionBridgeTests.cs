using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Report.Expressions;
using XLibur.Report.Functions;

namespace XLibur.Report.Tests.Functions;

public class ExcelFunctionBridgeTests
{
    private static ScribanExpressionEngine BridgedEngine()
    {
        var engine = new ScribanExpressionEngine();
        ExcelFunctionBridge.Register(engine);
        return engine;
    }

    private static ExpressionScope Scope(params (string Name, object? Value)[] values) =>
        new(values.Select(v => new KeyValuePair<string, object?>(v.Name, v.Value)));

    private static double Number(object? value) => Convert.ToDouble(value, CultureInfo.InvariantCulture);

    [Test]
    public async Task SumAddsItsArguments()
    {
        var result = BridgedEngine().Evaluate("SUM(1, 2, 3)", ExpressionScope.Empty);

        await Assert.That(Number(result)).IsEqualTo(6d);
    }

    /// <summary>
    /// <c>SUM(items.Total)</c> hands over one collection where Excel's SUM expects a run of
    /// arguments, so a collection argument is spread across several.
    /// </summary>
    [Test]
    public async Task SumAcceptsACollection()
    {
        var result = BridgedEngine().Evaluate("SUM(values)", Scope(("values", new[] { 1d, 2d, 3.5d })));

        await Assert.That(Number(result)).IsEqualTo(6.5d);
    }

    [Test]
    public async Task SumOverAProjectedCollection()
    {
        var items = new List<SaleItem>
        {
            new() { Product = "A", Quantity = 2, UnitPrice = 5m },
            new() { Product = "B", Quantity = 3, UnitPrice = 10m },
        };

        var result = BridgedEngine().Evaluate("SUM(items | array.map \"Total\")", Scope(("items", items)));

        await Assert.That(Number(result)).IsEqualTo(40d);
    }

    [Test]
    public async Task AverageWorksOverACollection()
    {
        var result = BridgedEngine().Evaluate("AVERAGE(values)", Scope(("values", new[] { 2d, 4d, 6d })));

        await Assert.That(Number(result)).IsEqualTo(4d);
    }

    [Test]
    [Arguments("ROUND(3.14159, 2)", 3.14)]
    [Arguments("ABS(-5)", 5d)]
    [Arguments("MAX(1, 9, 4)", 9d)]
    [Arguments("MIN(1, 9, 4)", 1d)]
    [Arguments("COUNT(1, 2, 3, 4)", 4d)]
    [Arguments("POWER(2, 10)", 1024d)]
    public async Task NumericFunctionsEvaluate(string expression, double expected)
    {
        var result = BridgedEngine().Evaluate(expression, ExpressionScope.Empty);

        await Assert.That(Number(result)).IsEqualTo(expected);
    }

    /// <summary>
    /// <c>if</c> is a Scriban keyword, which is why the bridge registers Excel names in upper
    /// case; <c>IF</c> parses as an ordinary function call.
    /// </summary>
    [Test]
    public async Task IfIsCallableDespiteBeingAScribanKeyword()
    {
        var engine = BridgedEngine();

        var bulk = engine.Evaluate("IF(12 > 10, \"bulk\", \"unit\")", ExpressionScope.Empty);
        var unit = engine.Evaluate("IF(2 > 10, \"bulk\", \"unit\")", ExpressionScope.Empty);

        await Assert.That(bulk).IsEqualTo("bulk");
        await Assert.That(unit).IsEqualTo("unit");
    }

    [Test]
    public async Task TextFunctionsEvaluate()
    {
        var engine = BridgedEngine();

        await Assert.That(engine.Evaluate("UPPER(\"widget\")", ExpressionScope.Empty)).IsEqualTo("WIDGET");
        await Assert.That(engine.Evaluate("LEFT(\"widget\", 3)", ExpressionScope.Empty)).IsEqualTo("wid");
        await Assert.That(engine.Evaluate("CONCATENATE(\"a\", \"b\")", ExpressionScope.Empty)).IsEqualTo("ab");
    }

    [Test]
    public async Task FunctionsReceiveModelValues()
    {
        var item = new SaleItem { Product = "Widget", Quantity = 3, UnitPrice = 9.5m };

        var result = BridgedEngine().Evaluate("ROUND(item.Total, 0)", Scope(("item", item)));

        await Assert.That(Number(result)).IsEqualTo(29d);
    }

    [Test]
    public async Task ResultsStayTypedForTheCell()
    {
        var value = XLibur.Report.Excel.ReportValueConverter.ToCellValue(
            BridgedEngine().Evaluate("SUM(1, 2, 3)", ExpressionScope.Empty));

        await Assert.That(value.IsNumber).IsTrue();
        await Assert.That(value.GetNumber()).IsEqualTo(6d);
    }

    /// <summary>Dates reach the function library as Excel serial numbers, as its date functions expect.</summary>
    [Test]
    public async Task DateArgumentsAreConvertedToSerialNumbers()
    {
        var result = BridgedEngine().Evaluate("YEAR(when)", Scope(("when", new DateTime(2026, 3, 14))));

        await Assert.That(Number(result)).IsEqualTo(2026d);
    }

    [Test]
    public async Task ErrorsComeBackAsExcelErrors()
    {
        var result = BridgedEngine().Evaluate("SQRT(-1)", ExpressionScope.Empty);

        await Assert.That(result).IsTypeOf<XLError>();
        await Assert.That((XLError)result!).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task WrongArgumentCountReportsAValueError()
    {
        var result = BridgedEngine().Evaluate("ROUND(1, 2, 3, 4)", ExpressionScope.Empty);

        await Assert.That(result).IsTypeOf<XLError>();
    }

    /// <summary>
    /// A template expression is evaluated before there is a grid, so a function that needs one has
    /// to say so rather than return something misleading.
    /// </summary>
    [Test]
    public async Task FunctionsNeedingAWorksheetAreRejected()
    {
        await Assert.That(() => BridgedEngine().Evaluate("ROW()", ExpressionScope.Empty))
            .Throws<ExpressionEvaluationException>();
    }

    [Test]
    public async Task TheBridgeCoversTheWholeFunctionLibrary()
    {
        var engine = BridgedEngine();

        // A sample spanning the registry's categories: maths, text, logical, date, statistical,
        // financial and information.
        foreach (var name in new[] { "SUM", "UPPER", "IF", "YEAR", "MEDIAN", "PMT", "ISNUMBER" })
        {
            var isRegistered = engine.Evaluate($"{name} != null", ExpressionScope.Empty);
            await Assert.That(isRegistered).IsTypeOf<bool>();
            await Assert.That((bool)isRegistered!).IsTrue();
        }
    }

    [Test]
    public async Task EnginesThatCannotHostFunctionsAreLeftAlone()
    {
        var engine = new FunctionlessEngine();

        ExcelFunctionBridge.Register(engine);

        await Assert.That(engine.AddFunctionCalls).IsEqualTo(0);
    }

    private sealed class FunctionlessEngine : IExpressionEngine
    {
        public int AddFunctionCalls { get; private set; }

        public bool SupportsFunctions => false;

        public object? Evaluate(string expression, ExpressionScope scope) => null;

        public string Interpolate(string text, ExpressionScope scope) => text;

        public void AddFunction(string name, Delegate function) => AddFunctionCalls++;
    }
}
