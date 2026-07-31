using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tests.Expressions;

public class ScribanExpressionEngineTests
{
    private static ExpressionScope Scope(params (string Name, object? Value)[] values) =>
        new(values.Select(v => new KeyValuePair<string, object?>(v.Name, v.Value)));

    private static ExpressionScope ItemScope(SaleItem item) => Scope(("item", item));

    private static SaleItem SampleItem() => new()
    {
        Product = "Widget",
        Quantity = 3,
        UnitPrice = 9.5m,
        SoldOn = new DateTime(2026, 3, 14),
        IsExport = true,
    };

    [Test]
    public async Task EvaluatePreservesDecimalType()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("item.UnitPrice", ItemScope(SampleItem()));

        await Assert.That(result).IsTypeOf<decimal>();
        await Assert.That(result).IsEqualTo(9.5m);
    }

    [Test]
    public async Task EvaluatePreservesDateTimeType()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("item.SoldOn", ItemScope(SampleItem()));

        await Assert.That(result).IsEqualTo(new DateTime(2026, 3, 14));
    }

    [Test]
    public async Task EvaluatePreservesBooleanType()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("item.IsExport", ItemScope(SampleItem()));

        await Assert.That(result).IsTypeOf<bool>();
        await Assert.That((bool)result!).IsTrue();
    }

    [Test]
    public async Task EvaluateComputesArithmetic()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("item.Quantity * item.UnitPrice", ItemScope(SampleItem()));

        await Assert.That(Convert.ToDecimal(result, CultureInfo.InvariantCulture)).IsEqualTo(28.5m);
    }

    [Test]
    public async Task EvaluateReadsComputedProperty()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("item.Total", ItemScope(SampleItem()));

        await Assert.That(result).IsEqualTo(28.5m);
    }

    /// <summary>
    /// Scriban renames PascalCase members to snake_case by default. The engine overrides that,
    /// because a template author binding a C# model expects to spell the C# names.
    /// </summary>
    [Test]
    public async Task MembersKeepTheirCSharpNames()
    {
        var engine = new ScribanExpressionEngine();

        var pascalCase = engine.Evaluate("item.UnitPrice", ItemScope(SampleItem()));
        var snakeCase = engine.Evaluate("item.unit_price", ItemScope(SampleItem()));

        await Assert.That(pascalCase).IsEqualTo(9.5m);
        await Assert.That(snakeCase).IsNull();
    }

    [Test]
    public async Task MissingMemberEvaluatesToNull()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("item.NoSuchMember", ItemScope(SampleItem()));

        await Assert.That(result).IsNull();
    }

    [Test]
    public async Task MissingVariableEvaluatesToNull()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Evaluate("missing", ExpressionScope.Empty);

        await Assert.That(result).IsNull();
    }

    [Test]
    public async Task NullPropertyEvaluatesToNull()
    {
        var engine = new ScribanExpressionEngine();
        var item = SampleItem();
        item.Notes = null;

        var result = engine.Evaluate("item.Notes", ItemScope(item));

        await Assert.That(result).IsNull();
    }

    [Test]
    public async Task InterpolateMixesTextAndExpressions()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Interpolate("Sold {{ item.Quantity }} of {{ item.Product }}", ItemScope(SampleItem()));

        await Assert.That(result).IsEqualTo("Sold 3 of Widget");
    }

    [Test]
    public async Task InterpolateLeavesPlainTextAlone()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Interpolate("no expressions here", ExpressionScope.Empty);

        await Assert.That(result).IsEqualTo("no expressions here");
    }

    [Test]
    public async Task InterpolateIsCultureInvariantByDefault()
    {
        var engine = new ScribanExpressionEngine();

        var result = engine.Interpolate("{{ value }}", Scope(("value", 1234.5d)));

        await Assert.That(result).IsEqualTo("1234.5");
    }

    [Test]
    public async Task InterpolateHonoursAnExplicitCulture()
    {
        var engine = new ScribanExpressionEngine(CultureInfo.GetCultureInfo("de-DE"));

        var result = engine.Interpolate("{{ value }}", Scope(("value", 1234.5d)));

        await Assert.That(result).IsEqualTo("1234,5");
    }

    /// <summary>
    /// The culture is readable, which is what lets sorting and group labels follow the report's
    /// culture rather than the machine's.
    /// </summary>
    [Test]
    public async Task CultureIsExposed()
    {
        var german = CultureInfo.GetCultureInfo("de-DE");

        await Assert.That(new ScribanExpressionEngine().Culture).IsEqualTo(CultureInfo.InvariantCulture);
        await Assert.That(new ScribanExpressionEngine(german).Culture).IsEqualTo(german);
    }

    [Test]
    public async Task InnerScopeShadowsOuterScope()
    {
        var engine = new ScribanExpressionEngine();
        var outer = Scope(("label", "outer"));
        var inner = outer.CreateChild("label", "inner");

        await Assert.That(engine.Evaluate("label", inner)).IsEqualTo("inner");
        await Assert.That(engine.Evaluate("label", outer)).IsEqualTo("outer");
    }

    [Test]
    public async Task OuterScopeStaysVisibleFromInnerScope()
    {
        var engine = new ScribanExpressionEngine();
        var outer = Scope(("company", "Contoso"));
        var inner = outer.CreateChild("item", SampleItem());

        var result = engine.Interpolate("{{ company }}: {{ item.Product }}", inner);

        await Assert.That(result).IsEqualTo("Contoso: Widget");
    }

    [Test]
    public async Task MalformedExpressionThrowsEvaluationException()
    {
        var engine = new ScribanExpressionEngine();

        await Assert.That(() => engine.Evaluate("item.", ItemScope(SampleItem())))
            .Throws<ExpressionEvaluationException>();
    }

    [Test]
    public async Task EvaluationExceptionCarriesTheExpression()
    {
        var engine = new ScribanExpressionEngine();

        var exception = Assert.Throws<ExpressionEvaluationException>(() => engine.Evaluate("item.", ItemScope(SampleItem())));

        await Assert.That(exception!.Expression).IsEqualTo("item.");
    }

    [Test]
    public async Task RegisteredFunctionIsCallable()
    {
        var engine = new ScribanExpressionEngine();
        engine.AddFunction("DOUBLE", new Func<double, double>(x => x * 2));

        var result = engine.Evaluate("DOUBLE(21)", ExpressionScope.Empty);

        await Assert.That(Convert.ToDouble(result, CultureInfo.InvariantCulture)).IsEqualTo(42d);
    }

    /// <summary>
    /// <c>if</c> is a Scriban keyword, so the Excel function bridge registers Excel names in
    /// upper case. This pins that an uppercase <c>IF</c> parses as an ordinary function call.
    /// </summary>
    [Test]
    public async Task UppercaseKeywordNamedFunctionIsCallable()
    {
        var engine = new ScribanExpressionEngine();
        engine.AddFunction("IF", new Func<bool, object, object, object>((condition, whenTrue, whenFalse) => condition ? whenTrue : whenFalse));

        var result = engine.Evaluate("IF(item.Quantity > 2, \"bulk\", \"unit\")", ItemScope(SampleItem()));

        await Assert.That(result).IsEqualTo("bulk");
    }

    [Test]
    public async Task FunctionsRegisteredAfterFirstEvaluationAreStillVisible()
    {
        var engine = new ScribanExpressionEngine();

        // Forces the template context to be created before the function is registered.
        engine.Evaluate("1 + 1", ExpressionScope.Empty);
        engine.AddFunction("TRIPLE", new Func<double, double>(x => x * 3));

        var result = engine.Evaluate("TRIPLE(5)", ExpressionScope.Empty);

        await Assert.That(Convert.ToDouble(result, CultureInfo.InvariantCulture)).IsEqualTo(15d);
    }

    [Test]
    public async Task SupportsFunctionsIsTrue()
    {
        await Assert.That(new ScribanExpressionEngine().SupportsFunctions).IsTrue();
    }

    [Test]
    public async Task RepeatedEvaluationOfTheSameExpressionReusesTheParsedTemplate()
    {
        var engine = new ScribanExpressionEngine();
        var items = Enumerable.Range(1, 50)
            .Select(i => new SaleItem { Product = $"P{i}", Quantity = i, UnitPrice = 2m })
            .ToList();

        var totals = items.Select(i => engine.Evaluate("item.Quantity * item.UnitPrice", ItemScope(i))).ToList();

        await Assert.That(totals.Count).IsEqualTo(50);
        await Assert.That(Convert.ToDecimal(totals[9], CultureInfo.InvariantCulture)).IsEqualTo(20m);
    }

    [Test]
    public async Task CollectionsAreEnumerableFromExpressions()
    {
        var engine = new ScribanExpressionEngine();
        var items = new List<SaleItem>
        {
            new() { Product = "A", Quantity = 2, UnitPrice = 5m },
            new() { Product = "B", Quantity = 4, UnitPrice = 5m },
        };

        var result = engine.Evaluate("items.size", Scope(("items", items)));

        await Assert.That(Convert.ToInt32(result, CultureInfo.InvariantCulture)).IsEqualTo(2);
    }
}
