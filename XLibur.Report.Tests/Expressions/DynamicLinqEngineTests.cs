using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Report.DynamicLinq;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tests.Expressions;

/// <summary>
/// The compatibility engine's expression language: ClosedXML.Report's syntax, which is C#.
/// </summary>
/// <remarks>
/// These are the unit tests for the engine itself. <see cref="BothEnginesTests"/> runs whole templates
/// through it, which is the test that matters for the promise the package makes.
/// </remarks>
public class DynamicLinqEngineTests
{
    private static readonly DynamicLinqExpressionEngine Engine = new();

    private static List<SaleItem> Items() => new()
    {
        new SaleItem { Product = "Trowel", Region = "North", Quantity = 96, UnitPrice = 4.20m },
        new SaleItem { Product = "Hoe", Region = "South", Quantity = 4, UnitPrice = 240.00m },
        new SaleItem { Product = "Twine", Region = "North", Quantity = 240, UnitPrice = 1.15m },
    };

    private static ExpressionScope Scope(params (string Name, object? Value)[] values) =>
        new(values.Select(v => new KeyValuePair<string, object?>(v.Name, v.Value)));

    private static ExpressionScope RowScope(int index = 0)
    {
        var items = Items();

        return Scope(("Company", "Contoso"))
            .CreateChild(new[]
            {
                new KeyValuePair<string, object?>("item", items[index]),
                new KeyValuePair<string, object?>("index", index),
                new KeyValuePair<string, object?>("items", (object?)items),
            });
    }

    [Test]
    public async Task AVariableIsReadByName()
    {
        await Assert.That(Engine.Evaluate("Company", Scope(("Company", "Contoso")))).IsEqualTo("Contoso");
    }

    /// <summary>
    /// Upstream templates prefix a workbook variable with <c>@</c> when reading it from inside a range.
    /// Both forms work, because a template mixing them is common and neither is wrong.
    /// </summary>
    [Test]
    public async Task AVariableIsAlsoReadableWithAnAtPrefix()
    {
        await Assert.That(Engine.Evaluate("@Company", RowScope())).IsEqualTo("Contoso");
    }

    [Test]
    public async Task APropertyOfTheCurrentItemIsRead()
    {
        await Assert.That(Engine.Evaluate("item.Product", RowScope())).IsEqualTo("Trowel");
    }

    /// <summary>
    /// The point of the whole package: a .NET method call, which the default engine's language has no
    /// way to express.
    /// </summary>
    [Test]
    public async Task AMethodOnAPropertyIsCalled()
    {
        await Assert.That(Engine.Evaluate("item.Product.ToUpper()", RowScope())).IsEqualTo("TROWEL");
    }

    [Test]
    public async Task ASubstringWorksTheWayCSharpWouldWriteIt()
    {
        await Assert.That(Engine.Evaluate("item.Product.Substring(0, 3)", RowScope())).IsEqualTo("Tro");
    }

    [Test]
    public async Task ArithmeticKeepsItsType()
    {
        var value = Engine.Evaluate("item.Quantity * item.UnitPrice", RowScope());

        await Assert.That(value).IsTypeOf<decimal>();
        await Assert.That((decimal)value!).IsEqualTo(403.20m);
    }

    [Test]
    public async Task AComputedPropertyIsRead()
    {
        await Assert.That((decimal)Engine.Evaluate("item.Total", RowScope())!).IsEqualTo(403.20m);
    }

    [Test]
    public async Task TheConditionalOperatorWorks()
    {
        await Assert.That(Engine.Evaluate("item.Quantity > 10 ? \"bulk\" : \"unit\"", RowScope()))
            .IsEqualTo("bulk");
    }

    [Test]
    public async Task TheRowIndexIsInScope()
    {
        await Assert.That(Engine.Evaluate("index", RowScope(index: 2))).IsEqualTo(2);
    }

    /// <summary>LINQ over the collection, which is the other half of what upstream syntax was for.</summary>
    [Test]
    public async Task LinqOverTheCollectionWorks()
    {
        await Assert.That((decimal)Engine.Evaluate("items.Sum(x => x.Total)", RowScope())!)
            .IsEqualTo(403.20m + 960.00m + 276.00m);
    }

    [Test]
    public async Task LinqCanFilterBeforeAggregating()
    {
        await Assert.That((decimal)Engine.Evaluate("items.Where(x => x.Region == \"North\").Sum(x => x.Total)", RowScope())!)
            .IsEqualTo(403.20m + 276.00m);
    }

    [Test]
    public async Task LinqCanCount()
    {
        await Assert.That(Engine.Evaluate("items.Count(x => x.Quantity > 50)", RowScope())).IsEqualTo(2);
    }

    /// <summary>
    /// A collection with no element type of its own gives LINQ nothing to work with, so the engine
    /// re-casts it to the type its items report — which is what makes an untyped list from a
    /// <c>DataTable</c> or an <c>ArrayList</c> usable.
    /// </summary>
    [Test]
    public async Task ANonGenericCollectionIsUsableWithLinq()
    {
        var list = new ArrayList(Items());
        var scope = Scope(("rows", list));

        await Assert.That((decimal)Engine.Evaluate("rows.Sum(x => x.Total)", scope)!)
            .IsEqualTo(403.20m + 960.00m + 276.00m);
    }

    [Test]
    public async Task InterpolationMixesTextAndExpressions()
    {
        await Assert.That(Engine.Interpolate("Sold {{ item.Quantity }} of {{ item.Product }}", RowScope()))
            .IsEqualTo("Sold 96 of Trowel");
    }

    /// <summary>
    /// Interpolation is culture-invariant because its other job is building a formula, and a formula
    /// cannot contain a locale's decimal comma.
    /// </summary>
    [Test]
    public async Task InterpolationIsCultureInvariant()
    {
        // Set here rather than by attribute: the suite resets the culture before every test, so this
        // affects nothing after it.
        TestDefaults.SetCulture(CultureInfo.GetCultureInfo("cs-CZ"));

        await Assert.That(Engine.Interpolate("{{ item.UnitPrice }}", RowScope())).IsEqualTo("4.20");
    }

    /// <summary>
    /// A date interpolates as its OLE Automation serial: the only form a formula can do arithmetic on,
    /// and what upstream produced.
    /// </summary>
    [Test]
    public async Task ADateInterpolatesAsItsSerialNumber()
    {
        var scope = Scope(("when", new DateTime(2026, 7, 30)));
        var expected = new DateTime(2026, 7, 30).ToOADate().ToString(CultureInfo.InvariantCulture);

        await Assert.That(Engine.Interpolate("{{ when }}", scope)).IsEqualTo(expected);
    }

    [Test]
    public async Task TextWithNoExpressionsIsReturnedUnchanged()
    {
        await Assert.That(Engine.Interpolate("just text", RowScope())).IsEqualTo("just text");
    }

    /// <summary>An unclosed marker is text, not a failure: the author will see it in the output.</summary>
    [Test]
    public async Task AnUnclosedMarkerIsLeftAlone()
    {
        await Assert.That(Engine.Interpolate("a {{ item.Product", RowScope())).IsEqualTo("a {{ item.Product");
    }

    [Test]
    public async Task ANullValueInterpolatesAsNothing()
    {
        await Assert.That(Engine.Interpolate("[{{ Notes }}]", Scope(("Notes", null)))).IsEqualTo("[]");
    }

    [Test]
    public async Task AnUnparseableExpressionIsReported()
    {
        await Assert.That(() => Engine.Evaluate("item.Quantity +", RowScope()))
            .Throws<ExpressionEvaluationException>();
    }

    [Test]
    public async Task AnUnknownNameIsReported()
    {
        await Assert.That(() => Engine.Evaluate("NoSuchVariable", Scope(("Company", "Contoso"))))
            .Throws<ExpressionEvaluationException>();
    }

    /// <summary>
    /// A member access that throws at runtime — a null in the middle of a path — is reported as an
    /// expression failure rather than escaping as whatever the CLR threw inside a reflection call.
    /// </summary>
    [Test]
    public async Task AFailureInsideTheExpressionIsReported()
    {
        var scope = Scope(("item", new SaleItem { Notes = null }));

        await Assert.That(() => Engine.Evaluate("item.Notes.Length", scope))
            .Throws<ExpressionEvaluationException>();
    }

    /// <summary>
    /// The engine declines to host functions, and says so rather than accepting them and ignoring them.
    /// </summary>
    [Test]
    public async Task TheEngineDeclinesFunctions()
    {
        await Assert.That(Engine.SupportsFunctions).IsFalse();
        await Assert.That(() => Engine.AddFunction("SUM", new Func<int, int>(x => x)))
            .Throws<NotSupportedException>();
    }

    /// <summary>
    /// The same expression against the same scope shape is parsed once. Not observable directly, so
    /// this asserts the consequence that matters: repeated evaluation stays correct, which a cache keyed
    /// only on the text would not be once the scope shape changed.
    /// </summary>
    [Test]
    public async Task TheSameExpressionIsCorrectAgainstDifferentScopeShapes()
    {
        var engine = new DynamicLinqExpressionEngine();

        await Assert.That(engine.Evaluate("value", Scope(("value", 42)))).IsEqualTo(42);
        await Assert.That(engine.Evaluate("value", Scope(("value", "text")))).IsEqualTo("text");
        await Assert.That(engine.Evaluate("value", Scope(("value", 7.5m)))).IsEqualTo(7.5m);
    }
}
