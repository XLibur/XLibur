using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tests.Expressions;

public class ExpressionScopeTests
{
    private static ExpressionScope Scope(params (string Name, object? Value)[] values) =>
        new(values.Select(v => new KeyValuePair<string, object?>(v.Name, v.Value)));

    [Test]
    public async Task TryGetValueFindsOwnValue()
    {
        var scope = Scope(("a", 1));

        await Assert.That(scope.TryGetValue("a", out var value)).IsTrue();
        await Assert.That(value).IsEqualTo(1);
    }

    [Test]
    public async Task TryGetValueFindsInheritedValue()
    {
        var child = Scope(("a", 1)).CreateChild("b", 2);

        await Assert.That(child.TryGetValue("a", out var value)).IsTrue();
        await Assert.That(value).IsEqualTo(1);
    }

    [Test]
    public async Task InnerValueShadowsOuterValue()
    {
        var child = Scope(("a", 1)).CreateChild("a", 2);

        child.TryGetValue("a", out var value);

        await Assert.That(value).IsEqualTo(2);
    }

    [Test]
    public async Task TryGetValueReportsMissingName()
    {
        var scope = Scope(("a", 1));

        await Assert.That(scope.TryGetValue("missing", out var value)).IsFalse();
        await Assert.That(value).IsNull();
    }

    [Test]
    public async Task NamesAreCaseSensitive()
    {
        var scope = Scope(("Total", 1));

        await Assert.That(scope.TryGetValue("total", out _)).IsFalse();
    }

    [Test]
    public async Task FromOutermostOrdersTheChainForShadowing()
    {
        var outer = Scope(("a", 1));
        var middle = outer.CreateChild("b", 2);
        var inner = middle.CreateChild("c", 3);

        var chain = inner.FromOutermost();

        await Assert.That(chain.Count).IsEqualTo(3);
        await Assert.That(chain[0]).IsSameReferenceAs(outer);
        await Assert.That(chain[1]).IsSameReferenceAs(middle);
        await Assert.That(chain[2]).IsSameReferenceAs(inner);
    }

    [Test]
    public async Task ValuesExcludesInheritedNames()
    {
        var child = Scope(("a", 1)).CreateChild("b", 2);

        await Assert.That(child.Values.Count).IsEqualTo(1);
        await Assert.That(child.Values.ContainsKey("b")).IsTrue();
        await Assert.That(child.Values.ContainsKey("a")).IsFalse();
    }

    [Test]
    public async Task EmptyScopeHasNoValuesAndNoParent()
    {
        await Assert.That(ExpressionScope.Empty.Values.Count).IsEqualTo(0);
        await Assert.That(ExpressionScope.Empty.Parent).IsNull();
    }
}
