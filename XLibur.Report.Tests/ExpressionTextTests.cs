using System.Threading.Tasks;

namespace XLibur.Report.Tests;

public class ExpressionTextTests
{
    [Test]
    [Arguments("{{ item.Total }}", true)]
    [Arguments("Total: {{ item.Total }}", true)]
    [Arguments("{{a}}{{b}}", true)]
    [Arguments("plain text", false)]
    [Arguments("", false)]
    [Arguments("{{ unterminated", false)]
    public async Task ContainsDetectsExpressions(string text, bool expected)
    {
        await Assert.That(ExpressionText.Contains(text)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("{{ item.Total }}", "item.Total")]
    [Arguments("{{item.Total}}", "item.Total")]
    [Arguments("  {{ item.Total }}  ", "item.Total")]
    public async Task WholeCellExpressionIsRecognised(string text, string expected)
    {
        await Assert.That(ExpressionText.TryGetSingleExpression(text, out var expression)).IsTrue();
        await Assert.That(expression).IsEqualTo(expected);
    }

    [Test]
    [Arguments("Total: {{ item.Total }}")]
    [Arguments("{{ a }} and {{ b }}")]
    [Arguments("{{ a }} trailing")]
    [Arguments("plain text")]
    [Arguments("")]
    [Arguments("{{}}")]
    public async Task MixedOrAbsentExpressionIsNotAWholeCellExpression(string text)
    {
        await Assert.That(ExpressionText.TryGetSingleExpression(text, out var expression)).IsFalse();
        await Assert.That(expression).IsEqualTo(string.Empty);
    }
}
