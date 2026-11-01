using System.Linq;
using System.Threading.Tasks;
using XLibur.Report.Tags;

namespace XLibur.Report.Tests.Tags;

public class TagParserTests
{
    [Test]
    [Arguments("<<Sum>>", true)]
    [Arguments("Total <<Sum>>", true)]
    [Arguments("no tags", false)]
    [Arguments("", false)]
    [Arguments("<<unclosed", false)]
    [Arguments("a < b > c", false)]
    public async Task ContainsDetectsTags(string text, bool expected)
    {
        await Assert.That(TagParser.Contains(text)).IsEqualTo(expected);
    }

    [Test]
    public async Task TagNameIsRead()
    {
        var tags = TagParser.Parse("<<Sum>>");

        await Assert.That(tags.Count).IsEqualTo(1);
        await Assert.That(tags[0].Name).IsEqualTo("Sum");
    }

    [Test]
    public async Task SeveralTagsInOneCellAreRead()
    {
        var tags = TagParser.Parse("<<Sort>><<Hidden>>");

        await Assert.That(tags.Select(t => t.Name)).IsEquivalentTo(new[] { "Sort", "Hidden" });
    }

    [Test]
    public async Task BareParameterIsAFlag()
    {
        var tag = TagParser.Parse("<<Sort desc>>").Single();

        await Assert.That(tag.Flag("desc")).IsTrue();
        await Assert.That(tag.Flag("asc")).IsFalse();
    }

    [Test]
    public async Task AssignedParameterIsRead()
    {
        var tag = TagParser.Parse("<<Sum over=D>>").Single();

        await Assert.That(tag.Value("over")).IsEqualTo("D");
    }

    [Test]
    public async Task QuotedParameterKeepsItsSpaces()
    {
        var tag = TagParser.Parse("<<Sort by=\"item.Customer Name\">>").Single();

        await Assert.That(tag.Value("by")).IsEqualTo("item.Customer Name");
    }

    [Test]
    public async Task ParameterNamesAreCaseInsensitive()
    {
        var tag = TagParser.Parse("<<Sum OVER=D>>").Single();

        await Assert.That(tag.Value("over")).IsEqualTo("D");
    }

    [Test]
    public async Task SeveralParametersAreRead()
    {
        var tag = TagParser.Parse("<<Sort by=item.Total desc>>").Single();

        await Assert.That(tag.Value("by")).IsEqualTo("item.Total");
        await Assert.That(tag.Flag("desc")).IsTrue();
    }

    [Test]
    public async Task FlagIsTrueWhenAssignedTrue()
    {
        var tag = TagParser.Parse("<<Delete keep=true>>").Single();

        await Assert.That(tag.Flag("keep")).IsTrue();
    }

    [Test]
    public async Task NumericParameterIsRead()
    {
        var tag = TagParser.Parse("<<Height value=28.5>>").Single();

        await Assert.That(tag.Number("value", 0)).IsEqualTo(28.5);
    }

    [Test]
    public async Task StripRemovesTheTagsAndKeepsTheRest()
    {
        await Assert.That(TagParser.Strip("Total <<Sum>>")).IsEqualTo("Total");
        await Assert.That(TagParser.Strip("<<Sum>>")).IsEqualTo(string.Empty);
        await Assert.That(TagParser.Strip("plain")).IsEqualTo("plain");
    }

    [Test]
    public async Task MissingParameterFallsBack()
    {
        var tag = TagParser.Parse("<<Sum>>").Single();

        await Assert.That(tag.Value("over", "fallback")).IsEqualTo("fallback");
        await Assert.That(tag.Has("over")).IsFalse();
    }
}
