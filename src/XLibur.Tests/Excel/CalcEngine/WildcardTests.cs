using System;
using XLibur.Excel.CalcEngine;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class WildcardTests
{
    [Test]
    [Arguments("")]
    [Arguments("abc")]
    public async Task Empty_Pattern_Matches_Any_String(string text)
    {
        await Assert.That(SearchWildcard(text, string.Empty)).IsEqualTo(0);
    }

    [Test]
    [Arguments("", "abc", 0)]
    [Arguments("a", "abc", 0)]
    [Arguments("ab", "abc", 0)]
    [Arguments("abc", "abc", 0)]
    [Arguments("bc", "abc", 1)]
    [Arguments("c", "abc", 2)]
    public async Task Substring_Of_Text_Matches_Text(string substringPattern, string text, int expectedIndex)
    {
        await Assert.That(SearchWildcard(text, substringPattern)).IsEqualTo(expectedIndex);
    }

    [Test]
    [Arguments("abcd", "abc")]
    public async Task Pattern_Not_In_Text_Returns_Negative_One(string pattern, string text)
    {
        await Assert.That(SearchWildcard(text, pattern)).IsEqualTo(-1);
    }

    [Test]
    public async Task Pattern_Comparison_Is_Case_Insensitive()
    {
        await Assert.That(SearchWildcard("zabcd", "AbCd")).IsEqualTo(1);
    }

    [Test]
    public async Task Tilde_Is_Escape_Char()
    {
        await Assert.That(SearchWildcard("_abc_", "~a~B~c")).IsEqualTo(1);
    }

    [Test]
    [Arguments("~*", "*", 0)]
    [Arguments("~*", "a", -1)]
    [Arguments("~?", "?", 0)]
    [Arguments("~?", "a", -1)]
    [Arguments("~a~b~", "ab", 0)]
    public async Task Escaped_Wildcards_Are_Matched_As_Chars(string pattern, string text, int expectedPosition)
    {
        await Assert.That(SearchWildcard(text, pattern)).IsEqualTo(expectedPosition);
    }

    [Test]
    public async Task Question_Mark_Wildcard_Matches_Any_Char()
    {
        await Assert.That(SearchWildcard("abc", "a?c")).IsEqualTo(0);
    }

    [Test]
    [Arguments("abcd", "ab*cd", 0)]
    [Arguments("aaab_____cd", "ab*cd", 2)]
    [Arguments("*abc*", "***a*b*c***", 0)]

    public async Task Star_Wildcard_Matches_Any_Number_Of_Chars(string text, string pattern, int index)
    {
        await Assert.That(SearchWildcard(text, pattern)).IsEqualTo(index);
    }

    [Test]
    public async Task Unpaired_Escape_Char_At_The_End_Of_Pattern_Is_Not_Char()
    {
        await Assert.That(SearchWildcard("a", "a~")).IsEqualTo(0);
    }

    [Test]
    public async Task Star_Wildcard_At_The_Beginning_Matches_First_Char()
    {
        await Assert.That(SearchWildcard("abcccd", "*ccd")).IsEqualTo(0);
    }

    [Test]
    public async Task Pattern_Size_Is_Limited_To_255_Chars()
    {
        await Assert.That(SearchWildcard(new string('a', 1000), new string('a', 255))).IsEqualTo(0);

        await Assert.That(SearchWildcard(new string('a', 1000), new string('a', 256))).IsEqualTo(-1);
    }

    [Test]
    [Arguments("?", "a", true)]
    [Arguments("?", "ab", false)]
    [Arguments("a?", "ab", true)]
    [Arguments("a?", "abc", false)]
    [Arguments("?b", "ab", true)]
    [Arguments("?b", "aab", false)]
    [Arguments("a*", "abc", true)]
    [Arguments("*a*", "abc", true)]
    [Arguments("*c", "abc", true)]
    [Arguments("*a*a", "abc", false)]
    [Arguments("*a*a", "aba", true)]
    [Arguments("*a*a", "zaba", true)]
    [Arguments("a*", "zaba", false)]
    public async Task Matches(string pattern, string text, bool matches)
    {
        await Assert.That(new Wildcard(pattern).Matches(text.AsSpan())).IsEqualTo(matches);
    }

    private static int SearchWildcard(string text, string pattern)
    {
        return new Wildcard(pattern).Search(text.AsSpan());
    }
}
