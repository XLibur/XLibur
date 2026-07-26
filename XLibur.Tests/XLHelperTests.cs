using XLibur.Excel;
using System;
using System.Threading.Tasks;

namespace XLibur.Tests;
// ReSharper disable once InconsistentNaming
public class XLHelperTests
{
    [Test]
    public async Task IsValidColumnTest()
    {
        await Assert.That(XLHelper.IsValidColumn("")).IsEqualTo(false);
        await Assert.That(XLHelper.IsValidColumn("1")).IsEqualTo(false);
        await Assert.That(XLHelper.IsValidColumn("A1")).IsEqualTo(false);
        await Assert.That(XLHelper.IsValidColumn("AA1")).IsEqualTo(false);
        await Assert.That(XLHelper.IsValidColumn("A")).IsEqualTo(true);
        await Assert.That(XLHelper.IsValidColumn("AA")).IsEqualTo(true);
        await Assert.That(XLHelper.IsValidColumn("AAA")).IsEqualTo(true);
        await Assert.That(XLHelper.IsValidColumn("Z")).IsEqualTo(true);
        await Assert.That(XLHelper.IsValidColumn("ZZ")).IsEqualTo(true);
        await Assert.That(XLHelper.IsValidColumn("XFD")).IsEqualTo(true);
        await Assert.That(XLHelper.IsValidColumn("ZAA")).IsEqualTo(false);
        await Assert.That(XLHelper.IsValidColumn("XZA")).IsEqualTo(false);
        await Assert.That(XLHelper.IsValidColumn("XFZ")).IsEqualTo(false);
    }

    [Test]
    public async Task ReplaceRelative1()
    {
        var result = XLHelper.ReplaceRelative("A1", 2, "B");
        await Assert.That(result).IsEqualTo("B2");
    }

    [Test]
    public async Task ReplaceRelative2()
    {
        var result = XLHelper.ReplaceRelative("$A1", 2, "B");
        await Assert.That(result).IsEqualTo("$A2");
    }

    [Test]
    public async Task ReplaceRelative3()
    {
        var result = XLHelper.ReplaceRelative("A$1", 2, "B");
        await Assert.That(result).IsEqualTo("B$1");
    }

    [Test]
    public async Task ReplaceRelative4()
    {
        var result = XLHelper.ReplaceRelative("$A$1", 2, "B");
        await Assert.That(result).IsEqualTo("$A$1");
    }

    [Test]
    public async Task ReplaceRelative5()
    {
        var result = XLHelper.ReplaceRelative("1:1", 2, "B");
        await Assert.That(result).IsEqualTo("2:2");
    }

    [Test]
    public async Task ReplaceRelative6()
    {
        var result = XLHelper.ReplaceRelative("$1:1", 2, "B");
        await Assert.That(result).IsEqualTo("$1:2");
    }

    [Test]
    public async Task ReplaceRelative7()
    {
        var result = XLHelper.ReplaceRelative("1:$1", 2, "B");
        await Assert.That(result).IsEqualTo("2:$1");
    }

    [Test]
    public async Task ReplaceRelative8()
    {
        var result = XLHelper.ReplaceRelative("$1:$1", 2, "B");
        await Assert.That(result).IsEqualTo("$1:$1");
    }

    [Test]
    public async Task ReplaceRelative9()
    {
        var result = XLHelper.ReplaceRelative("A:A", 2, "B");
        await Assert.That(result).IsEqualTo("B:B");
    }

    [Test]
    public async Task ReplaceRelativeA()
    {
        var result = XLHelper.ReplaceRelative("$A:A", 2, "B");
        await Assert.That(result).IsEqualTo("$A:B");
    }

    [Test]
    public async Task ReplaceRelativeB()
    {
        var result = XLHelper.ReplaceRelative("A:$A", 2, "B");
        await Assert.That(result).IsEqualTo("B:$A");
    }

    [Test]
    public async Task ReplaceRelativeC()
    {
        var result = XLHelper.ReplaceRelative("$A:$A", 2, "B");
        await Assert.That(result).IsEqualTo("$A:$A");
    }

    [Test]
    [Arguments("Sheet1", "Sheet1")]
    [Arguments("O'Brien's sales", "O'Brien's sales")]
    [Arguments(" data # ", " data # ")]
    [Arguments("data $1.00", "data $1.00")]
    [Arguments("data ", "data?")]
    [Arguments("abc def", "abc/def")]
    [Arguments("data 0 ", "data[0]")]
    [Arguments("data ", "data*")]
    [Arguments("abc def", "abc\\def")]
    [Arguments(" data", "'data")]
    [Arguments("data ", "data'")]
    [Arguments("d'at'a", "d'at'a")]
    [Arguments("sheet a4", "sheet:a4")]
    [Arguments("null", null)]
    [Arguments("empty", "")]
    [Arguments("1234567890123456789012345678901", "1234567890123456789012345678901TOOLONG")]
    public async Task CreateSafeSheetNames(string expected, string input)
    {
        var actual = XLHelper.CreateSafeSheetName(input);
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("Sheet1", "Sheet1")]
    [Arguments("O'Brien's sales", "O'Brien's sales")]
    [Arguments(" data # ", " data # ")]
    [Arguments("data $1.00", "data $1.00")]
    [Arguments("data?", "data_")]
    [Arguments("abc/def", "abc_def")]
    [Arguments("data[0]", "data_0_")]
    [Arguments("data*", "data_")]
    [Arguments("abc\\def", "abc_def")]
    [Arguments("'data", "_data")]
    [Arguments("data'", "data_")]
    [Arguments("d'at'a", "d'at'a")]
    [Arguments("sheet:a4", "sheet_a4")]
    [Arguments(null, "null")]
    [Arguments("", "empty")]
    [Arguments("1234567890123456789012345678901TOOLONG", "1234567890123456789012345678901")]
    public async Task CreateSafeSheetNamesWithUnderscore(string input, string expected)
    {
        await Assert.That(XLHelper.CreateSafeSheetName(input, replaceChar: '_')).IsEqualTo(expected);
    }

    [Test]
    public async Task CreateSafeSheetNamesInvalidReplacementChar()
    {
        await Assert.That(() => XLHelper.CreateSafeSheetName("abc\\def", replaceChar: ':')).Throws<ArgumentException>();
    }
}
