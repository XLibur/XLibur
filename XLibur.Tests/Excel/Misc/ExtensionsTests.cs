using System;
using DocumentFormat.OpenXml;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class ExtensionsTests
{
    [Test]
    public async Task FixNewLines()
    {
        await Assert.That("\n".FixNewLines()).IsEqualTo(Environment.NewLine);
        await Assert.That("\r\n".FixNewLines()).IsEqualTo(Environment.NewLine);
        await Assert.That("\rS\n".FixNewLines()).IsEqualTo("\rS" + Environment.NewLine);
        await Assert.That("\r\n\n".FixNewLines()).IsEqualTo(Environment.NewLine + Environment.NewLine);
    }

    [Test]
    public async Task DoubleSaveRound()
    {
        const double value = 1234.1234567;
        await Assert.That(Math.Round(value, 6)).IsEqualTo(value.SaveRound());
    }

    [Test]
    public async Task DoubleValueSaveRound()
    {
        const double value = 1234.1234567;
        await Assert.That(Math.Round(value, 6)).IsEqualTo(new DoubleValue(value).SaveRound().Value);
    }

    [Test]
    [Arguments("NoEscaping", "NoEscaping")]
    [Arguments("1", "'1'")]
    [Arguments("AB-CD", "'AB-CD'")]
    [Arguments(" AB", "' AB'")]
    [Arguments("Test sheet", "'Test sheet'")]
    [Arguments("O'Kelly", "'O''Kelly'")]
    [Arguments("A2+3", "'A2+3'")]
    [Arguments("A\"B", "'A\"B'")]
    [Arguments("A!B", "'A!B'")]
    [Arguments("A~B", "'A~B'")]
    [Arguments("A^B", "'A^B'")]
    [Arguments("A&B", "'A&B'")]
    [Arguments("A>B", "'A>B'")]
    [Arguments("A<B", "'A<B'")]
    [Arguments("A.B", "A.B")]
    [Arguments(".", "'.'")]
    [Arguments("A_B", "A_B")]
    [Arguments("_", "_")]
    [Arguments("=", "'='")]
    [Arguments("A,B", "'A,B'")]
    [Arguments("A@B", "'A@B'")]
    [Arguments("(Test)", "'(Test)'")]
    [Arguments("A#", "'A#'")]
    [Arguments("A$", "'A$'")]
    [Arguments("A%", "'A%'")]
    [Arguments("ABC1", "'ABC1'")]
    [Arguments("ABCD1", "ABCD1")]
    [Arguments("C05A", "'C05A'")]
    [Arguments("A1B", "'A1B'")]
    [Arguments("XFD1X", "'XFD1X'")]
    [Arguments("XFE1", "XFE1")]
    [Arguments("R1C1", "'R1C1'")]
    [Arguments("A{", "'A{'")]
    [Arguments("A}", "'A}'")]
    [Arguments("A`", "'A`'")]
    [Arguments("Русский", "Русский")]
    [Arguments("日本語", "日本語")]
    [Arguments("한국어", "한국어")]
    [Arguments("Slovenščina", "Slovenščina")]
    [Arguments("", "")]
    [Arguments(null, null)]
    public async Task CanEscapeSheetName(string sheetName, string expected)
    {
        await Assert.That(sheetName.EscapeSheetName()).IsEqualTo(expected);
    }

    [Test]
    [Arguments("TestSheet", "TestSheet")]
    [Arguments("'Test sheet'", "Test sheet")]
    [Arguments("'O''Kelly'", "O'Kelly")]
    public async Task CanUnescapeSheetName(string sheetName, string expected)
    {
        await Assert.That(sheetName.UnescapeSheetName()).IsEqualTo(expected);
    }
}
