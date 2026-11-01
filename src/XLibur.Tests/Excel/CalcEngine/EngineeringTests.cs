using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

[SetCulture("en-US")]
public class EngineeringTests
{
    #region HEX2DEC

    [Test]
    [Arguments("\"A\"", 10)]
    [Arguments("\"FF\"", 255)]
    [Arguments("\"AF0\"", 2800)]
    [Arguments("\"3DA408B9F\"", 16546565023)]
    [Arguments("\"0\"", 0)]
    [Arguments("\"1\"", 1)]
    [Arguments("\"FFFFFFFFFF\"", -1)] // 10 F's = -1 in two's complement
    [Arguments("\"FFFFFFFE00\"", -512)] // Negative via two's complement
    [Arguments("\"8000000000\"", -549755813888)] // Most negative 40-bit value
    [Arguments("\"7FFFFFFFFF\"", 549755813887)] // Most positive 40-bit value
    public async Task Hex2Dec(string input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"HEX2DEC({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"FFFFFFFFFFF\"")] // 11 chars, too long
    [Arguments("\"GG\"")] // Invalid hex char
    public async Task Hex2Dec_InvalidInput_ReturnsNumError(string input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"HEX2DEC({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region DEC2HEX

    [Test]
    [Arguments(100, "\"64\"")]
    [Arguments(0, "\"0\"")]
    [Arguments(-1, "\"FFFFFFFFFF\"")]
    [Arguments(549755813887, "\"7FFFFFFFFF\"")]
    [Arguments(-549755813888, "\"8000000000\"")]
    public async Task Dec2Hex(double input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2HEX({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(expected.Trim('"'));
    }

    [Test]
    [Arguments(100, 4, "0064")]
    [Arguments(10, 5, "0000A")]
    public async Task Dec2Hex_WithPlaces(double input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2HEX({input.ToString(System.Globalization.CultureInfo.InvariantCulture)},{places})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments(100, 1)] // Result is "64" which is 2 chars, but places=1
    public async Task Dec2Hex_PlacesTooSmall_ReturnsNumError(double input, int places)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DEC2HEX({input.ToString(System.Globalization.CultureInfo.InvariantCulture)},{places})")).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region HEX2BIN

    [Test]
    [Arguments("\"F\"", "1111")]
    [Arguments("\"A\"", "1010")]
    [Arguments("\"1\"", "1")]
    [Arguments("\"0\"", "0")]
    [Arguments("\"1FF\"", "111111111")] // 511
    [Arguments("\"FFFFFFFE00\"", "1000000000")] // -512 in two's complement
    public async Task Hex2Bin(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2BIN({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"F\"", 8, "00001111")]
    public async Task Hex2Bin_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2BIN({input},{places})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"200\"")] // 512, exceeds BIN range
    public async Task Hex2Bin_OutOfRange_ReturnsNumError(string input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"HEX2BIN({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region HEX2OCT

    [Test]
    [Arguments("\"F\"", "17")]
    [Arguments("\"3B4E\"", "35516")]
    [Arguments("\"0\"", "0")]
    [Arguments("\"FFFFFFFFFF\"", "7777777777")] // -1
    public async Task Hex2Oct(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2OCT({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"F\"", 4, "0017")]
    public async Task Hex2Oct_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2OCT({input},{places})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    #endregion

    #region BIN2DEC

    [Test]
    [Arguments("\"1010\"", 10)]
    [Arguments("\"0\"", 0)]
    [Arguments("\"1\"", 1)]
    [Arguments("\"111111111\"", 511)] // Max positive
    [Arguments("\"1000000000\"", -512)] // Most negative 10-bit
    [Arguments("\"1111111111\"", -1)] // -1 in two's complement
    public async Task Bin2Dec(string input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"BIN2DEC({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"10000000000\"")] // 11 digits, too long
    [Arguments("\"2\"")] // Invalid binary digit
    public async Task Bin2Dec_InvalidInput_ReturnsNumError(string input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"BIN2DEC({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region BIN2HEX

    [Test]
    [Arguments("\"1010\"", "A")]
    [Arguments("\"11111111\"", "FF")]
    [Arguments("\"1111111111\"", "FFFFFFFFFF")] // -1
    public async Task Bin2Hex(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BIN2HEX({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"1010\"", 4, "000A")]
    public async Task Bin2Hex_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BIN2HEX({input},{places})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    #endregion

    #region BIN2OCT

    [Test]
    [Arguments("\"1010\"", "12")]
    [Arguments("\"0\"", "0")]
    [Arguments("\"1111111111\"", "7777777777")] // -1
    public async Task Bin2Oct(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BIN2OCT({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    #endregion

    #region DEC2BIN

    [Test]
    [Arguments(10, "1010")]
    [Arguments(0, "0")]
    [Arguments(511, "111111111")]
    [Arguments(-512, "1000000000")]
    [Arguments(-1, "1111111111")]
    public async Task Dec2Bin(double input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2BIN({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments(10, 8, "00001010")]
    public async Task Dec2Bin_WithPlaces(double input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2BIN({input.ToString(System.Globalization.CultureInfo.InvariantCulture)},{places})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments(512)] // Out of range
    [Arguments(-513)]
    public async Task Dec2Bin_OutOfRange_ReturnsNumError(double input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DEC2BIN({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})")).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region DEC2OCT

    [Test]
    [Arguments(100, "144")]
    [Arguments(0, "0")]
    [Arguments(-1, "7777777777")]
    public async Task Dec2Oct(double input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2OCT({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    #endregion

    #region OCT2DEC

    [Test]
    [Arguments("\"77\"", 63)]
    [Arguments("\"0\"", 0)]
    [Arguments("\"7777777777\"", -1)]
    [Arguments("\"4000000000\"", -536870912)] // Most negative 30-bit
    [Arguments("\"3777777777\"", 536870911)] // Most positive 30-bit
    public async Task Oct2Dec(string input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"OCT2DEC({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    #endregion

    #region OCT2BIN

    [Test]
    [Arguments("\"12\"", "1010")]
    [Arguments("\"0\"", "0")]
    [Arguments("\"7777777777\"", "1111111111")] // -1
    public async Task Oct2Bin(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"OCT2BIN({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"1000\"")] // 512, out of BIN range
    public async Task Oct2Bin_OutOfRange_ReturnsNumError(string input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"OCT2BIN({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    #endregion

    #region OCT2HEX

    [Test]
    [Arguments("\"17\"", "F")]
    [Arguments("\"0\"", "0")]
    [Arguments("\"7777777777\"", "FFFFFFFFFF")] // -1
    public async Task Oct2Hex(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"OCT2HEX({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("\"17\"", 4, "000F")]
    public async Task Oct2Hex_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"OCT2HEX({input},{places})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    #endregion
}
