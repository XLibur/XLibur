using XLibur.Excel;
using NUnit.Framework;

namespace XLibur.Tests.Excel.CalcEngine;

[TestFixture]
[SetCulture("en-US")]
public class EngineeringTests
{
    #region HEX2DEC

    [TestCase("\"A\"", 10)]
    [TestCase("\"FF\"", 255)]
    [TestCase("\"AF0\"", 2800)]
    [TestCase("\"3DA408B9F\"", 16546565023)]
    [TestCase("\"0\"", 0)]
    [TestCase("\"1\"", 1)]
    [TestCase("\"FFFFFFFFFF\"", -1)] // 10 F's = -1 in two's complement
    [TestCase("\"FFFFFFFE00\"", -512)] // Negative via two's complement
    [TestCase("\"8000000000\"", -549755813888)] // Most negative 40-bit value
    [TestCase("\"7FFFFFFFFF\"", 549755813887)] // Most positive 40-bit value
    public void Hex2Dec(string input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"HEX2DEC({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"FFFFFFFFFFF\"")] // 11 chars, too long
    [TestCase("\"GG\"")] // Invalid hex char
    public void Hex2Dec_InvalidInput_ReturnsNumError(string input)
    {
        Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"HEX2DEC({input})"));
    }

    #endregion

    #region DEC2HEX

    [TestCase(100, "\"64\"")]
    [TestCase(0, "\"0\"")]
    [TestCase(-1, "\"FFFFFFFFFF\"")]
    [TestCase(549755813887, "\"7FFFFFFFFF\"")]
    [TestCase(-549755813888, "\"8000000000\"")]
    public void Dec2Hex(double input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2HEX({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})");
        Assert.AreEqual(expected.Trim('"'), actual);
    }

    [TestCase(100, 4, "0064")]
    [TestCase(10, 5, "0000A")]
    public void Dec2Hex_WithPlaces(double input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2HEX({input.ToString(System.Globalization.CultureInfo.InvariantCulture)},{places})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase(100, 1)] // Result is "64" which is 2 chars, but places=1
    public void Dec2Hex_PlacesTooSmall_ReturnsNumError(double input, int places)
    {
        Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"DEC2HEX({input.ToString(System.Globalization.CultureInfo.InvariantCulture)},{places})"));
    }

    #endregion

    #region HEX2BIN

    [TestCase("\"F\"", "1111")]
    [TestCase("\"A\"", "1010")]
    [TestCase("\"1\"", "1")]
    [TestCase("\"0\"", "0")]
    [TestCase("\"1FF\"", "111111111")] // 511
    [TestCase("\"FFFFFFFE00\"", "1000000000")] // -512 in two's complement
    public void Hex2Bin(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2BIN({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"F\"", 8, "00001111")]
    public void Hex2Bin_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2BIN({input},{places})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"200\"")] // 512, exceeds BIN range
    public void Hex2Bin_OutOfRange_ReturnsNumError(string input)
    {
        Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"HEX2BIN({input})"));
    }

    #endregion

    #region HEX2OCT

    [TestCase("\"F\"", "17")]
    [TestCase("\"3B4E\"", "35516")]
    [TestCase("\"0\"", "0")]
    [TestCase("\"FFFFFFFFFF\"", "7777777777")] // -1
    public void Hex2Oct(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2OCT({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"F\"", 4, "0017")]
    public void Hex2Oct_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"HEX2OCT({input},{places})");
        Assert.AreEqual(expected, actual);
    }

    #endregion

    #region BIN2DEC

    [TestCase("\"1010\"", 10)]
    [TestCase("\"0\"", 0)]
    [TestCase("\"1\"", 1)]
    [TestCase("\"111111111\"", 511)] // Max positive
    [TestCase("\"1000000000\"", -512)] // Most negative 10-bit
    [TestCase("\"1111111111\"", -1)] // -1 in two's complement
    public void Bin2Dec(string input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"BIN2DEC({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"10000000000\"")] // 11 digits, too long
    [TestCase("\"2\"")] // Invalid binary digit
    public void Bin2Dec_InvalidInput_ReturnsNumError(string input)
    {
        Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"BIN2DEC({input})"));
    }

    #endregion

    #region BIN2HEX

    [TestCase("\"1010\"", "A")]
    [TestCase("\"11111111\"", "FF")]
    [TestCase("\"1111111111\"", "FFFFFFFFFF")] // -1
    public void Bin2Hex(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BIN2HEX({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"1010\"", 4, "000A")]
    public void Bin2Hex_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BIN2HEX({input},{places})");
        Assert.AreEqual(expected, actual);
    }

    #endregion

    #region BIN2OCT

    [TestCase("\"1010\"", "12")]
    [TestCase("\"0\"", "0")]
    [TestCase("\"1111111111\"", "7777777777")] // -1
    public void Bin2Oct(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BIN2OCT({input})");
        Assert.AreEqual(expected, actual);
    }

    #endregion

    #region DEC2BIN

    [TestCase(10, "1010")]
    [TestCase(0, "0")]
    [TestCase(511, "111111111")]
    [TestCase(-512, "1000000000")]
    [TestCase(-1, "1111111111")]
    public void Dec2Bin(double input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2BIN({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase(10, 8, "00001010")]
    public void Dec2Bin_WithPlaces(double input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2BIN({input.ToString(System.Globalization.CultureInfo.InvariantCulture)},{places})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase(512)] // Out of range
    [TestCase(-513)]
    public void Dec2Bin_OutOfRange_ReturnsNumError(double input)
    {
        Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"DEC2BIN({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})"));
    }

    #endregion

    #region DEC2OCT

    [TestCase(100, "144")]
    [TestCase(0, "0")]
    [TestCase(-1, "7777777777")]
    public void Dec2Oct(double input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"DEC2OCT({input.ToString(System.Globalization.CultureInfo.InvariantCulture)})");
        Assert.AreEqual(expected, actual);
    }

    #endregion

    #region OCT2DEC

    [TestCase("\"77\"", 63)]
    [TestCase("\"0\"", 0)]
    [TestCase("\"7777777777\"", -1)]
    [TestCase("\"4000000000\"", -536870912)] // Most negative 30-bit
    [TestCase("\"3777777777\"", 536870911)] // Most positive 30-bit
    public void Oct2Dec(string input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"OCT2DEC({input})");
        Assert.AreEqual(expected, actual);
    }

    #endregion

    #region OCT2BIN

    [TestCase("\"12\"", "1010")]
    [TestCase("\"0\"", "0")]
    [TestCase("\"7777777777\"", "1111111111")] // -1
    public void Oct2Bin(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"OCT2BIN({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"1000\"")] // 512, out of BIN range
    public void Oct2Bin_OutOfRange_ReturnsNumError(string input)
    {
        Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"OCT2BIN({input})"));
    }

    #endregion

    #region OCT2HEX

    [TestCase("\"17\"", "F")]
    [TestCase("\"0\"", "0")]
    [TestCase("\"7777777777\"", "FFFFFFFFFF")] // -1
    public void Oct2Hex(string input, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"OCT2HEX({input})");
        Assert.AreEqual(expected, actual);
    }

    [TestCase("\"17\"", 4, "000F")]
    public void Oct2Hex_WithPlaces(string input, int places, string expected)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"OCT2HEX({input},{places})");
        Assert.AreEqual(expected, actual);
    }

    #endregion
}
