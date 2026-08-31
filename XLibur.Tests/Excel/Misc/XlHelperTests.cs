using System;
using System.Collections.Generic;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class XlHelperTests
{
    private static async Task CheckColumnNumber(int column)
    {
        await Assert.That(XLHelper.GetColumnNumberFromLetter(XLHelper.GetColumnLetterFromNumber(column))).IsEqualTo(column);
    }

    [Test]
    public async Task InvalidA1Addresses()
    {
        await Assert.That(XLHelper.IsValidA1Address("")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("a")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("-1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAAA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("XFG1")).IsFalse();

        await Assert.That(XLHelper.IsValidA1Address("@A1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AAA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[A1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AAA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{A1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AAA1")).IsFalse();

        // A '$' is an anchor and is only meaningful before the column letters or before the row
        // digits. These were all reported valid while the implementation stripped every '$' and
        // validated what was left, which put IsValidA1Address in direct contradiction with
        // IsValidRangeAddress about the same string (D36).
        await Assert.That(XLHelper.IsValidA1Address("A$2$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A2$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$A2$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A$$2")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$$A$$2")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$$")).IsFalse();

        await Assert.That(XLHelper.IsValidA1Address("A1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AA1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAA1{")).IsFalse();

        await Assert.That(XLHelper.IsValidA1Address("@A1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AAA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[A1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AAA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{A1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AA1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AAA1{")).IsFalse();
    }

    [Test]
    public async Task PlusAA1_Is_Not_an_address()
    {
        await Assert.That(XLHelper.IsValidA1Address("+AA1")).IsFalse();
    }

    [Test]
    public async Task TestConvertColumnLetterToNumberAnd()
    {
        await CheckColumnNumber(1);
        await CheckColumnNumber(27);
        await CheckColumnNumber(28);
        await CheckColumnNumber(52);
        await CheckColumnNumber(53);
        await CheckColumnNumber(1000);
        await CheckColumnNumber(1353);
    }

    [Test]
    public async Task ValidA1Addresses()
    {
        await Assert.That(XLHelper.IsValidA1Address("A1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("A" + XLHelper.MaxRowNumber)).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("Z1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("Z" + XLHelper.MaxRowNumber)).IsTrue();

        await Assert.That(XLHelper.IsValidA1Address("AA1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("AA" + XLHelper.MaxRowNumber)).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("ZZ1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("ZZ" + XLHelper.MaxRowNumber)).IsTrue();

        await Assert.That(XLHelper.IsValidA1Address("AAA1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("AAA" + XLHelper.MaxRowNumber)).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + "1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + XLHelper.MaxRowNumber)).IsTrue();

        // The four anchor placements Excel actually allows. Tightening the rejection of a
        // misplaced '$' must not cost the legitimate ones.
        await Assert.That(XLHelper.IsValidA1Address("$A$1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("$A1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("A$1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("$AAA$" + XLHelper.MaxRowNumber)).IsTrue();
    }

    [Test]
    public async Task TestColumnLetterLookup()
    {
        var columnLetters = new List<string>();
        for (var c = 1; c <= XLHelper.MaxColumnNumber; c++)
        {
            var columnLetter = NaiveGetColumnLetterFromNumber(c);
            columnLetters.Add(columnLetter);

            await Assert.That(XLHelper.GetColumnLetterFromNumber(c)).IsEqualTo(columnLetter);
        }

        foreach (var cl in columnLetters)
        {
            var columnNumber = NaiveGetColumnNumberFromLetter(cl);
            await Assert.That(XLHelper.GetColumnNumberFromLetter(cl)).IsEqualTo(columnNumber);
        }
    }

    [Test]
    [Arguments("R")]
    [Arguments("C")]
    [Arguments("RC")]
    [Arguments("R111C222")]
    [Arguments("R[]C")]
    [Arguments("RC[]")]
    [Arguments("R[]C[]")]
    [Arguments("R[111]C222")]
    [Arguments("R111C[222]")]
    [Arguments("R[111]C[222]")]
    [Arguments("R[-111]C[-222]")]
    public async Task ValidRCAddresses(string address)
    {
        await Assert.That(XLHelper.IsValidRCAddress(address)).IsTrue();
    }

    [Test]
    [Arguments("RD")]
    [Arguments("CC")]
    [Arguments("R[-]C222")]
    [Arguments("R[]C[-]")]
    [Arguments("_R111C222")]
    public async Task InvalidRCAddresses(string address)
    {
        await Assert.That(XLHelper.IsValidRCAddress(address)).IsFalse();
    }

    #region Old XLHelper methods

    private static readonly string[] Letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

    /// <summary>
    /// These used to be the methods in XLHelper, but were later changed.
    /// We now use them as a check against the new methods
    /// Gets the column number of a given column letter.
    /// </summary>
    /// <param name="columnLetter"> The column letter to translate into a column number. </param>
    private static int NaiveGetColumnNumberFromLetter(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentNullException(nameof(columnLetter));

        columnLetter = columnLetter.ToUpper();

        //Extra check because we allow users to pass row col positions in as strings
        if (columnLetter[0] <= '9')
        {
            var retVal = int.Parse(columnLetter, XLHelper.NumberStyle, XLHelper.ParseCulture);
            return retVal;
        }

        var sum = 0;

        foreach (var t in columnLetter)
        {
            sum *= 26;
            sum += t - 'A' + 1;
        }

        return sum;
    }

    /// <summary>
    /// Gets the column letter of a given column number.
    /// </summary>
    /// <param name="columnNumber">The column number to translate into a column letter.</param>
    /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
    private static string NaiveGetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
    {
        if (trimToAllowed) columnNumber = XLHelper.TrimColumnNumber(columnNumber);

        columnNumber--; // Adjust for start on column 1
        if (columnNumber <= 25)
        {
            return Letters[columnNumber];
        }
        var firstPart = (columnNumber) / 26;
        var remainder = ((columnNumber) % 26) + 1;
        return NaiveGetColumnLetterFromNumber(firstPart) + NaiveGetColumnLetterFromNumber(remainder);
    }

    #endregion Old XLHelper methods
}
