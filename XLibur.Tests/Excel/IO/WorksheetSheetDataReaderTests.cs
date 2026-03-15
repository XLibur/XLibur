using NUnit.Framework;
using XLibur.Excel;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.IO;

[TestFixture]
public class WorksheetSheetDataReaderTests
{
    [TestCase("yyyy-MM-dd", XLDataType.DateTime)]
    [TestCase("YYYY-MM-DD", XLDataType.DateTime)]
    [TestCase("Yyyy-Mm-Dd", XLDataType.DateTime)]
    [TestCase("hh:mm:ss", XLDataType.TimeSpan)]
    [TestCase("HH:MM:SS", XLDataType.TimeSpan)]
    [TestCase("#,##0.00", XLDataType.Number)]
    [TestCase("0.00%", XLDataType.Number)]
    [TestCase("mm:ss", XLDataType.TimeSpan)]
    [TestCase("MM:SS", XLDataType.TimeSpan)]
    [TestCase("[Red]0.00", XLDataType.Number)]
    [TestCase("\"Date: \"yyyy-MM-dd", XLDataType.DateTime)]
    [TestCase("[$-409]MMMM D, YYYY", XLDataType.DateTime)]
    public void GetDataTypeFromFormat_handles_mixed_case(string format, XLDataType expected)
    {
        var result = WorksheetSheetDataReader.GetDataTypeFromFormat(format);
        Assert.That(result, Is.EqualTo(expected));
    }

    [TestCase("General")]
    [TestCase("@")]
    [TestCase("")]
    public void GetDataTypeFromFormat_returns_null_for_non_numeric_date_formats(string format)
    {
        var result = WorksheetSheetDataReader.GetDataTypeFromFormat(format);
        Assert.That(result, Is.Null);
    }
}
