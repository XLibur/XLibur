using System.Collections.Generic;
using XLibur.Excel.InsertData;
using NUnit.Framework;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Tests.Excel.InsertData;

public class ArrayTypeReaderTests
{
    private readonly int[][] _data = new int[][]
    {
        [1, 2, 3],
        [4, 5, 6]
    };

    [Test]
    public void GetPropertyNameReturnsNull()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        Assert.IsNull(reader.GetPropertyName(0));
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        Assert.AreEqual(3, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        Assert.AreEqual(2, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        var result = reader.GetRecords();
        var enumerable = result as IEnumerable<XLCellValue>[] ?? result.ToArray();

        Assert.AreEqual(1, enumerable.First().First());
        Assert.AreEqual(3, enumerable.First().Last());
        Assert.AreEqual(4, enumerable.Last().First());
        Assert.AreEqual(6, enumerable.Last().Last());
    }
}
