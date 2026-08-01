using System.Collections.Generic;
using XLibur.Excel.InsertData;
using System.Linq;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class ArrayTypeReaderTests
{
    private readonly int[][] _data = new int[][]
    {
        [1, 2, 3],
        [4, 5, 6]
    };

    [Test]
    public async Task GetPropertyNameReturnsNull()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetPropertyName(0)).IsNull();
    }

    [Test]
    public async Task CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetPropertiesCount()).IsEqualTo(3);
    }

    [Test]
    public async Task CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetRecords().Count()).IsEqualTo(2);
    }

    [Test]
    public async Task CanReadValues()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        var result = reader.GetRecords();
        var enumerable = result as IEnumerable<XLCellValue>[] ?? result.ToArray();

        await Assert.That(enumerable.First().First()).IsEqualTo(1);
        await Assert.That(enumerable.First().Last()).IsEqualTo(3);
        await Assert.That(enumerable.Last().First()).IsEqualTo(4);
        await Assert.That(enumerable.Last().Last()).IsEqualTo(6);
    }
}
