using XLibur.Excel.InsertData;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class DataRowReaderTests
{
    private readonly DataTable _data;

    public DataRowReaderTests()
    {
        _data = new DataTable();
        _data.Columns.Add("Last name");
        _data.Columns.Add("First name");
        _data.Columns.Add("Age", typeof(int));

        _data.Rows.Add("Smith", "John", 33);
        _data.Rows.Add("Ivanova", "Olga", 25);
    }

    [Test]
    public async Task CanGetPropertyName()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetPropertyName(0)).IsEqualTo("Last name");
        await Assert.That(reader.GetPropertyName(1)).IsEqualTo("First name");
        await Assert.That(reader.GetPropertyName(2)).IsEqualTo("Age");
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
    public async Task CanReadValue()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        await Assert.That(enumerable.First().First()).IsEqualTo("Smith");
        await Assert.That(enumerable.First().Last()).IsEqualTo(33);
        await Assert.That(enumerable.Last().First()).IsEqualTo("Ivanova");
        await Assert.That(enumerable.Last().Last()).IsEqualTo(25);
    }
}
