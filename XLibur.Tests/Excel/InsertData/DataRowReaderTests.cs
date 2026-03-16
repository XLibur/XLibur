using XLibur.Excel.InsertData;
using NUnit.Framework;
using System.Data;
using System.Linq;

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
    public void CanGetPropertyName()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        Assert.AreEqual("Last name", reader.GetPropertyName(0));
        Assert.AreEqual("First name", reader.GetPropertyName(1));
        Assert.AreEqual("Age", reader.GetPropertyName(2));
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
    public void CanReadValue()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        Assert.AreEqual("Smith", enumerable.First().First());
        Assert.AreEqual(33, enumerable.First().Last());
        Assert.AreEqual("Ivanova", enumerable.Last().First());
        Assert.AreEqual(25, enumerable.Last().Last());
    }
}
