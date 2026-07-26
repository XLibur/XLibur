using XLibur.Excel.InsertData;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.InsertData;

public class DataRecordReaderTests
{
    private readonly string _connectionString = @"Data Source=(localdb)\MSSQLLocalDB;Integrated Security=True;Connect Timeout=1";

    private IEnumerable<IDataRecord> GetData()
    {
        const string queryString = @"
            select 'Value 1' as StringValue, 100 as NumericValue
            union all
            select 'Value 2', 200
            union all
            select 'Value 3', 300";

        using var connection = new SqlConnection(_connectionString);
        using var command = new SqlCommand(queryString, connection);
        try
        {
            connection.Open();
        }
        catch
        {
            Skip.Test("Could not connect to localdb");
        }

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            yield return reader;
        }
    }

    [Test]
    public async Task CanGetPropertyName()
    {
        var reader = InsertDataReaderFactory.CreateReader(GetData());
        await Assert.That(reader.GetPropertyName(0)).IsEqualTo("StringValue");
        await Assert.That(reader.GetPropertyName(1)).IsEqualTo("NumericValue");
    }

    [Test]
    public async Task CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(GetData());
        await Assert.That(reader.GetPropertiesCount()).IsEqualTo(2);
    }

    [Test]
    public async Task CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(GetData());
        await Assert.That(reader.GetRecords().Count()).IsEqualTo(3);
    }

    [Test]
    public async Task CanGetData()
    {
        var reader = InsertDataReaderFactory.CreateReader(GetData());
        var result = reader.GetRecords().ToArray();

        await Assert.That(result.First().First()).IsEqualTo("Value 1");
        await Assert.That(result.First().Last()).IsEqualTo(100);
        await Assert.That(result.Last().First()).IsEqualTo("Value 3");
        await Assert.That(result.Last().Last()).IsEqualTo(300);
    }
}
