using XLibur.Excel.InsertData;
using System.Collections;
using System.Linq;
using XLibur.Excel;
using XLibur.Tests.Excel.Tables;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;

namespace XLibur.Tests.Excel.InsertData;

public class UntypedObjectReaderTests
{
    private readonly ArrayList _data = new(new object[]
    {
        null!,
        new TablesTests.TestObjectWithAttributes
        {
            Column1 = "Value 1",
            Column2 = "Value 2",
            UnOrderedColumn = 3,
            MyField = 4,
        },
        null!,
        null!,
        null!,
        new[] { 1, 2, 3 },
        new[] { 4, 5, 6, 7 },
        "Separator",

        new TablesTests.TestObjectWithoutAttributes
        {
            Column1 = "Value 9",
            Column2 = "Value 10"
        },
    });

    [Test]
    [Arguments(0, "FirstColumn")]
    [Arguments(1, "SecondColumn")]
    [Arguments(2, "SomeFieldNotProperty")]
    [Arguments(3, "UnOrderedColumn")]
    public async Task CanGetPropertyName(int propertyIndex, string expectedPropertyName)
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        var actualPropertyName = reader.GetPropertyName(propertyIndex);
        await Assert.That(actualPropertyName).IsEqualTo(expectedPropertyName);
    }

    [Test]
    public async Task CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetPropertiesCount()).IsEqualTo(4);
    }

    [Test]
    public async Task CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);
        await Assert.That(reader.GetRecords().Count()).IsEqualTo(9);
    }

    [Test]
    public async Task CanGetData()
    {
        var reader = InsertDataReaderFactory.CreateReader(_data);

        var result = reader.GetRecords().ToArray();

        await Assert.That(result[0]).IsEquivalentTo(new XLCellValue[] { Blank.Value }, CollectionOrdering.Matching);
        await Assert.That(result[1]).IsEquivalentTo(new XLCellValue[] { "Value 2", "Value 1", 4, 3 }, CollectionOrdering.Matching);
        await Assert.That(result[2]).IsEquivalentTo(new XLCellValue[] { Blank.Value }, CollectionOrdering.Matching);
        await Assert.That(result[3]).IsEquivalentTo(new XLCellValue[] { Blank.Value }, CollectionOrdering.Matching);
        await Assert.That(result[4]).IsEquivalentTo(new XLCellValue[] { Blank.Value }, CollectionOrdering.Matching);
        await Assert.That(result[5]).IsEquivalentTo(new XLCellValue[] { 1, 2, 3 }, CollectionOrdering.Matching);
        await Assert.That(result[6]).IsEquivalentTo(new XLCellValue[] { 4, 5, 6, 7 }, CollectionOrdering.Matching);
        await Assert.That(result[7]).IsEquivalentTo(new XLCellValue[] { "Separator" }, CollectionOrdering.Matching);
        await Assert.That(result[8]).IsEquivalentTo(new XLCellValue[] { "Value 9", "Value 10" }, CollectionOrdering.Matching);
    }
}
