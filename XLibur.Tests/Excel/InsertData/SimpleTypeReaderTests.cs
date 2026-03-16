using XLibur.Excel.InsertData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Tests.Excel.InsertData;

public class SimpleTypeReaderTests
{
    private static readonly int[] IntData = [1, 2, 3];
    private static readonly List<double> DoubleData = [1.0, 2.0, 3.0];
    private static readonly decimal[] DecimalData = [1.0m, 2.0m, 3.0m];
    private static readonly string[] StringData = ["A", "B", "C"];
    private static readonly char[] CharData = ['A', 'B', 'C'];
    private static readonly DateTime[] DateTimeData = [new(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified)];

    private readonly int[] _data = [1, 2, 3];

    [TestCaseSource(nameof(SimpleSourceNames))]
    public string CanGetPropertyName<T>(IEnumerable<T> data)
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(data);
        return reader.GetPropertyName(0);
    }

    private static IEnumerable<TestCaseData> SimpleSourceNames
    {
        get
        {
            yield return new TestCaseData(IntData).Returns("Int32");
            yield return new TestCaseData(DoubleData).Returns("Double");
            yield return new TestCaseData(DecimalData).Returns("Decimal");
            yield return new TestCaseData(arg: StringData).Returns("String");
            yield return new TestCaseData(CharData).Returns("Char");
            yield return new TestCaseData(DateTimeData).Returns("DateTime");
        }
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
        Assert.AreEqual(1, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
        Assert.AreEqual(3, reader.GetRecords().Count());
    }

    [Test]
    public void CanReadValues()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(_data);
        var result = reader.GetRecords();

        var enumerable = result.ToList();
        Assert.AreEqual(1, enumerable.First().Single());
        Assert.AreEqual(3, enumerable.Last().Single());
    }
}
