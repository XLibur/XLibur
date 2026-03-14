using XLibur.Excel.InsertData;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace XLibur.Tests.Excel.InsertData;

public class InsertDataReaderFactoryTests
{
    [Test]
    public void CanInstantiateFactory()
    {
        var factory = InsertDataReaderFactory.Instance;

        Assert.IsNotNull(factory);
        Assert.AreSame(factory, InsertDataReaderFactory.Instance);
    }

    [TestCaseSource(nameof(SimpleSources))]
    public void CanCreateSimpleReader(IEnumerable data)
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(data);

        Assert.IsInstanceOf<SimpleTypeReader>(reader);
    }

    private static IEnumerable<object> SimpleSources
    {
        get
        {
            yield return new[] { 1, 2, 3 };
            yield return new List<double> { 1.0, 2.0, 3.0 };
            yield return new[] { "A", "B", "C" };
            yield return new[] { "A", "B", "C" };
            yield return new[] { 'A', 'B', 'C' };
        }
    }

    [TestCaseSource(nameof(SimpleNullableSources))]
    public void CanCreateSimpleNullableReader(IEnumerable data)
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(data);

        Assert.IsInstanceOf<SimpleNullableTypeReader>(reader);
    }

    private static IEnumerable<object> SimpleNullableSources
    {
        get
        {
            yield return new int?[] { 1, 2, null };
            yield return new List<double?> { 1.0, 2.0, null };
            yield return new char?[] { 'A', 'B', null };
            yield return new DateTime?[] { DateTime.MinValue, DateTime.MaxValue, null };
        }
    }

    [TestCaseSource(nameof(ArraySources))]
    public void CanCreateArrayReader<T>(IEnumerable<T> data)
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(data);

        Assert.IsInstanceOf<ArrayReader>(reader);
    }

    private static IEnumerable<object[]> ArraySources
    {
        get
        {
            yield return
            [
                new int[][]
                {
                    [1, 2, 3],
                    [4, 5, 6]
                }
            ];
            yield return [new List<List<double>> { new List<double> { 1.0, 2.0, 3.0 } }];
            yield return
            [
                (new int[][]
                {
                    [1, 2, 3],
                    [4, 5, 6]
                }).AsEnumerable()
            ];
            yield return
            [
                new[]
                {
                    new decimal[5],
                    new decimal[5],
                }
            ];
        }
    }

    private static readonly int[] SourceArray = [1, 2, 3];
    private static readonly double[] SourceArray0 = [1.0, 2.0, 3.0];

    [Test]
    public void CanCreateArrayReaderFromIEnumerableOfIEnumerables()
    {
        IEnumerable<IEnumerable> data = new List<IEnumerable>
        {
            SourceArray.AsEnumerable(),
            SourceArray0.AsEnumerable(),
        };
        var reader = InsertDataReaderFactory.Instance.CreateReader(data);

        Assert.IsInstanceOf<ArrayReader>(reader);
    }

    [Test]
    public void CanCreateSimpleReaderFromIEnumerableOfString()
    {
        IEnumerable<string> data = new[]
        {
            "String 1",
            "String 2",
        };
        var reader = InsertDataReaderFactory.Instance.CreateReader(data);

        Assert.IsInstanceOf<SimpleTypeReader>(reader);
    }

    [Test]
    public void CanCreateDataTableReader()
    {
        var dt = new DataTable();
        var reader = InsertDataReaderFactory.Instance.CreateReader(dt);

        Assert.IsInstanceOf<XLibur.Excel.InsertData.DataTableReader>(reader);
    }

    [Test]
    public void CanCreateDataRecordReader()
    {
        var dataRecords = Array.Empty<IDataRecord>();
        var reader = InsertDataReaderFactory.Instance.CreateReader(dataRecords);
        Assert.IsInstanceOf<DataRecordReader>(reader);
    }

    [Test]
    public void CanCreateObjectReader()
    {
        var entities = Array.Empty<TestEntity>();
        var reader = InsertDataReaderFactory.Instance.CreateReader(entities);
        Assert.IsInstanceOf<ObjectReader>(reader);
    }

    [Test]
    public void CanCreateObjectReaderForStruct()
    {
        var entities = Array.Empty<TestStruct>();
        var reader = InsertDataReaderFactory.Instance.CreateReader(entities);
        Assert.IsInstanceOf<ObjectReader>(reader);
    }

    [Test]
    public void CanCreateUntypedObjectReader()
    {
        var entities = new ArrayList(new object[]
        {
            new TestEntity(),
            "123",
        });
        var reader = InsertDataReaderFactory.Instance.CreateReader(entities);
        Assert.IsInstanceOf<UntypedObjectReader>(reader);
    }

    private class TestEntity;

    private struct TestStruct;
}