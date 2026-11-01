using PerfReport;

namespace PerfReport.Tests;

public class BdnJsonReaderTests
{
    // Trimmed from a BenchmarkDotNet 0.15 --exporters json file. Unknown properties must be ignored.
    private const string Json = """
        {
          "Title": "XLibur.Benchmarks.FirstEditAfterLoadBenchmarks-20261102-120000",
          "HostEnvironmentInfo": {
            "BenchmarkDotNetVersion": "0.15.8",
            "OsVersion": "ubuntu 24.04",
            "LogicalCoreCount": 4,
            "RuntimeVersion": ".NET 10.0.0"
          },
          "Benchmarks": [
            {
              "DisplayInfo": "FirstEditAfterLoadBenchmarks.Load: ShortRun(...) [Shape=1000x2]",
              "Namespace": "XLibur.Benchmarks",
              "Type": "FirstEditAfterLoadBenchmarks",
              "Method": "Load",
              "MethodTitle": "Load",
              "Parameters": "Shape=1000x2",
              "Statistics": { "Mean": 1234.5, "Median": 1200.0, "StandardDeviation": 3.0 },
              "Memory": {
                "Gen0Collections": 20,
                "Gen1Collections": 10,
                "Gen2Collections": 0,
                "TotalOperations": 1000,
                "BytesAllocatedPerOperation": 4096
              }
            },
            {
              "Namespace": "XLibur.Benchmarks",
              "Type": "StreamingWriteBenchmarks",
              "Method": "Write",
              "Parameters": "",
              "Statistics": { "Mean": 50.0 },
              "Memory": {
                "Gen0Collections": 0, "Gen1Collections": 0, "Gen2Collections": 0,
                "TotalOperations": 10, "BytesAllocatedPerOperation": 10
              }
            }
          ]
        }
        """;

    [Test]
    public async Task ReadsMeanAllocationAndCollectionsPerThousandOperations()
    {
        var run = BdnJsonReader.Read(Json, type => "hash-" + type);

        var entry = run.Benchmarks[0];
        await Assert.That(run.Benchmarks.Count).IsEqualTo(2);
        await Assert.That(entry.Key).IsEqualTo("FirstEditAfterLoadBenchmarks.Load(Shape=1000x2)");
        await Assert.That(entry.Method).IsEqualTo("Load");
        await Assert.That(entry.Parameters).IsEqualTo("Shape=1000x2");
        await Assert.That(entry.MeanNs).IsEqualTo(1234.5);
        await Assert.That(entry.AllocatedBytes).IsEqualTo(4096L);
        await Assert.That(entry.Gen0).IsEqualTo(20.0);
        await Assert.That(entry.Gen1).IsEqualTo(10.0);
        await Assert.That(entry.SourceHash).IsEqualTo("hash-FirstEditAfterLoadBenchmarks");
    }

    [Test]
    public async Task BenchmarkWithoutParametersHasPlainKey()
    {
        var run = BdnJsonReader.Read(Json, _ => "");

        await Assert.That(run.Benchmarks[1].Key).IsEqualTo("StreamingWriteBenchmarks.Write");
    }

    [Test]
    public async Task ReadsLogicalCoreCountFromHostInfo()
    {
        var run = BdnJsonReader.Read(Json, _ => "");

        await Assert.That(run.LogicalCores).IsEqualTo(4);
    }

    [Test]
    public async Task BenchmarkWithoutMemoryDiagnoserDataIsRejected()
    {
        const string noMemory = """{ "Benchmarks": [ { "Type": "T", "Method": "M", "Statistics": { "Mean": 1 } } ] }""";

        await Assert.That(() => BdnJsonReader.Read(noMemory, _ => ""))
            .Throws<InvalidDataException>();
    }
}
