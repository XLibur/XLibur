using PerfReport;

namespace PerfReport.Tests;

public class SourceHasherTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "hasher-" + Guid.NewGuid().ToString("N"));

    public SourceHasherTests() => Directory.CreateDirectory(_dir);

    public void Dispose() => Directory.Delete(_dir, recursive: true);

    private void Write(string name, string text) => File.WriteAllText(Path.Combine(_dir, name + ".cs"), text);

    private string Hash(string type) => new SourceHasher(_dir).HashFor(type);

    private const string Bench = "// run it\npublic class Bench { [Benchmark] public void Run() => Fixture.Build(); }\n";

    [Test]
    public async Task IgnoresLineEndingsCommentsAndWhitespace()
    {
        Write("Bench", Bench);
        var before = Hash("Bench");

        Write("Bench", "/* other */\r\npublic class Bench {\r\n  [Benchmark]   public void Run()\r\n => Fixture.Build(); }\r\n");

        await Assert.That(Hash("Bench")).IsEqualTo(before);
    }

    [Test]
    public async Task ACodeChangeChangesTheHash()
    {
        Write("Bench", Bench);
        var before = Hash("Bench");

        Write("Bench", Bench.Replace("Build()", "Build(2)"));

        await Assert.That(Hash("Bench")).IsNotEqualTo(before);
    }

    [Test]
    public async Task AChangeToAFixtureTheClassNamesChangesTheHash()
    {
        Write("Bench", Bench);
        Write("Fixture", "static class Fixture { public static int Build() => 1; }");
        var before = Hash("Bench");

        Write("Fixture", "static class Fixture { public static int Build() => 2; }");

        await Assert.That(Hash("Bench")).IsNotEqualTo(before);
    }

    [Test]
    public async Task AFixtureOfAFixtureCountsToo()
    {
        Write("Bench", Bench);
        Write("Fixture", "static class Fixture { public static int Build() => Data.Rows; }");
        Write("Data", "static class Data { public const int Rows = 1; }");
        var before = Hash("Bench");

        Write("Data", "static class Data { public const int Rows = 2; }");

        await Assert.That(Hash("Bench")).IsNotEqualTo(before);
    }

    [Test]
    public async Task AChangeToAnUnrelatedFileOrAnotherBenchmarkClassDoesNotChangeTheHash()
    {
        Write("Bench", Bench);
        Write("Other", "public class Other { [Benchmark] public void Run() => Bench.Nothing(); }");
        Write("Unrelated", "static class Unrelated { public static int X = 1; }");
        var before = Hash("Bench");

        Write("Other", "public class Other { [Benchmark] public void Run() => Bench.Something(); }");
        Write("Unrelated", "static class Unrelated { public static int X = 2; }");

        await Assert.That(Hash("Bench")).IsEqualTo(before);
    }

    [Test]
    public async Task ATypeWithNoFileIsUnknown()
    {
        await Assert.That(Hash("Missing")).IsEqualTo("unknown");
    }
}
