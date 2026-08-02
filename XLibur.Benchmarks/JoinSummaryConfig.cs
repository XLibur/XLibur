using BenchmarkDotNet.Configs;

namespace XLibur.Benchmarks;

/// <summary>
/// Emits one joined table for a run instead of a separate summary per benchmark class,
/// so results from the different XLibur benchmark classes can be read side by side.
/// </summary>
public class JoinSummaryConfig : ManualConfig
{
    public JoinSummaryConfig()
    {
        WithOption(ConfigOptions.JoinSummary, true);
    }
}
