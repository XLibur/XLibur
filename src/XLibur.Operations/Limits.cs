namespace XLibur.Operations;

public sealed record Limits
{
    public static Limits Default { get; } = new();

    public long MaxInputBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>The most bytes the entries of one .xlsx may inflate to, all entries together.</summary>
    public long MaxUncompressedBytes { get; init; } = 1024L * 1024 * 1024;

    /// <summary>The highest inflated-to-compressed ratio allowed for an entry larger than 1 MiB.</summary>
    public int MaxCompressionRatio { get; init; } = 1000;

    public int MaxZipEntries { get; init; } = 10_000;

    /// <summary>The most cells one read_range call returns, whatever the request asks for.</summary>
    public int MaxCells { get; init; } = 10_000;

    public int MaxResultChars { get; init; } = 1_000_000;
}
