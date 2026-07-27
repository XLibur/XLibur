using System.IO.Compression;

namespace XLibur.Excel.Streaming;

/// <summary>
/// How the streaming writer stores text values.
/// </summary>
public enum XLStreamingStringStorage
{
    /// <summary>
    /// Text is written to the workbook-wide shared string table and cells reference it by
    /// index. Produces the smallest file when text repeats, which is the usual case for an
    /// export. The dictionary of distinct strings is held in memory until
    /// <see cref="XLStreamingWorkbook.Finish"/>, so it is the one part of a streaming write
    /// whose cost grows with the data.
    /// </summary>
    SharedStrings = 0,

    /// <summary>
    /// Text is written into the cell as an inline string. Uses no memory beyond the current
    /// row at the cost of a larger file when text repeats. Choose this when the number of
    /// distinct strings is large enough that the shared string dictionary would not fit.
    /// </summary>
    Inline = 1
}

/// <summary>
/// Options for <see cref="XLStreamingWorkbook.Create(System.IO.Stream, XLStreamingOptions?)"/>.
/// </summary>
public sealed class XLStreamingOptions
{
    /// <summary>
    /// How text values are stored. Defaults to
    /// <see cref="XLStreamingStringStorage.SharedStrings"/>.
    /// </summary>
    public XLStreamingStringStorage StringStorage { get; set; } = XLStreamingStringStorage.SharedStrings;

    /// <summary>
    /// Write dates against the 1904 date system rather than the 1900 one. Defaults to
    /// <c>false</c>, matching <see cref="XLWorkbook"/>.
    /// </summary>
    public bool Use1904DateSystem { get; set; }

    /// <summary>
    /// How hard to compress the package. Defaults to
    /// <see cref="System.IO.Compression.CompressionLevel.Optimal"/>;
    /// <see cref="System.IO.Compression.CompressionLevel.Fastest"/> trades a noticeably larger
    /// file for a faster write, which is often the right call for a large export.
    /// </summary>
    /// <remarks>
    /// <see cref="SaveOptions.CompressionLevel"/> is the equivalent for an ordinary save. The two
    /// reach it differently: the streaming writer configures its own zip, while
    /// <see cref="XLWorkbook"/> passes the setting through <c>System.IO.Packaging</c>, which
    /// applies it only to parts that save creates.
    /// </remarks>
    public CompressionLevel CompressionLevel { get; set; } = CompressionLevel.Optimal;
}
