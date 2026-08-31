using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.IO;

namespace XLibur.Fuzz;

/// <summary>
/// Decides whether what a target just did was acceptable.
///
/// The rule that shapes everything here is that <em>the phase matters</em>. Handing XLibur a
/// corrupt file and having it refuse is correct behaviour. Handing XLibur a file it already
/// accepted and having it fail to write that file back out is not — by the time the save runs,
/// XLibur has claimed to understand the input, so almost nothing it throws there is excusable.
/// A single flat allowlist cannot express that difference, and the one this replaced discarded
/// write-path defects as though they were malformed-input rejections.
///
/// <para>
/// <b>No rule here may inspect <see cref="Exception.Message"/>.</b> The allowlist this replaced
/// tested for the text "does not exist in the package"; the message actually produced says
/// "doesn't exist in the package", so the branch never fired and a case the harness was written
/// to ignore was reported as a crash on every run for a week. Where a rejection cannot be
/// recognised by its type, the fix is to give it a type in the library — not to pattern-match
/// prose that belongs to somebody else's assembly and can change without notice.
/// </para>
/// </summary>
internal static class Oracle
{
    /// <summary>
    /// Exception types that mean "this input was not a workbook", raised while reading one.
    ///
    /// <see cref="PartStructureException"/> is XLibur's own statement that a package is not
    /// structured as a spreadsheet. <see cref="FileFormatException"/> comes from the packaging
    /// layer for an archive that is corrupt at the container level. The rest are the BCL's
    /// ordinary vocabulary for malformed data.
    /// </summary>
    public static bool IsRejectionDuringLoad(Exception exception)
    {
        return exception is PartStructureException
            or FileFormatException
            or InvalidDataException
            or XmlException
            or FormatException
            or OverflowException
            or ArgumentException
            or IOException;
    }

    /// <summary>
    /// Exception types tolerated while writing a workbook that has already loaded.
    ///
    /// Deliberately almost empty. <see cref="IOException"/> covers a destination stream that
    /// genuinely fails. <see cref="OutOfMemoryException"/> is tolerated because a fuzzer can
    /// describe a sheet whose declared extent is enormous — but see
    /// <see cref="ShouldReport"/>: tolerating it does not mean ignoring it.
    /// </summary>
    public static bool IsToleratedDuringSave(Exception exception)
    {
        return exception is IOException or OutOfMemoryException;
    }

    /// <summary>
    /// Nothing is tolerated when re-reading what XLibur just wrote. If XLibur produced those
    /// bytes then XLibur can read them; anything else is silent corruption in the write path,
    /// which is the failure mode no exception-based check on the first load can ever see.
    /// </summary>
    public static bool IsToleratedDuringReload(Exception exception)
    {
        _ = exception;
        return false;
    }

    /// <summary>
    /// Whether a tolerated exception is nonetheless worth putting in front of a human.
    ///
    /// An input of a few hundred kilobytes that exhausts memory inside the writer is not a crash,
    /// so it must not fail the run — but it is a resource-exhaustion result and losing it silently
    /// is how the previous harness came to look clean while finding nothing.
    /// </summary>
    public static bool ShouldReport(Exception exception)
    {
        return exception is OutOfMemoryException;
    }

    /// <summary>Write a tolerated-but-notable event where a human will find it after the run.</summary>
    public static void Report(string target, string phase, Exception exception)
    {
        // Console output is unreliable under a fuzzing driver, so this goes to a file. The
        // directory is supplied by fuzz.ps1; when it is absent the harness is being run by hand
        // and the console is fine.
        var line = $"{DateTime.UtcNow:O}\t{target}\t{phase}\t{exception.GetType().FullName}\t{exception.Message}";
        var directory = Environment.GetEnvironmentVariable("XLIBUR_FUZZ_REPORT_DIR");
        if (string.IsNullOrWhiteSpace(directory))
        {
            Console.Error.WriteLine(line);
            return;
        }

        try
        {
            Directory.CreateDirectory(directory);
            File.AppendAllText(Path.Combine(directory, "tolerated.tsv"), line + Environment.NewLine);
        }
        catch (IOException)
        {
            // Reporting must never be the reason a fuzzing run fails.
        }
    }

    /// <summary>
    /// Guard against a rule ever being written against an OpenXml type again. The library is
    /// supposed to convert those at its own boundary (see <c>XLWorkbook.OpenPackage</c>), so one
    /// arriving here means that conversion has a hole in it, and that is itself the finding.
    /// </summary>
    public static bool IsLeakedPackageReaderType(Exception exception)
    {
        return exception is OpenXmlPackageException
            || exception.GetType().Namespace?.StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal) == true;
    }
}
