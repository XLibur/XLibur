using XLibur.Excel;

namespace XLibur.Fuzz;

/// <summary>
/// Thrown when XLibur cannot read back a workbook it has just written.
///
/// A distinct type because this is the most serious thing the harness can find and the least
/// visible: the input was accepted correctly, the save reported success, and the damage exists
/// only in the bytes. Wrapping it also puts the phase in the message, which the stack trace
/// cannot supply — reading written bytes runs the same code as reading given bytes.
/// </summary>
internal sealed class ReloadFailedException(int writtenLength, Exception innerException)
    : Exception(
        $"XLibur could not read back the {writtenLength} bytes it wrote. " +
        "The input loaded and saved without complaint, so this is a defect in the write path.",
        innerException);

/// <summary>How far a candidate package got before the pipeline stopped with it.</summary>
internal enum WorkbookOutcome
{
    /// <summary>XLibur declined to load it. Correct behaviour for input that is not a workbook.</summary>
    Rejected,

    /// <summary>It loaded, but writing it back was stopped by a tolerated resource failure.</summary>
    SaveTolerated,

    /// <summary>It loaded, was written out, and was read back. A full round trip.</summary>
    RoundTripped,
}

/// <summary>
/// Drives one candidate package through load, save, and load-again, applying the phase-specific
/// rules in <see cref="Oracle"/> to each step. Both workbook targets share it, so the blind and
/// the structure-aware target can never drift into judging the same behaviour differently.
/// </summary>
internal static class WorkbookPipeline
{
    public static WorkbookOutcome Run(string target, byte[] candidate)
    {
        return Run(target, candidate, out _);
    }

    /// <summary>
    /// As <see cref="Run(string, byte[])"/>, but reporting <em>why</em> a candidate was rejected.
    ///
    /// A rejection is not a finding, so fuzzing has no use for the reason. Triage does: a
    /// structure-aware generator whose packages are all rejected on sight is reaching no more of
    /// the library than blind mutation, and without the reason there is no way to tell that from
    /// a generator that is working. Discarding it cost an hour the first time.
    /// </summary>
    public static WorkbookOutcome Run(string target, byte[] candidate, out string? rejection)
    {
        rejection = null;

        // The input stream must outlive the save. XLWorkbook keeps a reference to the stream it
        // loaded from and rewinds it during SaveAs, so disposing it earlier produces an
        // ObjectDisposedException out of the write path that looks exactly like a write-path
        // defect and is not one. An earlier version of this method scoped the stream to the load
        // and cost a false finding.
        using var input = new MemoryStream(candidate, writable: false);

        XLWorkbook workbook;
        try
        {
            workbook = new XLWorkbook(input);
        }
        catch (Exception exception) when (!Oracle.IsLeakedPackageReaderType(exception)
                                          && Oracle.IsRejectionDuringLoad(exception))
        {
            // XLibur declined the input. That is the whole point of a rejection: nothing to see.
            rejection = $"{exception.GetType().Name} at {StackSummary.FirstMeaningfulFrame(exception)}: {exception.Message}";
            return WorkbookOutcome.Rejected;
        }

        byte[] written;
        using (workbook)
        {
            try
            {
                using var output = new MemoryStream();
                workbook.SaveAs(output);
                written = output.ToArray();
            }
            catch (Exception exception) when (Oracle.IsToleratedDuringSave(exception))
            {
                if (Oracle.ShouldReport(exception))
                    Oracle.Report(target, "save", exception);

                return WorkbookOutcome.SaveTolerated;
            }
        }

        // XLibur wrote these bytes, so XLibur can read them. Anything thrown here is corruption
        // in the write path, and no check on the first load could have seen it.
        try
        {
            using var reread = new MemoryStream(written, writable: false);
            using var reloaded = new XLWorkbook(reread);
        }
        catch (Exception exception) when (Oracle.IsToleratedDuringReload(exception))
        {
            // Unreachable by construction; present so the phase reads the same as the others
            // and so widening the rule is a one-line change in Oracle rather than here.
        }
        catch (Exception exception)
        {
            // Say which phase this was. The stack cannot: reading the bytes XLibur just wrote runs
            // exactly the same code as reading the bytes it was given, so a reload failure and a
            // load failure produce an identical top frame while meaning opposite things — one is
            // XLibur correctly refusing someone else's bad file, the other is XLibur producing a
            // bad file of its own. Triage guessed wrong once before this was added.
            Dump(candidate, written);
            throw new ReloadFailedException(written.Length, exception);
        }

        return WorkbookOutcome.RoundTripped;
    }

    /// <summary>
    /// Write the package XLibur was given and the package it produced, side by side, when
    /// <c>XLIBUR_FUZZ_DUMP_DIR</c> asks for them.
    ///
    /// A write-path defect is a claim about two documents — what went in and what came out — and
    /// neither an exception type nor a stack frame carries either of them. Reconstructing the pair
    /// by hand from a fuzzer's input bytes means re-running the generator, so the harness hands
    /// them over instead.
    /// </summary>
    private static void Dump(byte[] given, byte[] produced)
    {
        var directory = Environment.GetEnvironmentVariable("XLIBUR_FUZZ_DUMP_DIR");
        if (string.IsNullOrWhiteSpace(directory))
            return;

        try
        {
            Directory.CreateDirectory(directory);
            var stamp = DateTime.UtcNow.ToString("HHmmss_fff", System.Globalization.CultureInfo.InvariantCulture);
            File.WriteAllBytes(Path.Combine(directory, $"{stamp}-given.xlsx"), given);
            File.WriteAllBytes(Path.Combine(directory, $"{stamp}-produced.xlsx"), produced);
        }
        catch (IOException)
        {
            // Diagnostics must never be the reason a run fails.
        }
    }
}
