using XLibur.Excel;

namespace XLibur.Fuzz;

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

        XLWorkbook workbook;
        try
        {
            using var input = new MemoryStream(candidate, writable: false);
            workbook = new XLWorkbook(input);
        }
        catch (Exception exception) when (!Oracle.IsLeakedPackageReaderType(exception)
                                          && Oracle.IsRejectionDuringLoad(exception))
        {
            // XLibur declined the input. That is the whole point of a rejection: nothing to see.
            rejection = $"{exception.GetType().Name} at {FirstFrame(exception)}: {exception.Message}";
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

        return WorkbookOutcome.RoundTripped;
    }

    /// <summary>
    /// The frame a rejection came from. The type and message alone cannot distinguish a
    /// deliberate refusal from an internal bounds check that happens to share the type.
    ///
    /// Prefers the outermost XLibur frame over the true outermost one. A bounds check reached
    /// through LINQ reports <c>System.Linq.ThrowHelper</c> at the top, which says nothing about
    /// which part of XLibur asked the question — and that is the only thing worth knowing.
    /// </summary>
    private static string FirstFrame(Exception exception)
    {
        var stack = exception.StackTrace;
        if (string.IsNullOrEmpty(stack))
            return "(no stack)";

        var frames = stack.Split('\n')
            .Select(line => line.Trim())
            .Where(line => line.Length > 0)
            .ToArray();

        if (frames.Length == 0)
            return "(no stack)";

        var xlibur = Array.Find(frames, f => f.Contains("XLibur.", StringComparison.Ordinal));
        return xlibur ?? frames[0];
    }
}
