namespace XLibur.Fuzz;

/// <summary>
/// Reduces a stack trace to the one frame worth reading.
///
/// Shared by the pipeline (which describes rejections) and by replay (which describes findings),
/// because the two must group inputs the same way. When each had its own copy, one preferred the
/// outermost frame and reported <c>System.Linq.ThrowHelper</c> for a defect that was actually in
/// the style decoder.
/// </summary>
internal static class StackSummary
{
    /// <summary>
    /// The outermost XLibur frame, falling back to the outermost frame of any kind.
    ///
    /// A precondition failure raised through a BCL helper reports that helper at the top of the
    /// stack, which identifies nothing: the question worth answering is which part of XLibur
    /// asked for something impossible, not which utility class noticed.
    /// </summary>
    public static string FirstMeaningfulFrame(Exception exception)
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

    /// <summary>
    /// A grouping key for one failure: the exception type and where it came from.
    ///
    /// Deliberately excludes the message, which routinely carries input-derived text and would
    /// split one defect into as many groups as there are inputs that reach it.
    /// </summary>
    public static string Signature(Exception exception)
    {
        return $"{exception.GetType().FullName} at {FirstMeaningfulFrame(exception)}";
    }
}
