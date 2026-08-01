using System;
using XLibur.Excel.Exceptions;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// A function was called without a worksheet to be relative to, and needs one.
/// </summary>
/// <remarks>
/// <para>
/// Thrown by <see cref="XLFunctionLibrary.TryInvoke"/> for the functions whose result depends on
/// where they are — <c>ROW</c>, <c>COLUMN</c>, <c>OFFSET</c>, <c>INDIRECT</c> and the like. There
/// is no answer to give outside a grid, and returning one would be misleading, so these belong in
/// a real cell formula.
/// </para>
/// <para>
/// This is the public face of the calc engine's internal missing-context signal. Catching it is
/// how a caller tells "this function cannot work here" from a function that ran and failed, which
/// reports itself as an <see cref="XLError"/> result instead.
/// </para>
/// </remarks>
public sealed class XLNoWorksheetContextException : XLiburException
{
    /// <summary>Creates the exception with a default message.</summary>
    public XLNoWorksheetContextException()
        : base("The function needs a worksheet to be relative to, and was called without one.")
    {
    }

    /// <summary>Creates the exception with <paramref name="message"/>.</summary>
    public XLNoWorksheetContextException(string message)
        : base(message)
    {
    }

    /// <summary>Creates the exception with <paramref name="message"/> and <paramref name="innerException"/>.</summary>
    public XLNoWorksheetContextException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
