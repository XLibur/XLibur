using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// A formula depends on its own value, either directly or through other formulas.
/// </summary>
/// <remarks>
/// <para>
/// Thrown when a cell that is part of a circular reference is read, or evaluated through
/// <see cref="IXLWorksheet.Evaluate(string, string)"/> or <see cref="IXLWorkbook.Evaluate(string)"/>.
/// <see cref="IXLWorkbook.RecalculateAllFormulas"/> does not throw it: it leaves the cells of a cycle
/// dirty and calculates the rest.
/// </para>
/// <para>
/// A cycle is a property of the workbook, not a failure of XLibur, so catching this type is how a
/// caller tells one from a defect. It derives from <see cref="InvalidOperationException"/>, the type
/// a cycle has always been reported as, so a <c>catch (InvalidOperationException)</c> written before
/// this type was public still catches it. It does not derive from
/// <see cref="XLibur.Excel.Exceptions.XLiburException"/>.
/// </para>
/// </remarks>
public sealed class XLCircularReferenceException : InvalidOperationException
{
    /// <summary>Creates the exception with a default message.</summary>
    public XLCircularReferenceException()
        : base("A formula depends on its own value.")
    {
    }

    /// <summary>Creates the exception with <paramref name="message"/>.</summary>
    public XLCircularReferenceException(string message)
        : base(message)
    {
    }

    /// <summary>Creates the exception with <paramref name="message"/> and <paramref name="innerException"/>.</summary>
    public XLCircularReferenceException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
