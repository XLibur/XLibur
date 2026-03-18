using System;

namespace XLibur.Excel.Exceptions;

public class EmptyTableException : XLiburException
{
    public EmptyTableException()
    { }

    public EmptyTableException(string message)
        : base(message)
    { }

    public EmptyTableException(string message, Exception innerException)
        : base(message, innerException)
    { }
}
