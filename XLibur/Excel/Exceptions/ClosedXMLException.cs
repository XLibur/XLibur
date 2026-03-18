using System;

namespace XLibur.Excel.Exceptions;

public abstract class XLiburException : Exception
{
    protected XLiburException()
    { }

    protected XLiburException(string message)
        : base(message)
    { }

    protected XLiburException(string message, Exception innerException)
        : base(message, innerException)
    { }
}
