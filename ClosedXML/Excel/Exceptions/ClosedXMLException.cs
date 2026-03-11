#nullable disable

using System;

namespace ClosedXML.Excel.Exceptions;

public abstract class ClosedXMLException : Exception
{
    protected ClosedXMLException()
    { }

    protected ClosedXMLException(string message)
        : base(message)
    { }

    protected ClosedXMLException(string message, Exception innerException)
        : base(message, innerException)
    { }
}
