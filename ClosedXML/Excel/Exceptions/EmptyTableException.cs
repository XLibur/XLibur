#nullable disable

using System;

namespace ClosedXML.Excel.Exceptions;

public class EmptyTableException : ClosedXMLException
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
