using System;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The exception that is thrown when the strings to be parsed to an expression is invalid.
/// </summary>
public class ExpressionParseException : Exception
{
    /// <summary>
    /// Initializes a new instance of the ExpressionParseException class with a
    /// specified error message.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public ExpressionParseException(string message)
        : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the ExpressionParseException class with a
    /// specified error message and the exception that caused it.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="innerException">
    /// The parser's own exception, kept so that the detail it carries — the position in the formula,
    /// the token it did not expect — is still reachable once the message has been read.
    /// </param>
    public ExpressionParseException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
