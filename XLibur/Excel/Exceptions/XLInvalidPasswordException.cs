using System;

namespace XLibur.Excel.Exceptions;

/// <summary>
/// A workbook is encrypted and the password needed to decrypt it was missing or wrong.
/// </summary>
/// <remarks>
/// Deliberately distinct from the exceptions a malformed file produces: a caller prompting for a
/// password needs to tell "you typed the wrong password" apart from "this file is broken".
/// </remarks>
public class XLInvalidPasswordException : XLiburException
{
    public XLInvalidPasswordException()
        : base("The workbook is encrypted and the password is missing or incorrect.")
    {
    }

    public XLInvalidPasswordException(string message)
        : base(message)
    {
    }

    public XLInvalidPasswordException(string message, Exception inner)
        : base(message, inner)
    {
    }
}
