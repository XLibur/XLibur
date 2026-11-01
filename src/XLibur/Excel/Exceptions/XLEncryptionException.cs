using System;

namespace XLibur.Excel.Exceptions;

/// <summary>
/// An encrypted workbook could not be read or written: the container is malformed, the encryption
/// descriptor uses a cipher or hash combination outside the profile Excel emits, or the integrity
/// check over the encrypted package failed.
/// </summary>
/// <remarks>
/// A failed integrity check reports through this type rather than
/// <see cref="XLInvalidPasswordException"/>: the password was right, the bytes were altered after
/// they were written. Reporting it as a bad password would send the caller off to re-prompt for a
/// password that was never the problem.
/// </remarks>
public class XLEncryptionException : XLiburException
{
    public XLEncryptionException()
        : base("The encrypted workbook could not be processed.")
    {
    }

    public XLEncryptionException(string message)
        : base(message)
    {
    }

    public XLEncryptionException(string message, Exception inner)
        : base(message, inner)
    {
    }
}
