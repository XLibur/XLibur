using System.IO;

namespace XLibur.Excel.IO.Encryption;

/// <summary>
/// Turns the .xlsx bytes of a built package into the encrypted compound file that carries them.
/// </summary>
/// <remarks>
/// A named boundary around the one step of a save that can fail after the package exists but before
/// the destination has been touched. <see cref="XLWorkbook"/> holds one of these rather than calling
/// <see cref="WorkbookEncryption"/> directly, so that a test can substitute an encryption that fails
/// and observe what the destination looks like afterwards.
/// </remarks>
internal interface IWorkbookEncryptor
{
    /// <summary>
    /// Encrypts <paramref name="package"/> into a compound file written to
    /// <paramref name="destination"/>.
    /// </summary>
    void Encrypt(Stream destination, byte[] package, string password);
}

/// <summary>
/// The encryption every workbook uses unless something has substituted another: agile encryption
/// with the parameters Excel writes.
/// </summary>
internal sealed class WorkbookEncryptor : IWorkbookEncryptor
{
    /// <summary>
    /// The instance <see cref="XLWorkbook"/> starts with. Stateless, so one is enough.
    /// </summary>
    public static readonly WorkbookEncryptor Default = new();

    private WorkbookEncryptor()
    {
    }

    public void Encrypt(Stream destination, byte[] package, string password) =>
        WorkbookEncryption.Encrypt(destination, package, password);
}
