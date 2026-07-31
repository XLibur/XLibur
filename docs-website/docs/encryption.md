---
id: encryption
title: Encryption and Passwords
sidebar_label: Encryption
description: Open and save password-protected workbooks with ECMA-376 encryption — load and save passwords, the exceptions to handle, and which encryption schemes are supported.
---

# Encryption and Passwords

A password-protected workbook is encrypted on disk. Excel asks for the password before it will
show you anything, and without it the contents cannot be recovered by any tool — the file is a
container of ciphertext rather than a spreadsheet.

XLibur reads these files and writes them, through `LoadOptions.Password` and
`SaveOptions.Password`.

:::warning Encryption is not the same as protection
This page is about **encrypting the file**. That is different from
[workbook and sheet protection](./workbook-settings.md#protection), which controls what a user is
allowed to *edit* in a file that anyone can open.

Protection is stored in plain text inside the `.xlsx` and any tool — including XLibur — can read
straight past it. Encryption is the only one of the two that actually keeps a file's contents
from being read.
:::

## Opening an encrypted workbook

Set `Password` on `LoadOptions`:

```csharp
using XLibur.Excel;

using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions
{
    Password = "s3cret",
});

var value = workbook.Worksheet("Data").Cell("A1").GetString();
```

The same works for a stream:

```csharp
using var stream = File.OpenRead("Confidential.xlsx");
using var workbook = new XLWorkbook(stream, new LoadOptions { Password = "s3cret" });
```

`Password` is only consulted when the file turns out to be encrypted, so setting it against a file
that isn't is harmless. That makes it safe to pass a password for input of mixed provenance
without sniffing each file first:

```csharp
// Fine whether or not Report.xlsx is encrypted
using var workbook = new XLWorkbook(path, new LoadOptions { Password = suppliedPassword });
```

## Saving an encrypted workbook

Set `Password` on `SaveOptions`:

```csharp
using var workbook = new XLWorkbook();
var ws = workbook.AddWorksheet("Data");
ws.Cell("A1").Value = "Commercially sensitive";

workbook.SaveAs("Confidential.xlsx", new SaveOptions
{
    Password = "s3cret",
});
```

Files are written with **agile encryption** using the parameters Excel itself writes —
AES-256-CBC, SHA-512, a spin count of 100,000, and a freshly generated random key per save — so
the result is indistinguishable from a file Excel produced, and opens in Excel with no complaint.

Saving to a stream works the same way:

```csharp
using var ms = new MemoryStream();
workbook.SaveAs(ms, new SaveOptions { Password = "s3cret" });
```

## `Save` keeps the encryption, `SaveAs` states it

The two save methods treat a missing password differently, because they are asking for different
things.

**`Save()` puts a workbook back where it came from, as it came.** A file opened with a password is
written back to that file encrypted, under the same password:

```csharp
using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions { Password = "s3cret" });

workbook.Worksheet("Data").Cell("A1").Value = "Updated";

workbook.Save();          // Confidential.xlsx — still encrypted, still "s3cret"
```

**`SaveAs` describes the file it is about to write.** A password means encrypt with it, no password
means plain, whatever the workbook was loaded as:

```csharp
workbook.SaveAs("Public.xlsx");                                              // plain
workbook.SaveAs("Copy.xlsx", new SaveOptions { Password = "s3cret" });       // encrypted
```

So a null `Password` means *unchanged* to `Save` and *none* to `SaveAs`. The asymmetry is the
point: a plain `SaveAs` can never silently produce a file you cannot open, and a plain `Save` can
never silently drop the protection off a file that had it.

### Changing the password

Give `Save` a password and it replaces the one the file already had, in place:

```csharp
using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions { Password = "s3cret" });
workbook.Save(new SaveOptions { Password = "n3wsecret" });
```

The same call on a workbook that was *not* loaded encrypted encrypts it for the first time:

```csharp
using var workbook = new XLWorkbook("Public.xlsx");
workbook.Save(new SaveOptions { Password = "s3cret" });   // Public.xlsx is now encrypted
```

### Removing the password

`Save` cannot remove encryption — there is no way to express it, which is what stops it happening
by accident. Decrypting is a `SaveAs` without a password, and the path can be the one you are
already at:

```csharp
using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions { Password = "s3cret" });

workbook.SaveAs("Confidential.xlsx");   // decrypted in place
workbook.SaveAs("Public.xlsx");         // or written out plain somewhere else
```

After that the workbook is no longer an encrypted one, so a later `Save()` keeps writing plain
files. `SaveAs` sets the encryption state as well as using it.

### Streams behave the same way

`Save()` on a workbook loaded from a stream writes the re-encrypted container back to that same
stream, so the stream has to be writable and seekable — `Save` says so with an
`InvalidOperationException` if it is not, rather than dropping the changes:

```csharp
using var file = File.Open("Confidential.xlsx", FileMode.Open, FileAccess.ReadWrite);
using var workbook = new XLWorkbook(file, new LoadOptions { Password = "s3cret" });

workbook.Worksheet("Data").Cell("A1").Value = "Updated";
workbook.Save();          // the stream now holds the re-encrypted workbook
```

## Handling a wrong password

Two exceptions, distinguished by whether the password is the problem:

| Exception | Means | What to do |
|---|---|---|
| `XLInvalidPasswordException` | The password is wrong, or none was supplied for an encrypted file | Re-prompt the user |
| `XLEncryptionException` | The password was right but the file is malformed, tampered with, or uses an unsupported scheme | Report the file as unreadable — do not re-prompt |

Both derive from `XLiburException`.

The distinction matters at the point where you would ask the user to type the password again. A
failed integrity check means the bytes were altered after they were written; sending the user off
to re-enter a password that was never the problem wastes their time and hides the real fault.

```csharp
using XLibur.Excel;
using XLibur.Excel.Exceptions;

static XLWorkbook OpenWithPrompt(string path, Func<string> promptForPassword)
{
    for (var attempt = 0; attempt < 3; attempt++)
    {
        try
        {
            return new XLWorkbook(path, new LoadOptions { Password = promptForPassword() });
        }
        catch (XLInvalidPasswordException)
        {
            // Wrong password — worth another try
        }
    }

    throw new InvalidOperationException($"Could not open {path} after three attempts.");
}
```

`XLEncryptionException` is deliberately not caught there: it will propagate, which is what you
want for a corrupt file.

Each exception carries a message naming the specific problem — an unexpected cipher, chaining
mode, hash or key length is rejected by name rather than approximated into something that
half-works.

## What is supported

| Scheme | Read | Write |
|---|---|---|
| Agile encryption (Office 2010 and later) | Yes | Yes |
| Standard encryption (Office 2007) | Yes | No — re-saved as agile |
| RC4 encryption (Office 97–2003) | No | No |
| Legacy `.xls` workbooks | No | No |

Reading covers both schemes in current use. Writing is agile only, which is what Excel has
produced for well over a decade — so a standard-encrypted file you open and save comes back out
as agile, which Excel opens with the same password.

An RC4-encrypted file reports that re-saving from Excel will upgrade it. A legacy `.xls` is a
compound file too, so it would otherwise look like an encrypted workbook; it is detected and named
rather than reported as a mysterious decryption failure.

The specification permits far more cipher and hash combinations than Excel emits. Anything outside
that profile is rejected with a message saying what the file asked for, rather than being decrypted
into plausible-looking garbage.

## Streams must be seekable

Encryption is detected by sniffing the compound-file signature at the start of the file, which
requires seeking. A non-seekable stream is passed straight through as an ordinary package:

```csharp
// Works — MemoryStream and FileStream are seekable
using var workbook = new XLWorkbook(seekableStream, new LoadOptions { Password = "s3cret" });
```

If you are reading from something non-seekable — a network or compressed stream — copy it into a
`MemoryStream` first, or the encrypted file will fail as a corrupt archive rather than asking for
the password:

```csharp
using var buffer = new MemoryStream();
await responseStream.CopyToAsync(buffer);
buffer.Position = 0;

using var workbook = new XLWorkbook(buffer, new LoadOptions { Password = "s3cret" });
```

## Practical notes

- **There is no recovery path.** A lost password means lost data; XLibur has no backdoor and
  neither does Excel. Whatever stores your passwords needs to be as durable as the files.
- **Encryption is deliberately slow.** The spin count of 100,000 iterations exists to make
  brute-forcing expensive, and it applies on every open and every save. Budget for it when
  encrypting files in bulk, and do not put an encrypted save inside a tight loop.
- **Derived keys are zeroed** after use with `CryptographicOperations.ZeroMemory`. Passwords are
  ordinary strings, matching the rest of the API, so they live in memory until the GC reclaims
  them. A workbook opened from an encrypted file holds its load password for as long as the
  workbook lives, because that is what lets `Save()` put the file back encrypted — if that matters
  in your threat model, keep the workbook short-lived.
- **The whole file is encrypted**, not individual sheets. Excel offers no way to encrypt part of a
  workbook, so neither does XLibur. Split the data across files if different audiences need
  different access.

## A worked example

Generating a confidential report, then reading it back:

```csharp
using XLibur.Excel;
using XLibur.Excel.Exceptions;

const string Password = "correct horse battery staple";

// Write
using (var workbook = new XLWorkbook())
{
    var ws = workbook.AddWorksheet("Salaries");
    ws.Cell("A1").Value = "Employee";
    ws.Cell("B1").Value = "Salary";
    ws.Range("A1:B1").Style.Font.Bold = true;

    ws.Cell("A2").Value = "A. Example";
    ws.Cell("B2").Value = 62_000;
    ws.Cell("B2").Style.NumberFormat.Format = "#,##0";

    // Protection stops casual edits; encryption stops the file being read at all.
    workbook.Protect(Password, XLProtectionAlgorithm.Algorithm.SHA512);

    workbook.SaveAs("Salaries.xlsx", new SaveOptions
    {
        Password = Password,
        EvaluateFormulasBeforeSaving = true,
    });
}

// Read
try
{
    using var workbook = new XLWorkbook("Salaries.xlsx", new LoadOptions { Password = Password });
    var salary = workbook.Worksheet("Salaries").Cell("B2").GetDouble();
    Console.WriteLine($"Salary: {salary:N0}");
}
catch (XLInvalidPasswordException)
{
    Console.Error.WriteLine("Wrong password.");
}
catch (XLEncryptionException ex)
{
    Console.Error.WriteLine($"The file could not be decrypted: {ex.Message}");
}
```

## Where to next

- [Workbook Settings](./workbook-settings.md#protection) — workbook and sheet protection, and the
  load and save options encryption sits alongside
- [Worksheets](./worksheets.md#protecting-a-sheet) — protecting the contents of a single sheet
- [Importing and Exporting](./importing-exporting.md) — the other ways data gets in and out
