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

## The password is never carried from load to save

Opening a workbook with a password does **not** cause it to be re-encrypted when you save it. To
keep a file encrypted, supply the password again:

```csharp
using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions { Password = "s3cret" });

workbook.Worksheet("Data").Cell("A1").Value = "Updated";

// Still encrypted — the password has to be given again
workbook.SaveAs("Confidential.xlsx", new SaveOptions { Password = "s3cret" });
```

This is deliberate. If the password carried over implicitly, a plain `SaveAs` would silently
produce a file the caller could not open without a password they never mentioned. The flip side
is that forgetting it decrypts the file, so make the save password explicit wherever a file is
meant to stay protected.

Removing the protection is therefore just a save without a password:

```csharp
using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions { Password = "s3cret" });
workbook.SaveAs("Public.xlsx");                    // plain, unencrypted
```

And changing it is a save with a different one:

```csharp
workbook.SaveAs("Confidential.xlsx", new SaveOptions { Password = "n3wsecret" });
```

:::warning `Save()` does not update an encrypted file in place
A workbook opened from an encrypted file is backed by the decrypted copy held in memory, because
the file on disk is a compound file rather than a package the save path can patch. `Save()` writes
to that in-memory copy and leaves the file on disk untouched — your changes go nowhere.

Always use `SaveAs` for a workbook that was loaded encrypted:

```csharp
// Wrong — the file on disk is not updated
workbook.Save();

// Right
workbook.SaveAs("Confidential.xlsx", new SaveOptions { Password = "s3cret" });
```
:::

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
  them — if that matters in your threat model, keep their lifetime short.
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
