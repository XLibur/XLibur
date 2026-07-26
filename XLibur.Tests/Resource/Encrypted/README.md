# Encrypted workbook test corpus

Files here are password-protected workbooks produced by *other* software. They exist to prove
XLibur reads what real applications write, which a round trip through XLibur's own encrypt and
decrypt cannot show: a shared bug in both directions would cancel itself out and the test would
still pass.

`XLibur.Tests/Excel/Encryption/WorkbookEncryptionTests.cs` covers the round trip and the error
paths. Those tests need nothing from this folder.

## Status

**The corpus is not populated yet.** Acceptance criteria 1 and 2 of
`docs/specs/06-workbook-encryption.md` stay open until it is, and the file-backed tests below stay
unwritten. This is recorded rather than quietly skipped, because "the tests pass" currently means
"XLibur agrees with itself".

## What is needed

Every file uses the password `xlibur-test` unless its row says otherwise. Add a row per file.

| File | Produced by | Encryption | Password | Notes |
|---|---|---|---|---|
| _(none yet)_ | | | | |

Wanted, in rough order of value:

| Producer | Why it matters |
|---|---|
| Excel 2016 or later, default settings | Agile AES-256 + SHA-512. The overwhelmingly common case. |
| Excel 2007 | Standard encryption. The only way to exercise that path against a real file. |
| LibreOffice Calc | Writes agile encryption with its own parameter choices, so it catches assumptions baked in from Excel's defaults. |
| Excel, non-ASCII password | Confirms the UTF-16LE password encoding on a real file rather than only in a round trip. |
| Excel, empty-ish or very long password | Boundary cases in key derivation. |

Also worth capturing, since the error paths are as much a part of the contract as the happy one:

- a file truncated mid-`EncryptedPackage`
- a file whose `EncryptedPackage` has had a byte flipped, to fail the HMAC

## Producing a file

In Excel: **File → Info → Protect Workbook → Encrypt with Password**, then save as `.xlsx`.

In LibreOffice Calc: **File → Save As**, tick **Save with password**.

Record the exact application version in the table above. When a file later fails to load, the
version is what makes the difference between "our reader is wrong" and "that release wrote
something unusual" — and it cannot be recovered from the file afterwards.

## Wiring a file into the tests

Add a case to `WorkbookEncryptionTests` that opens the resource with its password and asserts on a
known cell. Resources are embedded, so no test project change is needed beyond the file itself —
see `TestHelper.GetStreamFromResource`.
