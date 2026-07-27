# Spec 06 — Password-Protected Workbook Support (ECMA-376 Encryption)

**Area:** Feature + Compatibility
**Effort:** L (3–4 weeks including interop testing)
**Dependencies:** None.
**Status:** ✅ Implemented in PR #245 (tests #246, guide #249) — see [Implementation notes](#implementation-notes)

## Summary

XLibur cannot open or save password-encrypted Excel files — a hard blocker for users handling files from regulated environments. Encrypted .xlsx files are OLE Compound File Binary (CFB) containers holding an `EncryptionInfo` stream and an `EncryptedPackage` stream (the real .xlsx, AES-encrypted). Zero encryption code exists in the repo today (repo-wide grep for encrypt/agile/CryptoAPI: no hits). Note: this is distinct from sheet/workbook *protection* (password hashes controlling edit permissions), which already exists in `XLibur/Excel/Protection/`.

**Authorization context:** this implements the documented ECMA-376 / MS-OFFCRYPTO standard so users can open/save their own password-protected spreadsheets — the same capability EPPlus, NPOI, and Aspose ship.

## Scope

| Capability | In scope |
|---|---|
| Read: **Agile Encryption** (Office 2010+, AES-128/256-CBC, SHA-512 key derivation) | ✅ required |
| Read: **Standard Encryption** (Office 2007, AES-128-ECB derivation) | ✅ required (still common in the wild) |
| Write: Agile Encryption (AES-256, SHA-512, Excel-default parameters) | ✅ required |
| Write: Standard Encryption | ❌ no (obsolete; Excel hasn't written it since 2007) |
| XOR-obfuscation / RC4 (legacy .xls-era) | ❌ no |
| VBA project password, sheet protection | ❌ out of scope (different mechanisms) |

## Design

### Container layer: CFB (MS-CFB)

Decision required up front (task 1): **evaluate OpenMcdf** (mature .NET CFB library, MPL-2.0) vs writing a minimal in-house CFB reader/writer.
- OpenMcdf pro: battle-tested, fast. Con: MPL-2.0 dependency in an MIT project (MPL is file-level copyleft — legally compatible as a dependency, but the project has been conservative about licenses; flag to maintainer for sign-off before proceeding).
- In-house pro: no dependency; encrypted-workbook CFB usage is a narrow subset (2 streams, small directory). MS-CFB is well documented; a read+write implementation limited to what Office emits is ~600–900 lines. Reference implementations to study (clean-room, don't copy GPL code): EPPlus (PolyForm), NPOI (Apache-2.0 — safe to study), msoffcrypto-tool (MIT).

Recommendation: in-house minimal CFB (`XLibur/Excel/IO/Encryption/CompoundFile/`), unless maintainer approves the dependency.

### Crypto layer (MS-OFFCRYPTO)

All primitives exist in `System.Security.Cryptography` (Aes, SHA1/SHA512, HMACSHA*, RandomNumberGenerator) — no new packages.

- **Agile read/write** (`EncryptionInfo` version 4.4, XML descriptor):
  - Parse/emit the `<encryption>` XML (keyData, dataIntegrity, keyEncryptors/password).
  - Key derivation: iterated SHA-512 spin (spinCount, default 100,000) over password (UTF-16LE) + salt; block-key variants for verifier hash input/value, key value, HMAC key/value.
  - Package encryption: AES-CBC in 4096-byte segments, IV = H(keyData.salt ‖ segmentIndex).
  - Integrity: HMAC over the encrypted package, verified on read, emitted on write.
- **Standard read** (version 3.2/4.2 binary descriptor): AES key from SHA-1 iterated 50,000× + block index; verify via encryptedVerifier/Hash; package decrypted ECB (per spec, with the stream-size prefix).
- Wrong password → throw a dedicated `XLInvalidPasswordException` (distinguish from corrupt-file errors).

### API surface

```csharp
// Load
new XLWorkbook(path, new LoadOptions { Password = "secret" });
new XLWorkbook(stream, new LoadOptions { Password = "secret" });
// Save
workbook.SaveAs(path, new SaveOptions { Password = "secret" });  // Agile AES-256
```

- `LoadOptions.Password` (`string?`); `SaveOptions.Password` (`string?`).
- Load flow: sniff CFB signature (`D0 CF 11 E0 ...`) → if encrypted and no password, throw `XLInvalidPasswordException` with a clear message → decrypt `EncryptedPackage` into a `MemoryStream` → existing load path unchanged.
- Save flow: existing save into `MemoryStream` → encrypt → write CFB container. (Streaming encryption is possible later; MemoryStream first — files that fit in memory are the 99% case and the whole workbook is in memory anyway.)
- Re-saving a workbook loaded with a password does **not** silently re-encrypt — password must be given explicitly on save (matches EPPlus behavior; least surprise, documented).

## Work plan

| # | Task | Size |
|---|------|------|
| 1 | CFB decision spike (OpenMcdf license sign-off vs in-house); if in-house: CFB reader | M |
| 2 | CFB writer (only what Office needs: header, FAT, directory, 2 streams, mini-FAT) | M |
| 3 | Agile `EncryptionInfo` XML parse/emit + key derivation + verifier | M |
| 4 | Agile package decrypt/encrypt (segmented AES-CBC) + HMAC integrity | M |
| 5 | Standard-encryption read path | S |
| 6 | `LoadOptions.Password` / `SaveOptions.Password` wiring + `XLInvalidPasswordException` | S |
| 7 | Test corpus: files encrypted by real Excel (agile AES-256 default, long/unicode/empty-edge passwords), by LibreOffice, by EPPlus; wrong-password, truncated-container, tampered-HMAC cases | M |
| 8 | Round-trip: save-encrypted → open in Excel manually once + automated reopen via XLibur; document in PR | S |

Tasks 1–2 and 3–4 can run in parallel (container vs crypto). Test files go in `XLibur.Tests/Resource/Encrypted/` with a generation README (which Excel version produced each).

## Acceptance criteria

1. Opens agile- and standard-encrypted files produced by Excel 2007–2024 and LibreOffice; wrong password throws `XLInvalidPasswordException`.
2. Files saved with a password open in Excel (manual verification, recorded in PR) and in EPPlus/openpyxl (automated where licensing permits — openpyxl via a CI-optional script is fine to skip; EPPlus test project already references EPPlus for benchmarks, check license mode).
3. Tampered ciphertext (flipped byte) fails the HMAC check with a clear error, not garbage data.
4. No new NuGet dependency without maintainer license sign-off.
5. Passwords handled as `string` per API convention but never logged; derived keys zeroed where practical (`CryptographicOperations.ZeroMemory`).

## Risks

- MS-OFFCRYPTO has many optional knobs; Excel only emits a narrow profile — implement to the profile, validate against the corpus, and reject unknown cipher/hash combinations explicitly rather than guessing.
- CFB writer subtleties (mini-stream cutoff at 4096 bytes, sector chains) — the reopen-in-Excel test is the gate.

## References

- MS-OFFCRYPTO, MS-CFB (Microsoft open specifications).
- NPOI `POIFS`/`CryptoAPI` (Apache-2.0) as a study reference.
- Existing (different) feature: `XLibur/Excel/Protection/` — password *hashes* for edit protection, untouched by this spec.

## Implementation notes

> The **Summary** and **Design** sections above describe the codebase *before* this spec was
> implemented. In particular, "zero encryption code exists in the repo today" is no longer true —
> read them as the historical problem statement, not as current state.

Landed in PR #245, with test coverage in #246 and a user guide in #249.

What shipped, against the scope table:

| Capability | Shipped |
|---|---|
| Read Agile (AES-CBC, SHA-512) | ✅ `IO/Encryption/Agile/` |
| Read Standard (Office 2007) | ✅ `IO/Encryption/Standard/` |
| Write Agile | ✅ `Agile/AgileCrypto.cs`, `AgileEncryptionDescriptor.cs` |
| Write Standard | ❌ excluded by design |
| XOR / RC4 legacy | ❌ excluded by design |

Public surface is `LoadOptions.Password` and `SaveOptions.Password`; a workbook opened with a
password is deliberately *not* re-encrypted on save unless one is supplied.

### CFB layer: OpenMcdf, approved

The design recommended writing a minimal in-house CFB reader/writer, and made acceptance
criterion 4 ("no new NuGet dependency without maintainer license sign-off") the gate on the
alternative. The alternative was taken: `OpenMcdf 3.1.4`, referenced from `XLibur/XLibur.csproj`.

**The maintainer sign-off was given** — MPL-2.0 is file-level copyleft and imposes no obligation on
the MIT-licensed code that consumes it as an unmodified package dependency. Criterion 4 is
therefore met rather than deviated from, and the in-house CFB recommendation is closed as
not-taken rather than outstanding.

The practical consequence for anyone reading this later: `XLibur/Excel/IO/Encryption/` contains no
compound-file implementation, because CFB container handling lives in OpenMcdf. Only the crypto
layer is in-house.
