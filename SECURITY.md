# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| latest  | :white_check_mark: |

Only the latest release receives security fixes.

## Reporting a Vulnerability

If you discover a security vulnerability in XLibur, please report it responsibly.

**Do not open a public GitHub issue for security vulnerabilities.**

Instead, please report vulnerabilities by emailing the maintainers or by using [GitHub's private vulnerability reporting](https://github.com/XLibur/XLibur/security/advisories/new).

When reporting, please include:

- A description of the vulnerability
- Steps to reproduce the issue
- The potential impact
- Any suggested fix (if available)

We will acknowledge receipt within 72 hours and aim to provide a fix or mitigation plan within 30 days, depending on severity.

## Software Bill of Materials

Every published package carries an SBOM, in two forms.

**Embedded, in the package.** Each `.nupkg` contains an SPDX 2.2 manifest generated at pack
time, so it is available to anyone who installs from NuGet without going back to this
repository:

```bash
unzip -p XLibur.0.200.0.nupkg '_manifest/spdx_2.2/manifest.spdx.json' | jq .
```

`_manifest/spdx_2.2/manifest.spdx.json.sha256` beside it is the manifest's own checksum. Note
that it attests to the manifest, not to the package — nuget.org rewrites `.nupkg` files when
they are published, so no signature produced here can cover the file you download from it.

**Attached to each release.** Every GitHub Release also carries a CycloneDX 1.7 document per
package, named to match the package it describes (`XLibur.0.200.0.cdx.json`). These resolve the
project graph rather than scanning build output, so they are the more precise record of what a
given package depends on, and the better input for automated tooling.

Both are produced by the release workflows, from the same commit the packages are built from.
See `Directory.Build.targets` and `.github/scripts/generate-sbom.sh`.

## Scope

XLibur processes `.xlsx` and `.xlsm` files, which are ZIP-based XML packages. Security concerns include but are not limited to:

- XML External Entity (XXE) attacks via crafted OpenXML content
- Zip bomb / decompression bomb attacks
- Path traversal via malicious package part names
- Denial of service via excessively large or deeply nested structures
