#!/usr/bin/env bash
#
# Generates a CycloneDX SBOM for each shipped package, one file per package, named to
# match the .nupkg it describes:
#
#   XLibur.0.200.0.nupkg  ->  XLibur.0.200.0.cdx.json
#
# The release workflows attach these to the GitHub Release. They complement the SPDX
# manifest that Microsoft.Sbom.Targets embeds inside each .nupkg (see
# Directory.Build.targets): the embedded one travels to anyone installing from
# nuget.org, this one is the higher-fidelity document for tooling, because CycloneDX
# resolves the ProjectReference graph rather than reading a flat component scan.
#
# Usage: generate-sbom.sh <version> <output-dir> <project>...
#
#   version     Version being released, with or without a leading "v". Used for the
#               filenames and stamped into each document.
#   output-dir  Created if absent.
#   project     Project directory name, e.g. XLibur.Bundle. Expected to contain
#               <project>/<project>.csproj.
#
# Requires the CycloneDX tool from .config/dotnet-tools.json — run `dotnet tool restore`
# first.
#
# Exits non-zero if any expected document is missing or empty, so a silently failing
# tool cannot ship a release whose SBOMs are absent.

set -euo pipefail

version=${1:?usage: generate-sbom.sh <version> <output-dir> <project>...}
output=${2:?usage: generate-sbom.sh <version> <output-dir> <project>...}
shift 2

if [ "$#" -eq 0 ]; then
  echo "::error::No projects given — nothing to generate" >&2
  exit 1
fi

# The tag prefix is not part of a package version.
version=${version#v}

repo_root="$(cd "$(dirname "$0")/../.." && pwd)"
mkdir -p "$output"

status=0

for project in "$@"; do
  csproj="$repo_root/$project/$project.csproj"
  filename="$project.$version.cdx.json"

  if [ ! -f "$csproj" ]; then
    echo "::error::$csproj does not exist"
    status=1
    continue
  fi

  echo "::group::SBOM for $project"

  # --recursive matters: XLibur.Bundle and XLibur.Report reach most of their closure
  # through ProjectReference, and without it their documents list almost nothing.
  # --set-name/--set-version pin the document to the package identity rather than
  # letting the tool infer it from the project, which at pack time is a MinVer
  # pre-release on any commit that is not exactly a release tag.
  if ! dotnet tool run dotnet-CycloneDX "$csproj" \
      --output "$output" \
      --filename "$filename" \
      --output-format Json \
      --recursive \
      --exclude-dev \
      --exclude-test-projects \
      --set-name "$project" \
      --set-version "$version"; then
    echo "::error::CycloneDX failed for $project"
    status=1
  fi

  echo "::endgroup::"

  if [ ! -s "$output/$filename" ]; then
    echo "::error::$filename was not produced"
    status=1
  fi
done

if [ "$status" -ne 0 ]; then
  echo "::error::SBOM generation failed — refusing to continue"
  exit 1
fi

echo "Generated SBOMs in $output:"
ls -la "$output"
