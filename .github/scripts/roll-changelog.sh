#!/usr/bin/env bash
#
# Rolls the "Unreleased" section of CHANGELOG.md into a dated version heading and
# leaves a fresh, empty "Unreleased" section at the top.
#
#   ## Unreleased          ->   ## Unreleased
#
#   ### Fixed                    ## v0.106.0 - 2026-07-25
#   - ...
#                                ### Fixed
#                                - ...
#
# With --report it rolls the "## XLibur.Report — Unreleased" section instead.
# XLibur.Report releases on its own report-v* tag stream, so its section keeps its
# place in the file. That section opens with a paragraph explaining the separate
# stream, which stays under the Unreleased heading; the new version heading goes in
# just before the section's first "### " heading.
#
#   ## XLibur.Report — Unreleased     ->   ## XLibur.Report — Unreleased
#
#   (intro paragraph)                        (intro paragraph)
#
#   ### Fixed                                ## XLibur.Report v0.400.0 - 2026-09-26
#   - ...
#                                            ### Fixed
#                                            - ...
#
# Usage: roll-changelog.sh [--report] <version> [date] [changelog-path]
#
#   --report        Roll the XLibur.Report section rather than core's.
#   version         Version being released, with or without a leading "v".
#   date            Release date as YYYY-MM-DD. Defaults to today (UTC).
#   changelog-path  Defaults to CHANGELOG.md in the repository root.
#
# Exits non-zero if the changelog has no matching "Unreleased" heading, or if that
# section is empty — releasing with nothing written down is almost always a mistake.

set -euo pipefail

report=false
if [[ ${1:-} == --report ]]; then
  report=true
  shift
fi

version=${1:?usage: roll-changelog.sh [--report] <version> [date] [changelog-path]}
date=${2:-$(date -u +%Y-%m-%d)}
changelog=${3:-"$(dirname "$0")/../../CHANGELOG.md"}

# Normalise to a leading "v" so headings read "## v0.106.0".
version=${version#report-}
version=${version#v}

if $report; then
  title='## XLibur.Report — Unreleased'
  prefix='## XLibur.Report '
else
  title='## Unreleased'
  prefix='## '
fi
heading="${prefix}v${version} - ${date}"

if [[ ! -f $changelog ]]; then
  echo "::error::Changelog not found: $changelog" >&2
  exit 2
fi

# Headings are compared as fixed strings, ignoring trailing whitespace — the file
# is checked out with CRLF endings, so every line carries a trailing \r.
has_line() {
  awk -v want="$1" '
    { line = $0; sub(/[[:space:]]+$/, "", line) }
    line == want { found = 1; exit }
    END { exit !found }
  ' "$changelog"
}

has_prefix() {
  awk -v want="$1" '
    index($0, want) == 1 { found = 1; exit }
    END { exit !found }
  ' "$changelog"
}

if ! has_line "$title"; then
  echo "::error::No '$title' heading found in $changelog" >&2
  exit 3
fi

if has_prefix "${prefix}v${version} "; then
  echo "::error::$changelog already contains a heading for ${prefix#'## '}v${version}" >&2
  exit 4
fi

# Everything between the Unreleased heading and the next "## " heading (or EOF).
unreleased_body=$(awk -v title="$title" '
  { line = $0; sub(/[[:space:]]+$/, "", line) }
  line == title { inside = 1; next }
  inside && /^## / { exit }
  inside { print }
' "$changelog")

# The Report section always has its intro paragraph, so it counts as empty until
# it has a "### " heading of changes.
if $report; then
  if ! grep -qE '^### ' <<<"$unreleased_body"; then
    echo "::error::The '$title' section of $changelog has no '### ' heading — nothing to release" >&2
    exit 5
  fi
elif [[ -z ${unreleased_body//[[:space:]]/} ]]; then
  echo "::error::The Unreleased section of $changelog is empty — nothing to release" >&2
  exit 5
fi

tmp=$(mktemp)
trap 'rm -f "$tmp"' EXIT

if $report; then
  awk -v title="$title" -v heading="$heading" '
    { line = $0; sub(/[[:space:]]+$/, "", line) }
    line == title { inside = 1; print; next }
    inside && /^## / { inside = 0 }
    inside && !done && /^### / {
      print heading
      print ""
      done = 1
    }
    { print }
  ' "$changelog" >"$tmp"
else
  awk -v title="$title" -v heading="$heading" '
    { print }
    { line = $0; sub(/[[:space:]]+$/, "", line) }
    !done && line == title {
      print ""
      print heading
      done = 1
    }
  ' "$changelog" >"$tmp"
fi

mv "$tmp" "$changelog"
trap - EXIT

# Rebuild the "Contents" list from the "## " headings, so it cannot drift out of
# date as releases are rolled. Anchors follow GitHub's rule: lower-case, drop every
# character that is not a letter, digit, space or hyphen, then spaces to hyphens.
contents=$(
  grep -E '^## ' "$changelog" | while IFS= read -r line; do
    text=${line#'## '}
    text=${text%$'\r'}
    [[ $text == Contents ]] && continue
    anchor=$(printf '%s' "$text" | tr '[:upper:]' '[:lower:]' | tr -cd 'a-z0-9 -' | tr ' ' '-')
    # Link text is the heading without its " - YYYY-MM-DD" suffix.
    label=${text% - [0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]}
    printf -- '- [%s](#%s)\n' "$label" "$anchor"
  done
)

if [[ -n $contents ]]; then
  tmp=$(mktemp)
  trap 'rm -f "$tmp"' EXIT

  awk -v contents="$contents" '
    !seen && /^## Contents[[:space:]]*$/ {
      print; print ""; print contents; print ""
      seen = 1; skip = 1
      next
    }
    skip && /^## / { skip = 0 }
    skip { next }
    { print }
  ' "$changelog" >"$tmp"

  mv "$tmp" "$changelog"
  trap - EXIT
fi

echo "Rolled '${title}' into '${heading}' in ${changelog}"
