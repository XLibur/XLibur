---
name: changelog-generator
description: Generate or update CHANGELOG.md files from git history. Use this skill whenever the user asks to create a changelog, generate release notes, summarize commits for a release, document changes since the last version/tag, or update an existing CHANGELOG.md — even if phrased informally like "what changed since v2.4", "write up the release notes", or "add the recent commits to the changelog". Also trigger when the user mentions "Keep a Changelog", version release documentation, or wants commit history translated into user-facing language.
---
 
# Changelog Generator
 
Generate and maintain professional, user-facing `CHANGELOG.md` files by analyzing git history, categorizing changes, and translating developer commits into customer language.
 
## Workflow Overview
 
1. **Determine the range** — figure out which commits to include
2. **Scan git history** — collect commits with full context, plus author and PR for each
3. **Filter noise** — exclude internal-only changes
4. **Categorize** — group by change type (emoji sections), then by feature area within each
5. **Translate** — rewrite commit messages as user-facing entries, each attributed to its PR and author
6. **Format & write** — create or update `CHANGELOG.md` following the conventions below
7. **Review with the user** — show the draft before finalizing
## Step 1: Determine the Commit Range
 
Resolve what the user means by "since last release" / "for version X" / "past week":
 
```bash
# Most recent tag (usually the last release)
git describe --tags --abbrev=0
 
# All tags, newest first, with dates
git tag --sort=-creatordate --format='%(refname:short)  %(creatordate:short)'
 
# Check whether CHANGELOG.md already exists and read its latest entry
# (its most recent version heading tells you where to resume from)
```
 
Common range patterns:
 
| User request | Git range |
|---|---|
| "since last release" | `<latest-tag>..HEAD` |
| "for version 2.5.0" | `<previous-tag>..v2.5.0` (or `..HEAD` if not yet tagged) |
| "past week" | `--since="1 week ago"` |
| "between 2.4 and 2.5" | `v2.4.0..v2.5.0` |
 
If no tags exist and no existing changelog gives a starting point, ask the user what range they want rather than dumping the entire history.
 
## Step 2: Scan Git History
 
Collect commits with enough context to write meaningful entries:
 
```bash
# Subject + body + author + date, machine-parseable
git log <range> --no-merges --pretty=format:'%H%x09%ad%x09%an%x09%s' --date=short
 
# For commits whose subject alone is unclear, pull the full message and files touched
git show <hash> --stat --pretty=format:'%s%n%n%b'
```
 
Read commit bodies, not just subjects — the body often explains the user impact that the subject omits. If a commit references a PR or issue number (`#123`), keep the reference; every entry is attributed to it (see Step 5).

For large ranges (100+ commits), work in batches and consider whether merge commits of PRs (`--merges --first-parent`) give a cleaner unit of change than individual commits — one PR usually equals one changelog entry.

### Collect attribution

Every entry carries a PR link and an author, so gather both while you scan.

```bash
# Author and PR of record, straight from GitHub — not inferred from the git author
gh pr view <number> --json number,author,url --jq '"\(.number)\t\(.author.login)\t\(.url)"'
```

Resolve these before writing a single link:

- **Which repo do the PRs live in?** Check `git remote -v`. In a fork, `origin` is the personal
  remote and the PRs live on `upstream` — linking to `origin` produces 404s. Confirm with one
  `gh pr view`: its `url` field is authoritative.
- **Who authored it?** Squash merges rewrite the git author, so `%an` can name the merger rather
  than the contributor. `gh pr view --json author` is the reliable source. Fall back to `%an` only
  when the PR is unavailable (no `gh`, no network, or the commit landed without a PR).
- **Updating an existing changelog?** Attribute entries that are already in the file by tracing
  which commit introduced each line, rather than guessing from subject matter:
  ```bash
  git log <range> --reverse --format='@@@ %h %s' -p -- CHANGELOG.md | grep -E "^@@@|^\+- "
  ```
  One feature built across several PRs cites all of them.

If a change landed with no PR (direct push), link the commit SHA instead and say so when you
present the draft.
 
## Step 3: Filter Noise
 
Exclude commits that don't affect users. Typical exclusions:
 
- Test-only changes (`test:`, `tests/` only file changes)
- Refactoring with no behavior change (`refactor:`, "cleanup", "rename internal")
- CI/build tooling (`ci:`, `build:`, `.github/`, pipelines, linting config)
- Dependency bumps with no user impact (but **include** bumps that fix security issues or unlock features)
- Formatting, typo fixes in code comments, `chore:` commits
- Merge commits and version-bump commits ("bump version to 2.5.0")
- Work-in-progress noise ("wip", "fix typo from previous commit" — fold into the parent change)
**Judgment calls:** a refactor that improves performance noticeably, or a dependency bump that patches a CVE, *is* user-facing. When in doubt, include it in a draft and flag it for the user to decide.
 
Multiple commits for one feature (initial commit + follow-up fixes + review feedback) collapse into a **single** changelog entry.
 
## Step 4: Categorize — Two Levels

Entries are grouped **twice**: by change type first, then by feature area within each type. A reader
scanning for "what will break" reads one section; a reader who only cares about charting scans one
subheading. Both groupings are required — a flat list of 30 bug fixes is not a changelog.

### Level 1 — change type (`###`, always emoji)

In this order, omitting empty sections:

1. **⚠️ Breaking Changes** — anything requiring user action to upgrade (removed APIs, changed defaults, migration steps). Always first, always prominent. Anything the reader must edit code for carries a before/after example (Step 5).
2. **🔒 Security** — vulnerability fixes. State the impact and affected versions if known; don't include exploit details.
3. **✨ New Features** — new capabilities users can now use.
4. **⚡ Performance** — measured speed or resource wins. Quote the numbers and the workload they came from. (Merge into **💪 Improvements** — usability, better errors, expanded limits — if the release has few of either.)
5. **🐛 Bug Fixes** — things that were broken and now work.
6. **🗑️ Deprecations** — features marked for future removal, with the recommended alternative and a before/after example of the swap (Step 5).

The emoji are part of the heading, not decoration to be dropped. Detection hints: conventional-commit
prefixes map naturally (`feat:` → Features, `fix:` → Bug Fixes, `perf:` → Performance, `feat!:` /
`BREAKING CHANGE:` footer → Breaking Changes), but don't rely on them exclusively — read the actual
change.

### Level 2 — feature area (`####`, no emoji)

Within each type section, group by the part of the product the change touches — the areas a user
would recognise, derived from *this* codebase rather than a fixed list. For a spreadsheet library
that might be *Charts*, *Formulas and references*, *Text and number parsing*, *Rich text and shared
strings*, *Colours and styles*, *Conditional formatting*, *Security and encryption*, *Saving*. For a
web service it might be *Authentication*, *Search*, *Billing*, *Admin API*.

Rules:

- **Use the same area vocabulary across every section** so *Charts* under Bug Fixes is recognisably
  the same area as *Charts* under New Features. Keep a consistent area ordering too.
- **One PR can produce areas in several type sections** — a PR that adds an enum member, renumbers
  it and obsoletes the old name puts an area under New Features, Breaking Changes *and*
  Deprecations. That repetition is correct; type wins over area at the top level.
- **Name the area for what it affects, not for the PR that changed it.** If an entry's substance is
  half colours and half conditional formatting, name it *Colours and conditional formatting*.
- **Keep single-entry subheadings** for structural consistency, unless the user asks otherwise.
- **Don't invent areas to look thorough.** Three or four well-chosen areas beat ten of one entry
  each. If a section's entries genuinely share one area, that section gets one subheading.

*(If the project's existing CHANGELOG.md uses different section names, match the existing style
instead — and see Step 6 on leaving already-released sections alone.)*
 
## Step 5: Translate Technical → User-Friendly
 
Rewrite each entry for the **user reading the changelog**, not the developer who wrote the commit.
 
Rules:
- Lead with the user benefit or visible behavior, not the implementation
- Use plain language; drop internal jargon, class names, and file paths unless the audience is developers using those APIs
- Start entries with a verb in past tense or "Added/Fixed/Improved" style — pick one style and be consistent
- Keep entries to one or two lines; link to docs or PRs for detail
- For breaking changes, say **what breaks and what to do about it**
- **Show a before/after code example whenever the change requires the reader to edit code** (see below)
- **End every entry with its PR link and author** (see below)

### Attribution

Each entry closes with the PR that delivered it and who wrote it:

```markdown
- **Fixed a crash when signing in after a long period of inactivity.** … ([#238](https://github.com/org/repo/pull/238) by [@handle](https://github.com/handle))
```

- Link the **PR**, not the merge commit — `#238` is the conventional target and carries the review
  discussion. Use a commit SHA only when there is no PR.
- Use the author's **GitHub handle**, linked, rather than their display name.
- An entry collapsed from several PRs lists them all before a single `by`:
  `([#220](…), [#221](…), [#222](…) by [@handle](…))`
- Absolute URLs, against the repo the PRs actually live in (Step 2).
- If everything in the range is by one author, keep the per-entry attribution anyway — it stays
  correct as outside contributions arrive.

### Migration examples

**If an entry requires the reader to change their code, show the change.** Prose alone ("use
`RawColor()` instead") makes every reader translate the sentence back into code themselves, and each
of them can get it wrong. A two-line before/after removes that step.

This applies to any entry with a migration, wherever it sits — most often ⚠️ Breaking Changes and
🗑️ Deprecations, but also a ✨ New Feature that supersedes an older call, or a 🐛 Bug Fix whose
correct behaviour needs a different call site.

````markdown
- **`IndexColor` is replaced by `RawColor()`.** The property returned a palette index that silently
  went stale when the theme changed; the method resolves against the active theme at call time.
  ([#312](https://github.com/org/repo/pull/312) by [@handle](https://github.com/handle))

  ```csharp
  // Before
  var c = cell.Style.Font.IndexColor;

  // After
  var c = cell.Style.Font.RawColor();
  ```
````

Rules for the example:

- **Smallest thing that compiles.** One or two lines each side, `// Before` and `// After`. No
  surrounding class, no `using` block, no invented business context.
- **Only the line that changes.** If the migration is a single renamed argument, show that call, not
  the whole method it sits in.
- **Real names from the diff.** Read the actual signature rather than guessing at the new API's
  shape — a plausible-looking wrong example is worse than no example.
- **Label a non-mechanical migration.** If the replacement is not a drop-in — different return type,
  different null behaviour, now throws where it used to return a default — say so in one sentence
  under the example. That is the part the reader cannot infer from before/after.
- **Skip it when there is nothing to type.** A behaviour change that needs no code edit gets prose
  only; an example there is noise.
- **Several entries sharing one migration** get the example once, on the entry that introduces it,
  and the others reference it — or, if the release is a large coordinated migration, collect them in
  an `### Upgrade Guide` section at the end of the version and link to it.

**Example 1:**
Input: `fix(auth): handle null refresh token in TokenService.RenewAsync causing NRE`
Output: `Fixed a crash that could occur when signing in after a long period of inactivity`
 
**Example 2:**
Input: `perf: replace per-row dictionary lookup with cached array indexer in export path`
Output: `Improved export performance — large exports now complete up to 3× faster`
 
**Example 3:**
Input: `feat!: remove legacy /v1/search endpoint`
Output: `**Breaking:** The legacy /v1/search endpoint has been removed. Migrate to /v2/search — see the migration guide.`
 
**Example 4 (collapsing commits):**
Input: `feat: add CSV export`, `fix: CSV export encoding`, `fix: CSV header row missing`
Output: `Added CSV export for report data`
 
### Brand voice
 
Default to a professional, concise, positive tone. Before writing, check for project-specific voice guidance in this priority order:
1. Explicit instructions from the user in this conversation
2. Existing `CHANGELOG.md` entries — match their tone, tense, and entry formatting exactly. This
   governs *prose*: sentence length, whether entries lead with a bolded phrase, how much
   implementation detail they carry. It does not override the two-level emoji structure or
   attribution — those apply to the section you are writing even when older sections lack them.
3. A `CONTRIBUTING.md` or docs style guide in the repo
## Step 6: Format and Write CHANGELOG.md
 
Follow [Keep a Changelog](https://keepachangelog.com) conventions adapted to the sections above:
 
```markdown
# Changelog
 
All notable changes to this project will be documented in this file.

## Contents

- [Unreleased](#unreleased)
- [2.5.0](#250---2026-07-26)
- [2.4.0](#240---2026-06-02)

## [2.5.0] - 2026-07-26

### ⚠️ Breaking Changes

#### Search API

- **Removed the legacy `/v1/search` endpoint.** Migrate to `/v2/search`; the response shape is unchanged apart from the `facets` key. ([#244](https://github.com/org/repo/pull/244) by [@dana](https://github.com/dana))

  ```csharp
  // Before
  var results = await client.GetAsync("/v1/search?q=widget");

  // After
  var results = await client.GetAsync("/v2/search?q=widget");
  ```

### ✨ New Features

#### Reporting

- **CSV export for report data**, including scheduled reports and the saved-view picker. ([#241](https://github.com/org/repo/pull/241) by [@sam](https://github.com/sam))

#### Search API

- **Faceted search** via `/v2/search?facets=`, returning counts per category alongside the results. ([#244](https://github.com/org/repo/pull/244) by [@dana](https://github.com/dana))

### ⚡ Performance

#### Reporting

- **Large exports complete up to 3× faster** (250K-row export, 41s → 13s): the per-row dictionary lookup in the export path is now a cached array indexer. ([#239](https://github.com/org/repo/pull/239) by [@sam](https://github.com/sam))

### 🐛 Bug Fixes

#### Authentication

- **Fixed a crash when signing in after a long period of inactivity.** A null refresh token reached the renewal path instead of triggering a re-login. ([#238](https://github.com/org/repo/pull/238) by [@lee](https://github.com/lee))
```

Rules for writing the file:
- **A `## Contents` table of contents sits at the top**, between the intro and the newest version — see below
- **Newest version at the top**, directly under the table of contents
- Version heading format: `## [X.Y.Z] - YYYY-MM-DD`. If the release isn't tagged/dated yet, use `## [Unreleased]`
- **Updating an existing changelog:** insert the new section above the previous newest version. Never rewrite or reorder existing entries. Preserve the file's existing heading style, bullet style, and link conventions exactly — consistency beats these guidelines.
- **A file whose older releases predate this format:** apply the two-level emoji structure and attribution to the section you are writing, and leave already-released sections untouched. Mixed heading styles across releases are expected and fine. Offer to backfill older sections as a separate pass — it means mapping every historical entry to a PR, which is its own job.

### Table of contents

One flat list of versions, newest first, `Unreleased` included whenever the section exists. Versions
only — do not list the type sections or feature areas beneath them; the point is to jump between
releases, and a TOC that mirrors every heading is longer than the content it indexes.

**Derive each anchor from the heading you actually wrote**, applying GitHub's slug rules: lowercase
it, drop every character that is not alphanumeric, space or hyphen, then turn spaces into hyphens.
Brackets and dots vanish, and the ` - ` before a date becomes three hyphens:

| Heading | Anchor |
|---|---|
| `## Unreleased` | `#unreleased` |
| `## [2.5.0] - 2026-07-26` | `#250---2026-07-26` |
| `## v0.106.0 - 2026-07-25` | `#v01060---2026-07-25` |

Getting this wrong yields a link that silently scrolls nowhere, so check the derivation against the
literal heading text rather than against the version number you have in mind.

**Keep it in sync.** Adding a version section means adding its TOC line in the same edit; renaming a
heading (`Unreleased` → `1.0.0 - 2026-08-01`) means updating both the heading and its entry. On any
run that touches an existing changelog, confirm every TOC line still resolves to a heading and that
no heading is missing from the list — a stale TOC is the first thing a reader notices.

**Adding a TOC to a file that has none** is a safe, self-contained improvement: it inserts a block
and changes no existing entry. Do it as part of the update and mention it. If the file already has a
TOC in a different style (nested, table, bullet-free), extend that style rather than replacing it.
- If the existing file maintains link references at the bottom (`[2.5.0]: https://github.com/.../compare/v2.4.0...v2.5.0`), add the new one and update the `[Unreleased]` comparison link
- No trailing whitespace; one blank line between sections
## Step 7: Review Before Finalizing
 
Present the draft to the user before (or immediately after) writing the file, and flag:
- Entries you were unsure whether to include or exclude
- Commits you couldn't confidently translate (ask what the user-facing impact was)
- Anything that looks like it might be a breaking change but wasn't marked as one
- Feature areas you had to invent, or entries that straddle two areas
- Any entry you could not attribute (no PR, or an author you had to take from `%an`)
- Which repo you linked PRs against, when the project has more than one remote
- Any migration example whose replacement API you could not verify against the diff, and any
  migration you judged non-mechanical — those are the two the user most needs to check

Report the section/area layout and the entry count so the user can see the shape at a glance without
re-reading the file. After restructuring an existing section, verify the entry count is unchanged —
regrouping must not drop entries.
## Edge Cases
 
- **Empty range** (no commits since last release): say so; don't fabricate entries
- **Monorepo**: ask which package/path to scope to, then use `git log <range> -- <path>`
- **No CHANGELOG.md and user asked to "update" it**: create it, seeded with the requested range only — offer to backfill older releases as a separate step
- **Release notes vs changelog**: "release notes" for a single version = the same content as one changelog section, but delivered standalone (e.g., for GitHub Releases). Ask where it should go if unclear.
- **Sensitive info**: never copy internal ticket URLs, credentials, or internal system names from commit bodies into a public changelog