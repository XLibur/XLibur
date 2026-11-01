# Spec 57 — Release Benchmark History (variance between releases, on a GitHub runner)

**Area:** CI / performance observability
**Effort:** M
**Dependencies:** None. Relates to spec 19 (benchmark hotspot survey) and the external XLBench repo.
**Status:** Agreed (grilling session, 2026-11-02)

## Summary

Run a small, memory-lean set of XLibur benchmarks on a GitHub-hosted runner at each release, store one
record per release, and report the variance against earlier releases in `docs/performance.md`.
A shared runner is too noisy to compare timings reliably, so the goal is to *detect significant
change between releases*, not to publish authoritative numbers. Allocated bytes are deterministic and
carry the tight threshold; time is indicative only.

## Decisions

### Scope
- The in-repo job measures **XLibur-internal scenarios from source**: demand-driven read, first edit
  after load, shared-formula write, template and properties round trip, formula evaluation, streaming
  write, used-cell enumeration.
- Cross-library scenarios stay with XLBench (https://github.com/jafin/XLBench), which benchmarks the
  published NuGet package and has its own indicative CI workflow (`results-ci.md`) and `CHANGELOG.md`.
  `performance.md` links to both. **No cross-repo dispatch** (would need a PAT secret); revisit later.

### Benchmark set
- Opt in with `[BenchmarkCategory("CI")]`; the job runs `--anyCategories CI`.
- Each tagged benchmark runs at 10K–20K rows (or reduced `[Params]`). Nothing over ~100K rows, no
  picture benchmarks, no `LargeSheetRoundTrip`.
- Config: `ShortRunJob`, `MemoryDiagnoser`, in-process toolchain, **net10.0 only**, `ubuntu-latest`
  with `libfontconfig1`, 30-minute timeout, JSON exporter.

### Triggers
- New `.github/workflows/performance.yml` on `release: published` and `workflow_dispatch`. Never
  blocks `release.yml`.
- Manual runs are stored with `manual: true` and excluded from baselines, flags and the last-5 table.
- PRs touching `XLibur.Benchmarks/**` run a smoke job (`--anyCategories CI --job Dry`) so the tagged set
  cannot rot. It stores nothing and flags nothing.

### Records
- One JSON record per release on an orphan branch `perf-history`: `schema: 1`, version, commit, runner
  image, core count, per-benchmark mean, allocated bytes, Gen0/1/2, peak working set, and a
  per-benchmark **fingerprint** (method, parameters, hash of the benchmark source file).
- History is kept indefinitely. Pre-release tags are stored but excluded from baselines.
- The tool refuses to compare across schema versions it does not know.

### Comparison and flags
- Allocations: flag when more than **5%** above the previous release.
- Time: flag when more than **30%** above the median of the last 5 releases; labelled "indicative".
- A changed fingerprint shows "re-baselined" for that benchmark and is not flagged that release.
- The first run seeds history: no backfill, no flags.
- Nothing is a hard gate.

### Reporting
- `tools/PerfReport` (small .NET tool, TUnit tests): reads BDN JSON plus runner metadata, appends the
  record, compares, and generates `docs/performance.md` on `perf-history` (never hand-edited).
- Table per benchmark: last 5 releases, mean, allocated, delta vs previous, flag column. The same table
  goes to the job summary.
- A flagged allocation increase opens or updates one tracking issue labelled `performance`.

## Work plan

| # | Task | Size |
|---|------|------|
| 1 | `tools/PerfReport`: record model (schema 1), BDN JSON reader, compare/flag logic, markdown writer, tests | M |
| 2 | Tag the lean benchmark set `[BenchmarkCategory("CI")]`, trimming sizes where needed; confirm peak memory | S |
| 3 | `performance.yml`: release + dispatch run, smoke job on PRs, push to `perf-history`, job summary, issue on flag | M |
| 4 | Docs: README/docs-website link to `performance.md`, CHANGELOG entry, CLAUDE.md CI/CD note | S |

## Acceptance criteria

1. `dotnet run --project tools/PerfReport` over a fixture BDN JSON appends a schema-1 record and
   regenerates the markdown, with unit tests for: first run (no flags), allocation flag at >5%, time
   flag at >30% over a 5-release median, fingerprint change = re-baselined, manual and pre-release
   records excluded from baselines, unknown schema refused.
2. The tagged `CI` set runs end to end on a Linux runner within the 30-minute timeout.
3. `performance.yml` never triggers `release.yml` behaviour and cannot fail a release.
4. The PR smoke job runs only when `XLibur.Benchmarks/**` (or the tool/workflow) changes.
5. No new secrets are required.

## Implementation notes

- **Trigger.** GitHub does not fire `release: published` for a release that `GITHUB_TOKEN` publishes, which
  is how `release.yml` publishes. `release-publish.yml` therefore starts `performance.yml` with
  `workflow_dispatch` (input `tag`) in a job that runs after the release, does not wait for the benchmark run
  and has `continue-on-error: true`. A dispatch without a tag is a manual run. The `release` trigger stays for
  a release a person publishes by hand. A tag pushed by hand is not covered; dispatch Performance with the tag.
- **Order.** Records are ordered by semantic version, not by recording time, so a re-run of an old release or
  a hotfix tag cut after a newer release compares with the release below it. Versions that do not parse
  (manual runs) sort last, by recording time.
- **Fingerprint limits.** The fingerprint is method, parameters and a hash of the benchmark class's file plus
  the helper files it names (fixtures, data builders; a file that holds a benchmark class is never a helper),
  with comments and whitespace removed. It is approximate: code in the library under test is not hashed, and
  a helper reached without the class naming one of its types, or through reflection, is missed.
- **Issue.** One `performance` issue. Its body is a pointer to the history; each flagged release (not a
  pre-release, not a manual run) adds a comment.
