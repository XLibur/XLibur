# Spec 20 — Style key struct sizes: `XLColorKey` is the root, not `XLBorderKey`

**Area:** Performance (write · bulk styling) · Memory (copy cost, GC scanning)
**Effort:** M overall; tasks 1–4 are independently sized, independently ownable and independently revertible
**Dependencies:** Task 0 first (it is the measuring instrument for every other task). Tasks 1→2 are ordered; 3 and 4 are independent of everything. Task 5 gates the whole spec.
**Status:** Proposed. Sizes below are **measured**, not inferred. Everything downstream of "Proposed layout" is a hypothesis until task 0 re-measures it.

## Why this spec exists

A review of `XLBorderKey`/`XLBorderValue` asked whether the border key was oversized. It is — 264 bytes —
but the review was aimed one level too high. `XLBorderKey` is large almost entirely *because of what it
embeds*, and the same embedded type inflates the fill, font and style keys by the same mechanism. Fixing
the border key in isolation would deliver a fraction of the available win and leave three other types
holding the same defect.

**The finding: `XLColorKey` is 48 bytes and accounts for 91% of `XLBorderKey`, 92% of `XLFillKey`, 55% of
`XLFontKey`, and 384 of `XLStyleKey`'s 536 bytes.**

This spec is written as an investigation, not a mandate. The honest position is that the *sizes* are proven
and the *payoff* is not — see "What this will not do" and task 5, which is empowered to revert the lot.

## Measured baselines

Measured 2026-08-08 via `Unsafe.SizeOf<T>()` over `XLibur/bin/Release/net10.0/XLibur.dll`, net10.0 Release.
Probe harness retained at `scratchpad/sizeprobe/` and formalised by task 0.

| Type | Size | Of which is `XLColorKey` | GC-tracked refs |
|---|---:|---:|---:|
| `XLStyleKey` | **536** | 384 (72%) | 10 |
| `XLBorderKey` | **264** | 240 (91%) | 5 |
| `XLFillKey` | **104** | 96 (92%) | 2 |
| `XLFontKey` | **88** | 48 (55%) | 2 |
| `XLColorKey` | **48** | — | 1 |
| `XLAlignmentKey` | 28 | — | 0 |
| `XLNumberFormatKey` | 16 | — | 1 |
| `XLProtectionKey` | 2 | — | 0 |

### Why `XLColorKey` is 48 bytes

Two independent causes, both in `XLibur/Excel/Style/XLColorKey.cs:6-16`:

**1. A discriminated union stored as a flat record.** `ColorType` selects exactly one live arm — both
`GetHashCode` (`:31-53`) and `Equals` (`:61-83`) switch on it and read only that arm — but the layout
reserves space for all of them at once:

```
ColorType   XLColorType : Int32   =  4
Color       System.Drawing.Color  = 24
Indexed     Int32                 =  4
ThemeColor  XLThemeColor : Int32  =  4
ThemeTint   Double                =  8
                                  = 48 (44 + padding)
```

**2. `System.Drawing.Color` is 24 bytes and carries a string.** Measured layout:

```
System.Drawing.Color  size = 24
    name        String      <-- GC-tracked reference
    value       Int64
    knownColor  Int16
    state       Int16
```

So every colour drags a managed pointer. `XLStyleKey` holds **eight** object references purely from
colours, plus `FontName` and `Format` — ten pointers the GC walks on every scan of a live style key, for
data that `Equals` and `GetHashCode` **already reduce to `Color.ToArgb()`** (`XLColorKey.cs:48,69`).

That last sentence is the load-bearing one for task 1: `name`, `knownColor` and `state` are *already
semantically dead* in this type. Nothing in equality or hashing can observe them.

### Blast radius of the colour change (verified 2026-08-08)

`XLColorKey.Color` is read in exactly **two** places and written in **six**, all inside `Excel/Style/Colors/`:

| Site | Use |
|---|---|
| `XLColorKey.cs:69` | `Color.ToArgb() == other.Color.ToArgb()` — already ARGB-only |
| `XLColor_Public.cs:73` | the public `XLColor.Color` getter — the one real boundary |
| `XLColor_Internal.cs:25`, `XLColor_Static.cs:63` | `Color = color` on construction |
| `XLColor_Internal.cs:34`, `XLColor_Static.cs:124` | `Indexed = …` on construction |
| `XLColor_Internal.cs:41,49`, `XLColor_Static.cs:135,145` | `ThemeColor = …` on construction |

No site sets more than one arm. `XLColorKey` is `internal`, so none of this is public API.

## What this will not do

**This is not a heap-footprint win.** Style keys are deduplicated through the style repository, so the
number of live `XLStyleKey` instances is bounded by *distinct styles*, not by cells. Total allocated bytes
on the benchmarks may not move at all.

The claim being tested is narrower: **copy cost and cache behaviour on the bulk-styling path.** Every
`with` expression on `XLStyleKey` copies 536 bytes; `Key = Key with { LeftBorder = value }` in
`XLDeferredBorder.cs:72` copies 264 bytes *per property set*; the repository hashes and compares the key on
every style mutation. The codebase already works around the size with `ref` passing
(`XLBorderValue.FromKey(ref value)`, `XLColor.FromKey(ref colorKey)`), which is circumstantial evidence the
size is felt but is not proof it dominates.

**The noise floor is against us.** Spec 19 measured 9% run-to-run spread on `CreateFormattedAndSave` on
this machine. If this spec delivers less than ~10%, it cannot be distinguished from noise in a single
sitting, and task 5 must say so rather than claim a win. Specs 19-area-3 and 19-area-5 both declined their
own proposed tasks on measurement; that is an acceptable and expected outcome here.

---

## Task 0 — Size-probe test, before touching anything

**Effort:** XS · **Blocks:** every other task

Turn the throwaway probe into a real test so each later task has a before/after number and a regression gate.

- Add `XLibur.Tests/Excel/Style/StyleKeySizeTests.cs` asserting `Unsafe.SizeOf<T>()` for all eight key
  types against the table above.
- The test suite already has the `InternalsVisibleTo` grant, so no reflection is needed — reference the
  types directly.
- Assert **exact** sizes with the current values, and record the intended post-task values in a comment.
  Each task below updates the expected constants as part of its own PR; a size regression then fails CI
  instead of going unnoticed.
- Guard: sizes are ABI-stable for a given runtime but not contractually frozen across major .NET versions.
  If net8.0 and net10.0 disagree, assert per-TFM rather than deleting the test.

**Acceptance:** test passes on net8.0 and net10.0 and fails if any key type grows.

---

## Task 1 — `XLColorKey` 48 → 16 bytes

**Effort:** S · **Depends on:** task 0 · **This is the task that matters; the rest are follow-through.**

Collapse the union and drop `System.Drawing.Color` from storage.

```csharp
internal readonly struct XLColorKey : IEquatable<XLColorKey>
{
    private readonly byte _colorType;   // 1
    private readonly byte _themeColor;  // 1   (12 members)
                                        // 2   padding
    private readonly uint _value;       // 4   ARGB when Color, palette index when Indexed
    private readonly double _themeTint; // 8
                                        // = 16

    public XLColorType ColorType { get => (XLColorType)_colorType; init => _colorType = (byte)value; }
    public XLThemeColor ThemeColor { get => (XLThemeColor)_themeColor; init => _themeColor = (byte)value; }
    public Color Color { get => Color.FromArgb((int)_value); init => _value = (uint)value.ToArgb(); }
    public int Indexed { get => (int)_value; init => _value = (uint)value; }
    public double ThemeTint { get => _themeTint; init => _themeTint = value; }
}
```

**Note the technique, because tasks 2–4 reuse it:** the public-facing property types are unchanged
(`XLColorType`, `XLThemeColor`, `Color`, `int`). Only the *storage* narrows to a byte. This gets the full
packing win with **zero public API change** and no `PublicAPI.Shipped.txt` churn.

Rejected alternative: redeclaring the public enums as `: byte`. Same size outcome, but it is a
binary-breaking change requiring a shipped-API update and a coordinated release, for no additional gain.
Do not do this.

**Consequences to verify, not assume:**

- `Color` and `Indexed` now share `_value`. No current site sets both (table above), but an object
  initialiser that did would silently have last-writer-wins. Add a `Debug.Assert` in the `Indexed` init, or
  a targeted test, so a future caller cannot introduce it quietly.
- **`ToString()` changes for named colours.** `XLColorKey.cs:98` delegates to `Color.ToString()`, which
  prints `Color [Red]` when `knownColor` is set and `Color [A=255, R=255, G=0, B=0]` otherwise.
  `Color.FromArgb` cannot restore the name. Confirm nothing outside diagnostics reads it — in particular
  check the style writers and any test asserting on colour text — and if a name is genuinely needed,
  resolve it at the `XLColor` boundary rather than storing it.
- `ThemeTint` stays `double`. Narrowing it to `float` would reach 12 bytes but changes round-trip
  precision on a value that is written back to the file, and `Equals` already carries an
  `XLHelper.Epsilon` tolerance (`:78`) that would need re-deriving. **Out of scope — do not narrow it.**

**Acceptance:**
- `Unsafe.SizeOf<XLColorKey>() == 16`; the type is `unmanaged` (assert with a generic constraint helper).
- `XLBorderKey` falls to ~104, `XLFillKey` to ~40, `XLFontKey` to ~56, `XLStyleKey` to ~272 with no other
  change. Confirm with task 0's test rather than trusting these estimates.
- Full test suite green, including the colour round-trip and theme-colour tests.

---

## Task 2 — `XLBorderKey` ~104 → ~88

**Effort:** XS · **Depends on:** task 1

Apply the byte-storage technique from task 1 to the five `XLBorderStyleValues` fields (14 members, trivially
byte-sized) in `XLBorderKey.cs:5-27`. Public property types stay `XLBorderStyleValues`.

Do the same for `XLFillKey.PatternType` (`XLFillPatternValues`, 19 members) — though note `XLFillKey` is
already 8-byte aligned around its two colours, so this may measure as 40 → 40. **If it does not shrink,
skip it**; a no-op change that touches a hot type is negative value.

**Acceptance:** `XLBorderKey` ≤ 88 and fully `unmanaged` (zero GC references — both bools and all five
colours are now blittable). Border style round-trip tests green.

---

## Task 3 — `XLFontKey` ~56 → ~40

**Effort:** S · **Independent of tasks 2 and 4**

`XLFontKey.cs:9-31` has four separate `bool` fields (`Bold`, `Italic`, `Strikethrough`, `Shadow`) each
costing a byte plus contributing to padding, and five `int`-backed enums that all fit in a byte
(`XLFontCharSet` maxes at 255 / `Oem`; `XLFontFamilyNumberingValues` at 5; `XLFontScheme`,
`XLFontUnderlineValues`, `XLFontVerticalTextAlignmentValues` are all small).

- Collapse the four bools into one `[Flags] byte`, keeping `bool` public property types.
- Narrow the five enum fields to byte storage.
- `FontSize` (double) and `FontName` (string) stay as they are. `XLFontKey` therefore remains a managed
  type — that is expected and fine.

**Watch:** `GetHashCode` (`:50-72`) and `Equals` (`:33-48`) read every one of these fields. Both must be
updated in lockstep, and the existing `StringComparison.OrdinalIgnoreCase` treatment of `FontName` must be
preserved exactly — it is load-bearing for style deduplication.

**Acceptance:** `Unsafe.SizeOf<XLFontKey>() ≤ 40`; font style tests green; a style-repository test confirms
two fonts differing only in case of `FontName` still dedupe to one entry.

---

## Task 4 — `XLAlignmentKey` 28 → ~8

**Effort:** S · **Independent of everything except task 0**

`XLAlignmentKey.cs:5-21` is the one type whose waste is unrelated to colour. Its three `int` fields exceed
Excel's own limits by a wide margin, and its four bools pad badly:

| Field | Now | Excel's range | Proposed storage |
|---|---|---|---|
| `Indent` | `int` | 0–250 | `byte` |
| `TextRotation` | `int` | 0–180, plus sentinel 255 | `byte` |
| `RelativeIndent` | `int` | −255..255 | `short` |
| `JustifyLastLine`, `ShrinkToFit`, `WrapText` | 3 × `bool` | — | one `[Flags] byte` |
| `Horizontal`, `Vertical`, `ReadingOrder` | already `: byte` | — | unchanged |

**This type is `public`** (unlike every other key in this spec) — so keep the public property types as
`int`/`bool` and narrow only the private storage, exactly as in task 1. No `PublicAPI.Shipped.txt` change.

**Confirm the ranges before narrowing.** The table is from the Excel/OOXML spec, not from this codebase's
validation. If XLibur currently accepts out-of-range values and writes them, narrowing silently truncates.
Check the alignment setters and the styles writer first; if there is no validation, add range checks in the
same PR or leave the field wide.

**Acceptance:** `Unsafe.SizeOf<XLAlignmentKey>() ≤ 12`; alignment round-trip tests green; a test pins
behaviour for an out-of-range indent/rotation (whatever it is decided to be) so truncation cannot be silent.

---

## Task 5 — Measure, then decide whether any of it was worth it

**Effort:** S · **Depends on:** tasks 1–4 · **Empowered to revert.**

Projected end state, to be confirmed rather than assumed:

| Type | Now | After 1 | After 1–4 |
|---|---:|---:|---:|
| `XLColorKey` | 48 | 16 | 16 |
| `XLBorderKey` | 264 | ~104 | ~88 |
| `XLFillKey` | 104 | ~40 | ~40 |
| `XLFontKey` | 88 | ~56 | ~40 |
| `XLAlignmentKey` | 28 | 28 | ~8 |
| `XLStyleKey` | **536** | ~272 | **~224** |
| GC refs in `XLStyleKey` | 10 | 2 | 2 |

Benchmark, in one sitting, A/B against the parent commit:

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*CreateFormattedAndSave*'
```

`CreateFormattedAndSave` (50K × 10, half the rows styled) is the right target — spec 19 measures it at
1,005.5 ms / 322.0 MB and it is the only benchmark whose cost is dominated by style mutation. Also run
`CreateAndSave` as a control: it does almost no styling, so a change there is a signal that something
unintended moved.

**Decision rule, agreed in advance:**

- **≥10% on `CreateFormattedAndSave`** — keep everything, record the numbers in a Results section.
- **Under the 9% noise floor** — keep task 1 anyway (48→16 with 8 fewer GC references and a materially
  simpler type is a maintainability win that stands on its own), and **revert tasks 2–4** rather than
  carrying packing tricks that bought nothing. Say so plainly in Results.
- **Any regression** — revert to the last green task and record which one caused it. Byte-narrowing can
  cost speed where it forces widening conversions in a hot loop; that is a real outcome, not a failure to
  hide.

**Acceptance:** a Results section in this file stating what was measured, on what commit, and which tasks
survived — in spec 19's format, including the disproved parts.

---

## Explicitly out of scope

**`XLStyleKey`'s six memoised hash fields (24 bytes).** Tempting — they are 4.5% of the struct — but they
are the documented fast-reject path (`XLStyleKey.cs:8-16,130-143`) that made style mutation cheap in the
first place. Narrowing them to `short` would weaken rejection and cost more than the bytes are worth.
**Leave them alone.**

**`XLNumberFormatKey` (16 bytes)** — an `int` plus a string reference plus padding. Already minimal.

**`XLProtectionKey` (2 bytes)** — already minimal.

**`XLBorderValue` / `XLFontValue` / `XLStyleValue` etc.** — all measured as `class`, not struct. They hold a
key by reference and are already deduplicated by the repository; they are not a size problem.

**Narrowing `ThemeTint` to `float`** — see task 1. Round-trip precision risk for 4 bytes.

## Ground rules

Standard for this repo (see `docs/specs/README.md`): branch per task, never commit to main, warnings are
errors, nullable annotations required, no compound shell commands. Perf PRs carry before/after numbers.
Line numbers above were verified against `66307307` on 2026-08-08 — re-verify before editing.
