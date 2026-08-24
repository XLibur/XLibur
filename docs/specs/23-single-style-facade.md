# Spec 23 — One implementation per style interface: delete the deferred facades

**Area:** Architecture · Refactor · **Defect** (batch path writes wrong borders)
**Effort:** M (~1 week)
**Dependencies:** None hard. Shares files with **spec 20** (`XL*Key.cs` layout) — see Conflicts.
Spec 11 task 4 (bulk style propagation) is **done** and is the code this spec builds on.
**Status:** Done — see [Results](#results).

## Goal

Make `IXLStyle.Batch` a **flush policy on the one style facade** instead of a **second parallel
implementation** of seven interfaces, deleting 713 lines and closing a defect that exists only
because the two implementations must agree by hand.

## Why this spec exists

Every style interface has exactly two implementations:

| Interface | Real facade | Deferred twin |
|---|---|---|
| `IXLStyle` | `XLStyle.cs` (395) | `XLDeferredStyle.cs` (99) |
| `IXLBorder` | `XLBorder.cs` (738) | `XLDeferredBorder.cs` (201) |
| `IXLFont` | `XLFont.cs` (463) | `XLDeferredFont.cs` (121) |
| `IXLAlignment` | `XLAlignment.cs` (377) | `XLDeferredAlignment.cs` (131) |
| `IXLFill` | `XLFill.cs` (218) | `XLDeferredFill.cs` (79) |
| `IXLNumberFormat` | `XLNumberFormat.cs` (181) | `XLDeferredNumberFormat.cs` (45) |
| `IXLProtection` | `XLProtection.cs` (171) | `XLDeferredProtection.cs` (37) |
| | **2,543 lines** | **713 lines** |

Which one a caller gets is decided at `XLStyle.cs:212`:

```csharp
public IXLStyle Batch(Action<IXLStyle> modifications)
{
    if (!IsCellContainer)
    {
        // For ranges: fall back to normal behavior (each property triggers ModifyStyle)
        modifications(this);
        return this;
    }

    // For cells: use a deferred style that accumulates key changes
    var deferred = new XLDeferredStyle(Value.Key);
    modifications(deferred);
    ...
}
```

So `cell.Style.Border.X = v` and `cell.Style.Batch(s => s.Border.X = v)` run **different code** for
the same assignment, on the same container, through the same interface.

### The defect this has already produced

`XLBorder.InsideBorder` (`XLBorder.cs:195`) branches on the container:

```csharp
if (_container is null or XLWorksheet)          // note: XLCell is NOT in this list
    Modify(k => k with { TopBorder = value, BottomBorder = value,
                         LeftBorder = value, RightBorder = value });
else
    foreach (var r in _container.RangesUsed)
        using (new RestoreOutsideBorder(r))
            foreach (var cell in r.Cells())
                ((XLBorder)cell.Style.Border).Modify(k => k with { /* all four */ });
```

For a **cell** container the `else` branch runs. `RangesUsed` for a cell is a 1×1 range, so
`RestoreOutsideBorder` captures all four edges and restores all four on dispose — the net effect is
**no change**, which is correct: a single cell has no interior edges.

`XLDeferredBorder.InsideBorder` (`XLDeferredBorder.cs:45`) has no container and cannot do any of
that:

```csharp
public XLBorderStyleValues InsideBorder
{
    set => Key = Key with { TopBorder = value, BottomBorder = value,
                            LeftBorder = value, RightBorder = value };
}
```

That is byte-identical to its own `OutsideBorder` setter three lines above. So:

```csharp
cell.Style.Border.InsideBorder = XLBorderStyleValues.Thick;              // no-op   (correct)
cell.Style.Batch(s => s.Border.InsideBorder = XLBorderStyleValues.Thick); // all four (wrong)
```

The same holds for `InsideBorderColor`. **Exactly two properties diverge**; task 1 proves it and
task 4 deletes the code that can diverge.

### The duplication is already being worked around

`XLibur.Tests/Excel/Styles/XLAlignmentKeyTests.cs:113` says, of validation:

> `XLDeferredAlignment` is a second, independent facade over the same key, reached through
> `IXLStyle.Batch` for a cell rather than through the ordinary `XLAlignment` facade. Validation
> lives on the key rather than on either facade **precisely so a caller cannot reach an unguarded
> path through this one instead.**

Pushing validation down to the key is a sound defence against the duplication. It is not a fix for
it, and it does not help with `InsideBorder`, whose correctness is about the *container*, not the key.

`XLibur.Tests/Excel/Styles/BatchStyleTests.cs:140` already asserts the intended property —
`Batch_MatchesIndividualPropertySets` — but only for `OutsideBorder`. The one border property that
diverges is the one the test does not cover.

## The design: batching is a mode, not a second object graph

`XLStyle` gains a nullable pending key. While it is set, the component `Modify*` fast paths write
into it instead of resolving a style value and pushing it to the cell. At flush, one resolution and
one `SetStyleValue`.

```
Before                                   After
──────                                   ─────
XLStyle ──cached──> XLBorder ─┐          XLStyle ──cached──> XLBorder ─┐
                              ├─ 7 IXL*                                ├─ 7 IXL*
XLDeferredStyle ──> XLDeferredBorder ─┘   (deleted)
```

The facades are **not modified in their behaviour**. That is the whole point: `InsideBorder`'s
container-aware logic starts applying on the batch path because there is no longer a second path.
The defect is fixed by deletion, not by a patch.

### Why the facades stop caching an interned value

`XLBorder.SetKey` (`XLBorder.cs:589`) currently does:

```csharp
private void SetKey(XLBorderKey newKey)
{
    _style.ModifyBorder(newKey);
    _value = _style.Value.Border;
}
```

While batching, `_style.Value` does not change, so `_value` would go stale and the facade's getters
would report pre-batch values mid-batch. The fix is to stop caching: the facade reads its key from
the style, and derives colours from the key exactly as `XLDeferredBorder` already does
(`XLColor.FromKey(ref colorKey)`). This removes a repository round-trip from the non-batch path too —
`SetKey`'s own remarks already describe `_value = _style.Value.Border` as a wasted lookup on a
transition-cache hit.

## Non-goals

- **No public API change.** `IXLStyle.Batch` keeps its signature and its contract. Everything
  changed is `internal`.
- **No change to range batching.** `Batch` on a non-cell container still runs `modifications(this)`
  directly; this spec does not extend batching to ranges.
- **No key layout changes.** That is spec 20's territory.
- **No new validation.** Key-level validation stays exactly where it is.

## Current state

Verified against the tree at `d05b0753` (2026-08-23).

- `XLStyle.Batch` — `XLibur/Excel/Style/XLStyle.cs:211-233`
- `XLStyle.ModifyFont` / `ModifyBorder` / … — the per-component cell fast paths with the transition
  cache, `XLStyle.cs:88-180`
- `XLStyle.IsCellContainer` — `XLStyle.cs:242`
- Cached sub-facades — `XLStyle.cs:244-250`
- `XLBorder.SetKey` / `Modify` / `SyncValue` — `XLBorder.cs:589`, `:610`, `:124`
- `XLBorder.InsideBorder` / `InsideBorderColor` — `XLBorder.cs:195`, `:232`
- Existing tests: `XLibur.Tests/Excel/Styles/BatchStyleTests.cs` (203 lines, 9 tests)

Only one test file in the whole suite mentions the deferred types, and it does so in a comment.

## File structure

```
XLibur/Excel/Style/XLStyle.cs               modified — gains the pending-key mode
XLibur/Excel/Style/XLBorder.cs              modified — reads its key from the style
XLibur/Excel/Style/XLFont.cs                modified — as above
XLibur/Excel/Style/XLFill.cs                modified — as above
XLibur/Excel/Style/XLAlignment.cs           modified — as above
XLibur/Excel/Style/XLNumberFormat.cs        modified — as above
XLibur/Excel/Style/XLProtection.cs          modified — as above
XLibur/Excel/Style/XLDeferredStyle.cs       deleted
XLibur/Excel/Style/XLDeferredBorder.cs      deleted
XLibur/Excel/Style/XLDeferredFont.cs        deleted
XLibur/Excel/Style/XLDeferredFill.cs        deleted
XLibur/Excel/Style/XLDeferredAlignment.cs   deleted
XLibur/Excel/Style/XLDeferredNumberFormat.cs deleted
XLibur/Excel/Style/XLDeferredProtection.cs  deleted
```

## Interfaces

New internal members on `XLStyle`, consumed by the six component facades:

```csharp
internal sealed class XLStyle : IXLStyle
{
    /// <summary>Non-null while a <see cref="Batch"/> is accumulating. Holds the pending key.</summary>
    private XLStyleKey? _batchKey;

    /// <summary>True while a batch is accumulating; component fast paths write to the pending key.</summary>
    internal bool IsBatching => _batchKey.HasValue;

    /// <summary>
    /// The style key as the facades must see it: the pending batch key while batching, the resolved
    /// value's key otherwise. Every component facade reads its own slice of this.
    /// </summary>
    internal XLStyleKey CurrentKey => _batchKey ?? Value.Key;

    internal XLBorderKey CurrentBorderKey => CurrentKey.Border;
    internal XLFontKey CurrentFontKey => CurrentKey.Font;
    internal XLFillKey CurrentFillKey => CurrentKey.Fill;
    internal XLAlignmentKey CurrentAlignmentKey => CurrentKey.Alignment;
    internal XLNumberFormatKey CurrentNumberFormatKey => CurrentKey.NumberFormat;
    internal XLProtectionKey CurrentProtectionKey => CurrentKey.Protection;
}
```

Each existing `Modify*(XL*Key)` fast path gains a batching branch at the top and is otherwise
unchanged.

## Global constraints

- Warnings are errors; nullable enabled.
- Branch per task; never commit to main. Commit prefix `test:` for task 1, `refactor:` for 2–4,
  `perf:` for task 5 if it changes anything.
- No compound shell commands (`&&`, `;`) in agent tool calls.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Style subset: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Style*/*"`
- Border subset: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/BorderTests/*"`
- Use `--treenode-filter`, never `--filter`. Never filter at solution level.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Prove the divergence — exhaustive batch-vs-direct parity test | S | New test **fails** on `InsideBorder`/`InsideBorderColor` |
| 2 | `XLStyle` gains the pending-key batching mode | M | Suite green; batching mode unused so far |
| 3 | Facades read their key from the style instead of a cached value | M | Suite green |
| 4 | Route `Batch` through the mode; delete the seven deferred types | S | Task 1's test **passes**; suite green |
| 5 | Confirm batching still pays for itself | S | Benchmark within noise of baseline, or spec reverted |

Tasks are strictly ordered. Do not start task 4 before 2 and 3 are green.

---

### Task 1 — Prove the divergence

This task adds a **failing** test. It is the only task in this spec that is allowed to end red.

**Files:**
- Modify: `XLibur.Tests/Excel/Styles/BatchStyleTests.cs`

**Interfaces:**
- Produces: `Batch_and_direct_assignment_agree_for_every_border_property`, the gate task 4 turns green.

- [ ] **Step 1: Write the parity test across every border property**

`Batch_MatchesIndividualPropertySets` at line 140 already asserts this property for `OutsideBorder`.
Add its exhaustive sibling to `XLibur.Tests/Excel/Styles/BatchStyleTests.cs`:

```csharp
/// <summary>
/// Every IXLBorder property must reach the same cell style whether it is assigned directly or
/// inside a Batch. Before spec 23 these ran through two independent implementations of IXLBorder —
/// XLBorder and XLDeferredBorder — and InsideBorder/InsideBorderColor disagreed: a single cell has
/// no interior edges, so the direct path is a no-op, while the deferred path set all four.
/// </summary>
[Test]
[Arguments("OutsideBorder")]
[Arguments("InsideBorder")]
[Arguments("LeftBorder")]
[Arguments("RightBorder")]
[Arguments("TopBorder")]
[Arguments("BottomBorder")]
[Arguments("DiagonalBorder")]
public async Task Batch_and_direct_assignment_agree_for_every_border_property(string property)
{
    const XLBorderStyleValues value = XLBorderStyleValues.Thick;

    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Sheet1");

    var direct = ws.Cell("A1");
    ApplyBorder(direct.Style.Border, property, value);

    var batched = ws.Cell("B1");
    batched.Style.Batch(s => ApplyBorder(s.Border, property, value));

    await Assert.That(BorderSignature(batched)).IsEqualTo(BorderSignature(direct));
}

/// <summary>Colour variants of the same parity property.</summary>
[Test]
[Arguments("OutsideBorderColor")]
[Arguments("InsideBorderColor")]
[Arguments("LeftBorderColor")]
[Arguments("RightBorderColor")]
[Arguments("TopBorderColor")]
[Arguments("BottomBorderColor")]
[Arguments("DiagonalBorderColor")]
public async Task Batch_and_direct_assignment_agree_for_every_border_colour(string property)
{
    var value = XLColor.Red;

    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Sheet1");

    var direct = ws.Cell("A1");
    ApplyBorderColor(direct.Style.Border, property, value);

    var batched = ws.Cell("B1");
    batched.Style.Batch(s => ApplyBorderColor(s.Border, property, value));

    await Assert.That(BorderSignature(batched)).IsEqualTo(BorderSignature(direct));
}

private static void ApplyBorder(IXLBorder border, string property, XLBorderStyleValues value)
{
    switch (property)
    {
        case "OutsideBorder": border.OutsideBorder = value; break;
        case "InsideBorder": border.InsideBorder = value; break;
        case "LeftBorder": border.LeftBorder = value; break;
        case "RightBorder": border.RightBorder = value; break;
        case "TopBorder": border.TopBorder = value; break;
        case "BottomBorder": border.BottomBorder = value; break;
        case "DiagonalBorder": border.DiagonalBorder = value; break;
        default: throw new ArgumentOutOfRangeException(nameof(property), property, null);
    }
}

private static void ApplyBorderColor(IXLBorder border, string property, XLColor value)
{
    switch (property)
    {
        case "OutsideBorderColor": border.OutsideBorderColor = value; break;
        case "InsideBorderColor": border.InsideBorderColor = value; break;
        case "LeftBorderColor": border.LeftBorderColor = value; break;
        case "RightBorderColor": border.RightBorderColor = value; break;
        case "TopBorderColor": border.TopBorderColor = value; break;
        case "BottomBorderColor": border.BottomBorderColor = value; break;
        case "DiagonalBorderColor": border.DiagonalBorderColor = value; break;
        default: throw new ArgumentOutOfRangeException(nameof(property), property, null);
    }
}

/// <summary>
/// Reads the cell's border straight off the style key, so the comparison cannot be satisfied by
/// two facades that merely agree with themselves.
/// </summary>
private static string BorderSignature(IXLCell cell)
{
    var k = ((XLStyle)cell.Style).Key.Border;
    return string.Join('|',
        k.LeftBorder, k.RightBorder, k.TopBorder, k.BottomBorder, k.DiagonalBorder,
        k.LeftBorderColor, k.RightBorderColor, k.TopBorderColor, k.BottomBorderColor,
        k.DiagonalBorderColor, k.DiagonalUp, k.DiagonalDown);
}
```

Add `using System;` and `using XLibur.Excel;` if they are not already present.

- [ ] **Step 2: Run it and record exactly which cases fail**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/BatchStyleTests/*"`

Expected: **12 of 14 cases PASS; `InsideBorder` and `InsideBorderColor` FAIL.**

The failure message shows the direct path leaving the border untouched (all `None`) and the batch
path setting all four edges to `Thick` / `Red`.

**If a different set fails, stop and update this spec before continuing** — the divergence is wider
than the two properties analysed, and tasks 2–4 need to know that.

- [ ] **Step 3: Commit the red test**

```bash
git add XLibur.Tests/Excel/Styles/BatchStyleTests.cs
git commit -m 'test(style): pin batch-vs-direct parity for every border property (spec 23 task 1)'
```

Commit it red, with the failing cases named in the commit body. Task 4 is what turns it green, and
the commit history should show that the defect was demonstrated before it was fixed.

If the repository gates commits on a green suite, add
`[Skip("Fails until spec 23 task 4; see spec 23 task 1")]` to the two failing argument sets and
remove it in task 4.

---

### Task 2 — `XLStyle` gains the pending-key batching mode

**Files:**
- Modify: `XLibur/Excel/Style/XLStyle.cs`
- Test: `XLibur.Tests/Excel/Styles/BatchStyleTests.cs` (unmodified — existing tests are the gate)

**Interfaces:**
- Produces: `XLStyle.IsBatching`, `XLStyle.CurrentKey`, `XLStyle.CurrentBorderKey` and its five
  siblings, consumed by task 3.

- [ ] **Step 1: Add the pending key and the accessors**

Insert into the properties region of `XLibur/Excel/Style/XLStyle.cs`, after `Key`:

```csharp
    /// <summary>
    /// Non-null while <see cref="Batch"/> is accumulating. The component fast paths write their new
    /// component key into it instead of resolving a style value and pushing it to the cell, so a
    /// batch of N property assignments costs one resolution rather than N.
    /// </summary>
    private XLStyleKey? _batchKey;

    /// <summary>True while a batch is accumulating.</summary>
    internal bool IsBatching => _batchKey.HasValue;

    /// <summary>
    /// The key as the component facades must see it. While batching this is the pending key, so a
    /// facade's getter reports what has been assigned so far in the batch rather than the pre-batch
    /// value.
    /// </summary>
    internal XLStyleKey CurrentKey => _batchKey ?? Value.Key;

    internal XLBorderKey CurrentBorderKey => CurrentKey.Border;
    internal XLFontKey CurrentFontKey => CurrentKey.Font;
    internal XLFillKey CurrentFillKey => CurrentKey.Fill;
    internal XLAlignmentKey CurrentAlignmentKey => CurrentKey.Alignment;
    internal XLNumberFormatKey CurrentNumberFormatKey => CurrentKey.NumberFormat;
    internal XLProtectionKey CurrentProtectionKey => CurrentKey.Protection;
```

- [ ] **Step 2: Give every component fast path a batching branch**

`ModifyBorder` becomes — note the `Normalize()` call stays ahead of the branch, so the pending key
holds the same normalized form the repositories would:

```csharp
    internal void ModifyBorder(XLBorderKey newBorderKey)
    {
        newBorderKey = newBorderKey.Normalize();

        if (_batchKey.HasValue)
        {
            _batchKey = _batchKey.Value with { Border = newBorderKey };
            return;
        }

        var transitionHash = (newBorderKey.GetHashCode() * 397) ^ 1;
        Value = Value.GetTransition(transitionHash, in newBorderKey)
                ?? Value.StoreTransition(transitionHash, in newBorderKey, ResolveBorder(newBorderKey));
        ((XLCell)_container!).SetStyleValue(Value);
        return;

        // ... existing local function unchanged
    }
```

Apply the identical four-line branch to `ModifyFont` and to every other `Modify*` component fast
path in the file. Find them with:

Run: `grep -n 'internal void Modify' XLibur/Excel/Style/XLStyle.cs`

Every one of them must gain the branch. A fast path without it silently drops the assignment during
a batch.

- [ ] **Step 3: Give the general `Modify` a batching branch too**

`Modify(Func<XLStyleKey, XLStyleKey>)` is reached from the compound border operations, which run
per-cell through `((XLBorder)cell.Style.Border).Modify(...)`. Those inner facades belong to *other*
cells' styles and are not batching, so they are unaffected — but the outer style's own `Modify` must
honour the mode:

```csharp
    internal void Modify(Func<XLStyleKey, XLStyleKey> modification)
    {
        if (_batchKey.HasValue)
        {
            _batchKey = modification(_batchKey.Value);
            return;
        }

        Key = modification(Key);

        if (_container != null)
        {
            _container.ModifyStyle(modification);
        }
    }
```

- [ ] **Step 4: Build and run the full style suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Style*/*"`
Expected: PASS, unchanged from before this task — `_batchKey` is never set yet, so every branch added
here is dead code at this point. That is deliberate: this task is provably inert.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/Style/XLStyle.cs
git commit -m 'refactor(style): give XLStyle a pending-key batching mode (spec 23 task 2)'
```

---

### Task 3 — Facades read their key from the style

**Files:**
- Modify: `XLibur/Excel/Style/XLBorder.cs`, `XLFont.cs`, `XLFill.cs`, `XLAlignment.cs`,
  `XLNumberFormat.cs`, `XLProtection.cs`

**Interfaces:**
- Consumes: `XLStyle.CurrentBorderKey` and siblings from task 2.

- [ ] **Step 1: Repoint `XLBorder.Key` at the style**

`XLBorder.cs:89` currently reads:

```csharp
    internal XLBorderKey Key => _value.Key;
```

Replace with:

```csharp
    /// <remarks>
    /// Read from the style rather than from a locally cached value, so the facade reports pending
    /// assignments while a batch is accumulating. The style is the single source of truth for the
    /// key; <see cref="_value"/> survives only for the pending-colour bookkeeping below.
    /// </remarks>
    internal XLBorderKey Key => _style.CurrentBorderKey;
```

- [ ] **Step 2: Stop `SetKey` refreshing the cached value**

`XLBorder.cs:589` currently reads:

```csharp
    private void SetKey(XLBorderKey newKey)
    {
        _style.ModifyBorder(newKey);
        _value = _style.Value.Border;
    }
```

Replace with:

```csharp
    private void SetKey(XLBorderKey newKey) => _style.ModifyBorder(newKey);
```

The assignment it drops was already documented on this method as a wasted lookup on a
transition-cache hit, and `Key` now reads through the style, so nothing depends on it.

Apply the same two edits to `Modify` (`XLBorder.cs:610`): drop its trailing
`_value = _style.Value.Border;`.

- [ ] **Step 3: Derive the colour accessors from the key**

Every `XLColor` getter on `XLBorder` that reads `_value` must read `Key` instead. For example:

```csharp
    public XLColor LeftBorderColor
    {
        get
        {
            var colorKey = Key.LeftBorderColor;
            return XLColor.FromKey(ref colorKey);
        }
        set { /* unchanged */ }
    }
```

This is exactly what `XLDeferredBorder` already does, so the shape is proven.

Find every remaining read with:

Run: `grep -n '_value' XLibur/Excel/Style/XLBorder.cs`

What should be left: the `_pending*Color` bookkeeping and `SyncValue`. Nothing else may read
`_value`.

- [ ] **Step 4: Make `SyncValue` compare against the style's key**

`SyncValue` clears the pending colours when the incoming value's key differs from the one held. It
must keep working now that `Key` comes from the style. Change the comparison to capture the key
before assigning:

```csharp
    internal void SyncValue(XLBorderValue value)
    {
        if (!value.Key.Equals(_value.Key))
        {
            _pendingLeftBorderColor = null;
            _pendingRightBorderColor = null;
            _pendingTopBorderColor = null;
            _pendingBottomBorderColor = null;
            _pendingDiagonalBorderColor = null;
        }

        _value = value;
    }
```

This is unchanged — `SyncValue` is called by `XLStyle`'s cached-facade getter with the resolved
value, which is never the batch key, so it keeps comparing resolved keys to resolved keys.

- [ ] **Step 5: Repeat steps 1–3 for the other five facades**

For each of `XLFont`, `XLFill`, `XLAlignment`, `XLNumberFormat`, `XLProtection`:
point `Key` at the matching `XLStyle.Current*Key`, strip the `_value` refresh from its `SetKey` and
`Modify`, and derive any `XLColor` accessor from `Key`. The facades are smaller than `XLBorder` and
none of them has the pending-colour machinery.

- [ ] **Step 6: Build and run the full suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, except task 1's two known-failing cases which are still failing for the same reason.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Excel/Style/XLBorder.cs XLibur/Excel/Style/XLFont.cs XLibur/Excel/Style/XLFill.cs XLibur/Excel/Style/XLAlignment.cs XLibur/Excel/Style/XLNumberFormat.cs XLibur/Excel/Style/XLProtection.cs
git commit -m 'refactor(style): facades read their key from the style, not a cached value (spec 23 task 3)'
```

---

### Task 4 — Route `Batch` through the mode; delete the deferred types

This is the task that turns task 1's test green and deletes 713 lines.

**Files:**
- Modify: `XLibur/Excel/Style/XLStyle.cs:211-233`
- Delete: all seven `XLibur/Excel/Style/XLDeferred*.cs`
- Modify: `XLibur.Tests/Excel/Styles/BatchStyleTests.cs` — remove the `[Skip]` if task 1 added one

- [ ] **Step 1: Rewrite `Batch`**

```csharp
    /// <inheritdoc/>
    public IXLStyle Batch(Action<IXLStyle> modifications)
    {
        if (!IsCellContainer)
        {
            // For ranges: fall back to normal behavior (each property triggers ModifyStyle)
            modifications(this);
            return this;
        }

        // For cells: accumulate into a pending key and resolve once. The facades are the ordinary
        // ones, so container-aware operations — a compound border edit, say — behave exactly as
        // they do outside a batch.
        _batchKey = Value.Key;
        XLStyleKey newKey;
        try
        {
            modifications(this);
        }
        finally
        {
            newKey = _batchKey!.Value;
            _batchKey = null;
        }

        if (!Value.Key.Equals(newKey))
        {
            Value = XLStyleValue.FromKey(ref newKey);
            ((XLCell)_container!).SetStyleValue(Value);
        }

        return this;
    }
```

The `finally` matters: an exception thrown by a caller's lambda — for instance the
`ArgumentOutOfRangeException` that `XLAlignmentKeyTests` asserts on an undefined enum value — must
not leave the style stuck in batching mode. `XLAlignmentKeyTests.Setting_an_undefined_horizontal_value_through_the_batch_facade_throws`
is the existing test that covers this; it must still pass.

- [ ] **Step 2: Delete the seven deferred types**

```bash
git rm XLibur/Excel/Style/XLDeferredStyle.cs XLibur/Excel/Style/XLDeferredBorder.cs XLibur/Excel/Style/XLDeferredFont.cs XLibur/Excel/Style/XLDeferredFill.cs XLibur/Excel/Style/XLDeferredAlignment.cs XLibur/Excel/Style/XLDeferredNumberFormat.cs XLibur/Excel/Style/XLDeferredProtection.cs
```

- [ ] **Step 3: Build and confirm nothing else referenced them**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Expected: no errors. `XLStyle.cs:222` was the only production reference; if the build names another,
that call site also needs routing through the mode.

- [ ] **Step 4: Remove any `[Skip]` from task 1 and run the parity test**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/BatchStyleTests/*"`
Expected: **all 14 cases PASS**, including `InsideBorder` and `InsideBorderColor`.

- [ ] **Step 5: Run the full suite on both frameworks**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0. Pay particular attention to
`XLAlignmentKeyTests.Setting_an_undefined_horizontal_value_through_the_batch_facade_throws` — it now
exercises the real facade's validation rather than the deferred one's, and must still throw
`ArgumentOutOfRangeException`.

- [ ] **Step 6: Update the stale test comment**

`XLibur.Tests/Excel/Styles/XLAlignmentKeyTests.cs:113` describes `XLDeferredAlignment` as "a second,
independent facade over the same key". That type no longer exists. Rewrite the summary to say the
batch path now runs the same facade, and that key-level validation is kept because it guards the key
regardless of which caller reaches it.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m 'refactor(style): one implementation per style interface; delete the deferred facades (spec 23 task 4)'
```

---

### Task 5 — Confirm batching still pays for itself

`Batch` exists for performance. Task 4 changed how it achieves that, so the claim must be re-measured.
This task is empowered to **revert the spec** if batching regresses materially.

**Files:**
- Modify: `XLibur.Benchmarks/` — add a batch-styling case if none exists

- [ ] **Step 1: Find or add the benchmark**

Run: `grep -rn 'Batch' XLibur.Benchmarks --include=*.cs`

If nothing covers it, add a case that styles 50,000 cells with six properties each, once through
`Batch` and once through direct assignment.

- [ ] **Step 2: Measure the merge-base and the branch**

Run on the merge-base commit, then on this branch:

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*Batch*'
```

- [ ] **Step 3: Compare**

Expected: batch styling within noise of the baseline, and still materially faster than direct
assignment. Note the machine's ~40% run-to-run variance — a single run is not evidence. Take three
runs of each and compare medians.

**Decision rule.** If batching is more than 10% slower than the pre-spec baseline on the median of
three runs, the pending-key mode is not reproducing what the deferred graph did. Do not tune blindly:
record the numbers in this spec's Results section, and either fix the specific cause or revert tasks
2–4 and reopen the spec. A correct-but-slower batch is still a decision for the project owner, not
for the implementing agent.

- [ ] **Step 4: Record the numbers in a Results section and commit**

```bash
git add docs/specs/23-single-style-facade.md XLibur.Benchmarks
git commit -m 'perf(style): record batch-styling numbers after the facade merge (spec 23 task 5)'
```

---

## Results

Implemented 2026-08-24 on branch `task/23`. All five tasks landed; no task was reverted.

### What the parity test actually found (task 1)

**13 of the 14 argument sets passed, not 12.** Only `InsideBorder` failed:

```
direct:  None|None|None|None|None|...      (a 1x1 range has no interior edges)
batched: Thick|Thick|Thick|Thick|None|...
```

`InsideBorderColor` passed. `XLDeferredBorder.InsideBorderColor` does write all four edge colours,
exactly as this spec describes — but every edge is styleless, and `XLBorderKey.Normalize` resets a
styleless edge's colour to black when the accumulated key is resolved. The divergence is real and
the deleted code was still wrong; it was simply invisible in the resulting key. The parity set is
therefore **narrower** than analysed, not wider, so tasks 2–4 were unaffected.

### Two deviations from the design as written

Both were forced by the code, and both are documented at their site.

1. **`XLStyle.Key` reads through the pending key**, rather than a separate `CurrentKey` member
   leaving `Key` on `Value.Key`. `RestoreOutsideBorder` reads a cell's key back through
   `XLStyle.Key` mid-edit, so a compound inside-border operation inside a batch would otherwise
   capture the pre-batch key and restore over assignments the same batch had already made.

2. **The component facades read through the style only while it is batching**, keeping `_value` as
   the source of truth otherwise, instead of dropping `_value` entirely. `XLFont(IXLFontBase)`,
   `XLFill(null, d)`, `XLBorder(container, null, d)` and their siblings construct a facade over a
   key that the empty style they fall back to does not hold, so an unconditional read through the
   style reports the default key for every standalone facade. This also leaves the non-batch path
   byte for byte what it was.

### The pending key is not an `XLStyleKey?` field (task 5)

Task 5's benchmark caught the pending-key mode as first written costing **more** than the object
graph it replaced, in two separate ways:

* `XLStyleKey` is a large struct and `XLStyle` is allocated per cell on the ordinary styling path,
  so holding one inline grew every style. `DirectPerCell` — code this spec does not otherwise touch
  — went from 32.5 ms / 29.1 MB to **42.0 ms / 42.1 MB**.
* Every `with` on that key copies the whole struct and re-runs the assigned component's `init`
  accessor, which normalizes and re-hashes it. A six-property batch paid six copies and six
  component hashes that it then paid again at flush. `BatchPerCell` went from 29.5 ms to
  **72.0 ms** — a 2.4× regression, far past this task's 10% rule.

The fix is `XLStyle.PendingKey`: loose component fields behind one reference, assembled into an
`XLStyleKey` exactly once at flush, rented from a one-deep per-thread cache rather than allocated
per batch. The style grows by 8 bytes; a batch allocates nothing in the steady state.

### Numbers

**Read the ratio, not the milliseconds.** This machine drifts hard *between* sessions: `ValueOnly`
in `CellStylingBenchmarks` — a benchmark that touches no style code at all — moved 8.9% across two
sweeps of identical binaries, and pre-spec `BatchPerCell` measured 29.5 ms in one session and
32.3 ms in another. A cross-session millisecond delta under ~10% is therefore not evidence of
anything. The within-run `Batch`/`Direct` ratio cancels that drift, and the allocation columns are
exact.

`BatchStylingBenchmarks`, 50,000 cells × 6 properties, net10.0, medians of three runs, **all three
variants measured back to back in one thermal window**:

| Variant | `BatchPerCell` | `Batch`/`Direct` | Batch allocation |
|---|---|---|---|
| Pre-spec (`1b41cadd`) | 32.33 ms | 0.90 | 32.16 MB |
| Branch, before task 5's second pass | 32.20 ms | 0.93 | 29.49 MB |
| **Branch, final** | **31.41 ms** | **0.90** | **29.49 MB** |

Batch is back to its pre-spec ratio against direct assignment, allocates 8.3% less than the deferred
graph did, and no longer provokes gen1 collections on the batch path. `DirectPerCell` allocation is
29.48 MB against a pre-spec 29.10 MB — a deterministic **+8 bytes per `XLStyle`**, the one `_pending`
reference field. Every `CellStylingBenchmarks` variant that builds a per-cell facade gained exactly
+0.77 MB over 100,000 cells, and the three that do not are byte-identical, which is what makes that
attribution certain rather than inferred. **No revert required.**

### What task 5's second pass changed

The first pass left batching at a 0.93 ratio against a pre-spec 0.90. The cause was that the
deferred twins exposed their component key as a bare field and their setters just wrote it — no
equality guard, no container test, one read — whereas the real facades read `Key` twice per
assignment (guard, then the `with`), and each read had come to cost two null tests: `IsBatching`,
then a second one inside `Current*Key`. Two behaviour-preserving changes closed it:

* `XLStyle` exposes the pending holder itself, so a facade's `Key` getter is one field read and one
  null test. The six `Current*Key` accessors are gone.
* Every setter reads `Key` once into a local and reuses it for both the guard and the `with`.
  `XLFill` gained an `ApplyKeyUpdate` overload taking the already-read key, and its
  `ShouldAdjustPatternTypeForBackgroundColor` takes the key rather than going back through two
  properties that each read it again.

`XLNumberFormat` needed neither: its setters build a fresh key and never read `Key`.

### Acceptance criteria as met

| # | Criterion | Result |
|---|---|---|
| 1 | Seven `XLDeferred*.cs` deleted; 713 lines removed | ✅ exactly 713 |
| 2 | One implementation per style interface | ✅ seven files, one each |
| 3 | All 14 parity argument sets pass | ✅ |
| 4 | No public API change | ✅ `PublicAPI.Unshipped.txt` untouched |
| 5 | `XLAlignmentKeyTests`' two throw-tests still pass | ✅ |
| 6 | Full suite green on net8.0 and net10.0 | ✅ 23,756 passed, 8 skipped |
| 7 | Batch styling within 10% of its pre-spec median | ✅ at parity — 0.90 `Batch`/`Direct`, same as pre-spec |

## Acceptance criteria

1. Seven `XLDeferred*.cs` files deleted; 713 lines removed.
2. Each of the seven style interfaces has exactly **one** implementation. Gate:
   `grep -rlE "class \w+ *: *IXLBorder" XLibur --include=*.cs` returns one file, and the same for
   `IXLFont`, `IXLFill`, `IXLAlignment`, `IXLNumberFormat`, `IXLProtection`, `IXLStyle`.
3. `Batch_and_direct_assignment_agree_for_every_border_property` and its colour sibling pass for all
   14 argument sets.
4. No public API change: `PublicAPI.Unshipped.txt` untouched; `IXLStyle.Batch` keeps its signature.
5. Existing validation behaviour preserved — `XLAlignmentKeyTests`' two throw-tests still pass.
6. Full suite green on net8.0 and net10.0.
7. Batch styling within 10% of its pre-spec median, or the regression recorded and escalated.

## Conflicts

- **Spec 20** (`XLColorKey`/`XLBorderKey`/`XLFontKey`/`XLAlignmentKey` struct layout) edits the key
  types; this spec edits the facades that read and write them. The overlap is the `with`-expressions
  in the facades: if spec 20 renames or restructures a key field, those expressions change. **Land
  one, then rebase the other** — either order works, but do not run them concurrently. Spec 20's
  task 0 size probe is unaffected by this spec.
- Nothing else in `docs/specs/` touches `XLibur/Excel/Style/XL*.cs`. Spec 11 task 4's bulk
  propagation work (`XLStylizedBase.ModifyStyle`/`SetStyle`) is done and is not re-entered here.
