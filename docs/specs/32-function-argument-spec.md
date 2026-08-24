# Spec 32 — One argument spec per function, instead of three encodings that must agree by hand

**Area:** Architecture · Refactor · **Defect class (unrealized)** · **Effort:** L (~8–10 days, 411 call sites) · **Dependencies:** **Spec 30 first** (it owns `FunctionDefinition`'s array path, which reads the fields 32 replaces). No hard dependency otherwise. · **Status:** Proposed.

## Goal

Make a function's argument shape a single fact, written once, in the one place that already knows
it. Today it is written three times per registration — in the `Adapt*` overload chosen, in the
`minParams, maxParams` pair, and in the `AllowRange` + `markedParams` tail — and nothing checks the
three agree.

Two of the three are mechanically derivable from the first. This spec derives them instead of
requiring them, and deletes the `allowRanges`/`markedParams` tail of `RegisterFunction` from all
411 call sites.

**This is the largest blast radius of any spec in the series: 411 live `RegisterFunction` call sites
across 12 files.** And unlike specs 26, 28, 29 and 30 from this review round, **32 has no confirmed
shipped defect behind it.** Task 0's audit found zero wrong `markedParams` indices in the tree
today (see below). Its value is preventing a class of silent defect, not fixing a known one. That is
why it ranks below 26–30 despite touching more code than all of them combined.

## Why this spec exists

### The widest interface in the calc engine, with no test surface

`XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs` is **1,358 lines** and exposes **61
`public static CalcEngineFunction` overloads**:

| Family | Overloads | Lines |
|---|---:|---|
| `Adapt` | 31 | `:30`, `:35`, `:47`, `:59`, `:71`, `:87`, `:107`, `:127`, `:151`, `:163`, `:179`, `:195`, `:219`, `:239`, `:253`, `:276`, `:294`, `:299`, `:311`, `:327`, `:343`, `:357`, `:374`, `:394`, `:415`, `:900`, `:924`, `:987`, `:1039`, `:1055`, `:1118` |
| `AdaptLastOptional` | 14 | `:443`, `:457`, `:473`, `:493`, `:513`, `:529`, `:551`, `:577`, `:593`, `:611`, `:627`, `:958`, `:1138`, `:1159` |
| `AdaptLastTwoOptional` | 7 | `:800`, `:821`, `:842`, `:869`, `:1011`, `:1086`, `:1180` |
| `AdaptIfs` | 2 | `:652`, `:667` |
| `AdaptCoerced`, `AdaptIndex`, `AdaptMatch`, `AdaptSeriesSum`, `AdaptNumberValue`, `AdaptSubstitute`, `AdaptMultinomial` | 7 | `:431`, `:678`, `:696`, `:713`, `:733`, `:755`, `:786` |
| **Total** | **61** | |

The prompt for this spec said "9 one-offs … and one more — find it". The tenth is the **second
`AdaptIfs` overload** at `:667`; `AdaptIfs` is the only one-off with two.

Those 61 overloads cover **55 distinct `Func<>` shapes**. Six shapes appear twice, and all six pairs
are the same shape under `Adapt` and `AdaptLastOptional`, or under `AdaptLastOptional` and
`AdaptLastTwoOptional` — i.e. **every duplicate exists only to say "the last argument is optional"**:

| Shape | Appears at |
|---|---|
| `Func<CalcContext, double, double, ScalarValue>` | `Adapt:71`, `AdaptLastOptional:457` |
| `Func<CalcContext, double, double, double, ScalarValue>` | `Adapt:87`, `AdaptLastOptional:473` |
| `Func<CalcContext, double, double, string, ScalarValue>` | `Adapt:107`, `AdaptLastOptional:1138` |
| `Func<double, double, double, ScalarValue>` | `AdaptLastOptional:529`, `AdaptLastTwoOptional:800` |
| `Func<CalcContext, double, double, bool, ScalarValue>` | `AdaptLastOptional:493`, `AdaptLastTwoOptional:821` |
| `Func<CalcContext, double, double, double, double, double, ScalarValue>` | `AdaptLastOptional:958`, `AdaptLastTwoOptional:1011` |

Converters sit at the bottom of the file: `CoerceToLogical:1224`, `ToNumber:1237`, `ToText:1252`,
`ToScalarValue:1266`, `ToNonLogicalNumber:1280`, `ToSeriesSumCoefficients:1288`,
`ToOptionalNumber:1304`, `GetNonBlankScalars:1312`, `ToCriteria:1333`, in the `#region Value
converters` that opens at `:1219`.

> **Correction to the evidence this spec was commissioned from.** `ToNumber` is at `:1237`, not
> `:1243`. Everything else in the converter list checked out.

**Zero test files reference `SignatureAdapter`.** Verified:

```
grep -rl SignatureAdapter --include=*.cs XLibur.Tests   ->  (nothing)
grep -rl SignatureAdapter --include=*.cs XLibur         ->  12 files, all production
```

The 12 are the 11 `Functions/*.cs` registration files plus `SignatureAdapter.cs` itself. The widest
interface in the engine is exercised only incidentally, through whatever formula a function test
happens to write.

### The file already knows it is the wrong shape

`SignatureAdapter.cs:18-21`, the `S4136` suppression:

```csharp
// S4136 wants the Adapt, AdaptLastOptional and AdaptLastTwoOptional overloads each grouped in one
// run. They are instead ordered by arity within their regions, so an adapter for a new signature is
// found by counting its parameters. Grouping by name would interleave the three families and lose
// that. Unlike S2234 above, no restore for S4136 appears below, so this reaches the whole file.
```

**"found by counting its parameters"** is the discovery rule for a 61-member interface, and it is
written down as a deliberate ordering decision rather than as a problem.

`SignatureAdapter.cs:28`:

```csharp
    // through the value converters below. We can hopefully generate them at a later date, so try to keep them similar.
```

The intent to source-generate is recorded in the file. Nobody has. Section "Source generation vs
hand-written shims" below answers whether they should.

### One fact, three encodings — the worked example

`XLibur/Excel/CalcEngine/Functions/DateAndTime.cs:35`:

```csharp
ce.RegisterFunction("NETWORKDAYS", 2, 3, AdaptLastOptional(NetWorkDays), FunctionFlags.Range, AllowRange.Only, 2); // Returns the number of whole workdays between two dates
```

against the body at `DateAndTime.cs:365`:

```csharp
private static ScalarValue NetWorkDays(CalcContext ctx, ScalarValue startDate, ScalarValue endDate, AnyValue holidays)
```

which selects the overload at `SignatureAdapter.cs:593`:

```csharp
public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, AnyValue, ScalarValue> f)
{
    return (ctx, args) =>
    {
        var arg0Converted = ToScalarValue(args[0], ctx);
        if (!arg0Converted.TryPickT0(out var arg0, out var err0))
            return err0;

        var arg1Converted = ToScalarValue(args[1], ctx);
        if (!arg1Converted.TryPickT0(out var arg1, out var err1))
            return err1;

        var arg2 = args.Length > 2 ? args[2] : AnyValue.Blank;

        return f(ctx, arg0, arg1, arg2).ToAnyValue();
    };
}
```

The three encodings of one fact:

| # | Where | What it says |
|---|---|---|
| 1 | the chosen overload | three parameters; 0 and 1 are scalars; 2 is a raw `AnyValue` defaulting to `AnyValue.Blank` when absent |
| 2 | `2, 3` | minimum two arguments, maximum three |
| 3 | `AllowRange.Only, 2` | parameter 2 accepts a range; 0 and 1 are implicitly intersected first |

**Encodings 2 and 3 are both derivable from encoding 1**, and neither is checked against it.

What happens if the `2` in encoding 3 is wrong — say it reads `1`:

- It compiles. `markedParams` is `params int[]`; any int is legal.
- It passes the only validation there is. `FunctionDefinition.cs:26-27` throws exactly once, when
  `AllowRange.None` is paired with a non-empty `markedParams`. Nothing else is checked — not the
  index range, not the count, not agreement with the delegate's shape.
- It runs. `IntersectArguments` (`FunctionDefinition.cs:126-141`) would intersect argument 2 and
  leave argument 1 alone.
- The result changes only when someone passes a range: `NETWORKDAYS(A1:A5, B1)` would stop being
  intersected to `A1` and would instead arrive at `ToScalarValue`, which silently takes `array[0,0]`
  (`SignatureAdapter.cs:1266-1278`). Same answer here, different answer for
  `NETWORKDAYS(A1:C1, B1)` in a formula on row 1, where intersection would have picked the column
  under the formula rather than the first cell.

So a wrong index is a silently wrong answer on some inputs and the right answer on most. That is the
class of defect this spec makes unrepresentable — not by checking it, but by removing the index.

### An audit of all 53 marked registrations found no defect

Every registration carrying `AllowRange.Only` or `AllowRange.Except` was checked for an index at or
beyond its declared `maxParams`:

```
registrations with out-of-range markedParams index: 0
```

A stronger check — does the marked set agree with which of the delegate's parameters are declared
`AnyValue`/`Array`? — is mechanically decidable for only **21** of the 53, because the rest go
through a one-off adapter or a raw `CalcEngineFunction`. Of those 21, one hit: `INDEX`
(`Lookup.cs:24`), whose `AdaptIndex(Func<CalcContext, AnyValue, List<int>, AnyValue>)` at
`SignatureAdapter.cs:678` collects arguments 1..n-1 into a `List<int>`. `AllowRange.Only, 0` is
correct there; the heuristic false-positived on a variadic collector. **No defect found.**

Say this plainly: **32 is a prevention spec.** It ranks below 26, 28, 29 and 30 for exactly this
reason.

### A fourth encoding, and it is entirely inert

`FunctionFlags` is written at all 411 registrations. It has **exactly one consumer in the whole
library**:

```
grep -rn "FunctionFlags\." XLibur --include=*.cs | grep -v "CalcEngine/Functions/"
  -> XLibur/Excel/CalcEngine/FunctionDefinition.cs:54:
     if (_flags.HasFlag(FunctionFlags.ReturnsArray) && _allowRanges == AllowRange.All)
```

Only `ReturnsArray` is read. `Scalar`, `Range`, `SideEffect`, `Volatile` and `Future` are written
411 times and never read by anything. `FunctionFlags.Range` in particular is documented on
`FunctionFlags.cs` as "At least one of the arguments of the function accepts a range. It means that
implicit intersection works differently" — which is what `AllowRange` actually controls.

The predictable consequence, found by cross-checking the two:

```csharp
// MathTrig.cs:87 — the only registration in the tree where Scalar and a non-None AllowRange co-occur
ce.RegisterFunction("MULTINOMIAL", 1, 255, AdaptMultinomial(Multinomial), FunctionFlags.Scalar, AllowRange.All);
```

`FunctionFlags.Scalar` is documented as "designed for a single value argument"; `AllowRange.All`
says every parameter accepts a range; `AdaptMultinomial` (`SignatureAdapter.cs:786`) genuinely
collects `List<IEnumerable<ScalarValue>>`. The flag is wrong and the behaviour is right, because
nothing reads the flag. **A field that disagrees with reality and cannot be observed to is the
purest possible demonstration of this spec's thesis.**

This spec does not delete `FunctionFlags` — see Non-goals — but it records the finding, because a
later reader who adds a consumer for `FunctionFlags.Range` will inherit 411 unaudited values.

### Where the shape is used at runtime

| Step | Location |
|---|---|
| Arity check, at **parse** time | `FormulaParser.cs:285-313` — `minParams`/`maxParams` via `FunctionRegistry.TryGetFunc(name, out min, out max)` (`FunctionRegistry.cs:37`) |
| Arity check, for the public library | `XLFunctionLibrary.cs:87` |
| Hot path, per function call | `CalculationVisitor.cs:75` → `FunctionDefinition.CallFunction` (`:41`) or `CallAsArray` (`:52`) |
| Implicit intersection | `FunctionDefinition.cs:126-141`, gated on `CalcContext.UseImplicitIntersection`, which is `=> true` unconditionally (`CalcContext.cs:73`) |
| Array-path single/multi decision | `FunctionDefinition.cs:161-172` |
| Argument conversion | inside each `Adapt*` closure |

Note what is **not** there: arity is never re-checked at evaluation time, so every adapter reads
`args[0]`, `args[1]` … unguarded and relies on the parser having enforced `minParams`.

### The registration inventory

Recounted from source, resolving nested parentheses rather than by line:

| File | Registrations | via `Adapt*` | `AllowRange` ≠ `None` | `Only`/`Except` | File lines |
|---|---:|---:|---:|---:|---:|
| `MathTrig.cs` | 72 | 67 | 15 | 5 | 1336 |
| `Distributions.cs` | 62 | 54 | 8 | 4 | 826 |
| `Engineering.cs` | 54 | 54 | 0 | 0 | 645 |
| `Statistical.cs` | 53 | 21 | 47 | 20 | 1079 |
| `Text.cs` | 36 | 31 | 8 | 5 | 1425 |
| `Financial.cs` | 33 | 26 | 6 | 6 | 1013 |
| `DateAndTime.cs` | 25 | 23 | 4 | 4 | 1113 |
| `Regression.cs` | 23 | 1 | 23 | 4 | 966 |
| `DynamicArray.cs` | 18 | 0 | 18 | 0 | 818 |
| `Information.cs` | 15 | 14 | 5 | 0 | 194 |
| `Lookup.cs` | 11 | 8 | 9 | 4 | 799 |
| `Logical.cs` | 9 | 5 | 3 | 1 | 153 |
| **Total** | **411** | **304** | **146** | **53** | |

> **Corrections to the commissioned evidence.** `grep -c 'RegisterFunction('` returns **442**: one
> declaration (`FunctionRegistry.cs:61`) plus 441 lines. **30 of those 441 are commented-out
> placeholders** (`//ce.RegisterFunction("LOOKUP", , Lookup);` and the like). The live figure is
> **411**, not 441. Likewise the line-based counts 307/139 become **304** via `Adapt*` and **146**
> with a non-default `AllowRange` once multi-line calls are joined and commented lines dropped.
> These are the numbers every gate in this spec uses.

Arity distribution, which shapes the design:

| `maxParams` | Registrations |
|---|---:|
| 0–7 | 351 |
| 30 | 1 |
| 254 / 255 | 37 |
| `int.MaxValue` | 22 |

**351 of 411 are fixed-arity with at most 7 parameters.** The 60 variadic ones are the awkward case
and get an explicit escape hatch.

## Non-goals

- **No observable behaviour change for any function.** Same results, same errors, same intersection.
- **Not adding or removing functions.** That is spec 07 (waves A–F done; optional wave A2 open).
- **Not touching `FunctionDefinition`'s array path** — `CallAsArray`, `NormalizeArguments`,
  `EvaluateArrayElements`, `EvaluateSingleElement`, `GetScalarArgsMaxSize`
  (`FunctionDefinition.cs:52-124`). That is spec 30. 32 changes only what
  `IsParameterSingleValue` reads, not what the array path does with the answer.
- **Not implementing LET/LAMBDA.** Spec 08.
- **Not deleting `FunctionFlags`**, despite five of its six members being write-only. Removing it is
  a separate, smaller change with its own argument, and folding it into a 411-site sweep would make
  a mechanical diff into a semantic one.
- **Not writing a source generator.** See below.
- **No public API change.** `FunctionRegistry`, `FunctionDefinition`, `SignatureAdapter` and
  `AllowRange` are all `internal`. `XLFunctionLibrary` is public and its behaviour is unchanged.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24).

- `SignatureAdapter.cs` — 1,358 lines; 61 overloads; 55 distinct `Func<>` shapes; converters
  `:1219-1355`
- `SignatureAdapter.cs:18-21` — the "found by counting its parameters" discovery rule
- `SignatureAdapter.cs:28` — "We can hopefully generate them at a later date"
- `FunctionRegistry.cs:7-20` — `internal enum AllowRange { None, All, Except, Only }`
- `FunctionRegistry.cs:61-65` — `RegisterFunction(string functionName, int minParams, int maxParams,
  CalcEngineFunction fn, FunctionFlags flags, AllowRange allowRanges = AllowRange.None,
  params int[] markedParams)`
- `FunctionRegistry.cs:32`, `:37` — the two `TryGetFunc` overloads
- `FunctionDefinition.cs:16`, `:22` — `_allowRanges`, `_markedParams`
- `FunctionDefinition.cs:26-27` — the only validation of `markedParams`
- `FunctionDefinition.cs:126-141` — `IntersectArguments`, the per-argument runtime loop
- `FunctionDefinition.cs:161-172` — `IsParameterSingleValue`, the array path's reader
- `FunctionDefinition.cs:54` — the only read of `_flags`, and it reads only `ReturnsArray`
- `XLCalcEngine.cs:752` — `internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue> arg);`
- `XLCalcEngine.cs:78` — `internal FunctionRegistry Functions { get; }`
- `CalculationVisitor.cs:75-95` — the hot path into `CallFunction`/`CallAsArray`
- `CalcContext.cs:73` — `public static bool UseImplicitIntersection => true;`
- `XLFunctionLibrary.cs:77-117` — `TryInvoke`, the public path; arity check at `:87`
- `XLibur/Properties/AssemblyInfo.cs:3` — `InternalsVisibleTo("XLibur.Tests")`, so tests can reach
  all of the above
- `XLibur.Tests/Excel/CalcEngine/` — 48 test files; none references `SignatureAdapter`
- `XLibur.Benchmarks/FormulaEvaluationBenchmarks.cs` — three probes, all over `SUM`
- **No source generator exists in the repo.** The only Roslyn package is
  `Microsoft.CodeAnalysis.PublicApiAnalyzers` (`XLibur/XLibur.csproj:48`).

## File structure

```
XLibur/Excel/CalcEngine/ArgSpec.cs                      new — ArgKind, ArgSpec, AdaptedFunction
XLibur/Excel/CalcEngine/FunctionRegistry.cs             modified — two new overloads (task 1);
                                                                   old overload + AllowRange deleted (task 8)
XLibur/Excel/CalcEngine/FunctionDefinition.cs           modified — _args replaces _allowRanges/_markedParams
XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs   modified — every overload returns AdaptedFunction;
                                                                   AdaptLastOptional/AdaptLastTwoOptional folded away
XLibur/Excel/CalcEngine/Functions/Logical.cs            modified — 9 registrations
XLibur/Excel/CalcEngine/Functions/Information.cs        modified — 15
XLibur/Excel/CalcEngine/Functions/Lookup.cs             modified — 11
XLibur/Excel/CalcEngine/Functions/Statistical.cs        modified — 53
XLibur/Excel/CalcEngine/Functions/Regression.cs         modified — 23
XLibur/Excel/CalcEngine/Functions/DateAndTime.cs        modified — 25
XLibur/Excel/CalcEngine/Functions/Financial.cs          modified — 33
XLibur/Excel/CalcEngine/Functions/Text.cs               modified — 36
XLibur/Excel/CalcEngine/Functions/DynamicArray.cs       modified — 18
XLibur/Excel/CalcEngine/Functions/Engineering.cs        modified — 54
XLibur/Excel/CalcEngine/Functions/MathTrig.cs           modified — 72
XLibur/Excel/CalcEngine/Functions/Distributions.cs      modified — 62

XLibur.Tests/Excel/CalcEngine/FunctionSignatureTableTests.cs   new — task 0, the gate for 3–8
XLibur.Benchmarks/FunctionAdapterBenchmarks.cs                 new — task 2, the decision gate
```

## The design

### The adapter emits the spec, because it is the only thing that knows it

The mistake to avoid is writing an `ArgSpec[]` literal at each of the 411 call sites. That would be
a *fourth* encoding — better shaped than `AllowRange` + `markedParams`, and still capable of
disagreeing with the delegate.

Instead, each `Adapt*` overload returns its delegate **and** the spec its own generic signature
implies:

```csharp
/// <summary>
/// A function ready to register: the closure that invokes it, and the argument shape the adapter
/// that built the closure knows to be true.
/// </summary>
/// <remarks>
/// The spec is not written at the registration site. It is produced by the <c>Adapt*</c> overload,
/// which is the one place that already knows how many arguments it reads, which of them it converts
/// to a scalar, and which it hands through as a raw <see cref="AnyValue"/>. Deriving it there is
/// what makes arity and range-ness impossible to state wrongly at a call site.
/// </remarks>
internal readonly struct AdaptedFunction
{
    internal AdaptedFunction(CalcEngineFunction fn, ArgSpec[] args)
    {
        Fn = fn;
        Args = args;
    }

    internal CalcEngineFunction Fn { get; }

    internal ArgSpec[] Args { get; }
}
```

```csharp
/// <summary>What one function parameter accepts.</summary>
/// <remarks>
/// Two members on purpose. This is the only distinction anything downstream acts on:
/// <see cref="FunctionDefinition.IntersectArguments"/> intersects everything that is not
/// <see cref="Range"/>, and <see cref="FunctionDefinition.IsParameterSingleValue"/> asks the same
/// question the other way round. The *type* an argument is converted to — number, text, logical —
/// stays where it already is, inside the adapter closure, and is not duplicated here. A third
/// member describing conversion would be documentation that nothing reads, which is the defect
/// this file exists to remove.
/// </remarks>
internal enum ArgKind : byte
{
    /// <summary>A single value. A range in this position is implicitly intersected first.</summary>
    Value,

    /// <summary>A range or array, handed to the function untouched. Never implicitly intersected.</summary>
    Range,
}
```

```csharp
/// <summary>One parameter's shape: what it accepts, and whether it may be omitted.</summary>
internal readonly struct ArgSpec
{
    private ArgSpec(ArgKind kind, bool optional, bool repeating)
    {
        Kind = kind;
        Optional = optional;
        Repeating = repeating;
    }

    internal ArgKind Kind { get; }

    /// <summary>The argument may be omitted. Every parameter after an optional one is optional.</summary>
    internal bool Optional { get; }

    /// <summary>
    /// This spec applies to its own position and to every argument after it. Set on the last spec of
    /// a variadic function such as <c>SUM</c> or <c>CONCAT</c>, so a 255-parameter function needs one
    /// spec rather than 255.
    /// </summary>
    internal bool Repeating { get; }

    internal static ArgSpec Value(bool optional = false) => new(ArgKind.Value, optional, repeating: false);

    internal static ArgSpec Range(bool optional = false) => new(ArgKind.Range, optional, repeating: false);

    internal static ArgSpec ValueRest() => new(ArgKind.Value, optional: true, repeating: true);

    internal static ArgSpec RangeRest() => new(ArgKind.Range, optional: true, repeating: true);
}
```

`NETWORKDAYS` becomes:

```csharp
ce.RegisterFunction("NETWORKDAYS", AdaptLastOptional(NetWorkDays), FunctionFlags.Range); // Returns the number of whole workdays between two dates
```

The `2, 3` and the `AllowRange.Only, 2` are gone. Both are now produced by the overload at
`SignatureAdapter.cs:593`, once, for every function that shares that shape:

```csharp
private static readonly ArgSpec[] ScalarScalarRangeOpt =
    [ArgSpec.Value(), ArgSpec.Value(), ArgSpec.Range(optional: true)];

public static AdaptedFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, AnyValue, ScalarValue> f)
{
    return new AdaptedFunction((ctx, args) => { /* body unchanged */ }, ScalarScalarRangeOpt);
}
```

The array is `static readonly`, built once at type-init, shared by every function of that shape. 55
arrays replace 411 `(min, max, AllowRange, int[])` tuples.

### Arity and intersection derive from the spec

```csharp
    /// <summary>Minimum arguments: every parameter up to the first optional one.</summary>
    private static int MinFrom(ArgSpec[] args)
    {
        var min = 0;
        while (min < args.Length && !args[min].Optional)
            ++min;

        return min;
    }

    /// <summary>Maximum arguments, or <see cref="int.MaxValue"/> when the last spec repeats.</summary>
    private static int MaxFrom(ArgSpec[] args, int? declaredMax)
        => args.Length > 0 && args[^1].Repeating
            ? declaredMax ?? int.MaxValue
            : args.Length;
```

`declaredMax` is the escape hatch for the 37 registrations that cap a variadic at 254/255 and the
one at 30. They pass it explicitly:

```csharp
ce.RegisterFunction("CONCAT", Adapt(Concat), FunctionFlags.Future | FunctionFlags.Range, maxParams: 255);
```

`IntersectArguments` becomes strictly less work than it is today:

```csharp
    private void IntersectArguments(CalcContext ctx, Span<AnyValue> args)
    {
        var last = _args.Length - 1;
        for (var i = 0; i < args.Length; ++i)
        {
            // The last spec repeats for a variadic function, so an index past the end reads it.
            if (_args[i < last ? i : last].Kind != ArgKind.Range)
                args[i] = args[i].ImplicitIntersection(ctx);
        }
    }
```

Against today's body, which for every argument evaluates a four-arm `switch` on `_allowRanges` and
then calls `_markedParams.Contains(i)` — `Enumerable.Contains` over an `IReadOnlyCollection<int>`,
which lands on `ICollection<int>.Contains` through an interface dispatch, on an array, O(n). The new
form is one array index and one byte compare.

`IsParameterSingleValue` collapses the same way:

```csharp
    private bool IsParameterSingleValue(int paramIndex)
    {
        var last = _args.Length - 1;
        return _args[paramIndex < last ? paramIndex : last].Kind != ArgKind.Range;
    }
```

### The migration keeps one runtime representation from task 1 onward

The old `RegisterFunction` overload is not deleted until task 8. It is instead **translated**:

```csharp
    /// <summary>
    /// The pre-spec-32 registration form. Kept while the 411 call sites are converted; it builds the
    /// same <see cref="ArgSpec"/> array the new form takes, so both paths run identical code from
    /// task 1 onward and the sweep is a call-site rewrite with no behavioural surface.
    /// </summary>
    internal static ArgSpec[] Translate(int minParams, int maxParams, AllowRange allowRanges,
        IReadOnlyCollection<int> markedParams)
    {
        var variadic = maxParams > 8;
        var count = variadic ? Math.Max(minParams, 1) : maxParams;
        var specs = new ArgSpec[count];
        for (var i = 0; i < count; ++i)
        {
            var isRange = allowRanges switch
            {
                AllowRange.None => false,
                AllowRange.All => true,
                AllowRange.Only => markedParams.Contains(i),
                AllowRange.Except => !markedParams.Contains(i),
                _ => throw new InvalidOperationException($"Unexpected value {allowRanges}")
            };
            var repeating = variadic && i == count - 1;
            var optional = repeating || i >= minParams;
            specs[i] = (isRange, repeating) switch
            {
                (true, true) => ArgSpec.RangeRest(),
                (true, false) => ArgSpec.Range(optional),
                (false, true) => ArgSpec.ValueRest(),
                (false, false) => ArgSpec.Value(optional),
            };
        }

        return specs;
    }
```

The translation is lossless for `AllowRange.None`, `All` and, for a variadic function, `Only`/
`Except` — provided no variadic function marks an index past the repeat point. **That is a premise,
and task 1 step 3 is what would disprove it**: it asserts round-trip equality
(`Translate(...)` → derived `min`/`max`/per-index range-ness == the originals) for all 411
registrations. If any registration fails to round-trip, record which and stop; the design needs a
per-index spec for that one function rather than a repeating tail.

### What this design does **not** do: collapse 61 overloads into 1

The brief asked for "one argument-spec module … with `Adapt` reduced to a small set of typed invoke
shims or source-generated". Assessed honestly, **the 61 overloads cannot become one, and reducing
them below 55 requires generation.** The reason is structural, not stylistic:

Every function body is strongly typed. `NetWorkDays(CalcContext, ScalarValue, ScalarValue,
AnyValue)` is a four-parameter method with four distinct parameter types. C# offers exactly three
ways to call it after converting arguments at runtime:

1. **A shim per shape** — a method whose generic `Func<>` parameter has that arity and those types.
   That is what the 61 overloads are.
2. **Reflection / `DynamicInvoke`** — boxes every argument, allocates a `object[]` per call, and
   puts a reflection dispatch on the hot path of every formula. Not viable.
3. **Rewrite all 411 bodies** to a uniform `Func<CalcContext, in ConvertedArgs, ScalarValue>` and
   pull typed values by index (`args.Number(0)`, `args.Text(1)`). That is 411 method signatures and
   ~5,000 lines of function body, it discards the compiler's type checking of each function's own
   arguments, and it is a far larger and riskier change than the one this spec proposes.

So the reduction available without generation is: **fold optionality into `ArgSpec` and keep one
adapter per distinct `Func<>` shape — 61 → 55.** `AdaptLastOptional` (14) and
`AdaptLastTwoOptional` (7) stop being separate families; the six duplicated shapes tabulated above
merge; the nine remaining after that keep their own names because they differ in *collection
strategy*, not shape (`AdaptIfs`'s `ToCriteria` pairs, `AdaptIndex`'s `List<int>` tail,
`AdaptMultinomial`'s `List<IEnumerable<ScalarValue>>`).

**Six methods. That is the whole reduction in the adapter file, and it should be stated up front
rather than discovered at task 8.** The value of this spec is in the 411 call sites, not the 61
overloads.

### Source generation vs hand-written shims — recommendation: no generator

The file asks for one (`:28`). Do not write it, for four reasons:

1. **It would be the repo's first.** There is no `IIncrementalGenerator` anywhere in the tree; the
   only Roslyn dependency is `PublicApiAnalyzers` (`XLibur/XLibur.csproj:48`). A generator means a
   new `netstandard2.0` project, an analyzer `PackageReference` with `OutputItemType="Analyzer"`,
   packaging so consumers do not get it, and a build step that must work under net8.0, net9.0 and
   net10.0 SDKs and in CI.
2. **It would generate 55 shims to save writing 55 shims.** The 55 exist already and change roughly
   once per new function shape — spec 07's six waves added functions, but the shape count grew far
   more slowly than the function count (411 functions, 55 shapes).
3. **`TreatWarningsAsErrors=true` makes generator output a liability.** Any analyzer warning in
   generated code is a build break, and generated code is the least pleasant place to suppress one.
4. **The spec that would justify a generator is not this one.** If option 3 above is ever taken —
   uniform function bodies — generation becomes the obvious way to emit the per-function argument
   pulls. `ArgSpec` is the input such a generator would read. This spec leaves that door open and
   walks through none of it.

Record this decision in `SignatureAdapter.cs` where the "hopefully generate" comment is, so the next
reader does not re-open it from scratch.

### Runtime cost — the central risk, and the decision rule

The commissioned brief framed this as "compile-time shape resolution moved to a runtime loop, on the
hot path of every formula evaluation". Half of that is already false today: **`IntersectArguments`
is already a per-argument runtime loop**, and this design replaces it with a cheaper one. What
changes on the hot path:

| | Today | After |
|---|---|---|
| Per argument, intersection | `switch` on `_allowRanges` + `IReadOnlyCollection<int>.Contains(i)` (interface dispatch, O(n) scan) | one array index + one `byte` compare |
| Per call, conversion | inside the `Adapt*` closure, unchanged | inside the `Adapt*` closure, **unchanged** |
| Per call, delegate invocations | 2 (closure, then `f`) | 2 (identical) |
| Per registration, allocation | `int[]` markedParams for 53 of 411 | one shared `static readonly ArgSpec[]` per shape, 55 total |
| Structs embedded by value | none | none |

**Expected result: neutral to slightly faster.** No struct is embedded in another struct, no delegate
call is added, no closure changes.

That expectation is a premise, and **spec 21 is the reason this spec does not act on a premise about
the JIT.** Spec 21 predicted that removing an interface dispatch on the slice enumerator would make
enumeration faster. Its Results:

| Variant | Mean | Allocated |
|---|---:|---:|
| baseline — `Enumerator` is a class, held via `IEnumerator<Point>` | 5,053,270 ns | 88 B |
| struct, still boxed via `IEnumerator<Point>` | 5,068,989 ns | 88 B |
| struct **embedded by value** via a wrapper | 8,463,473 ns | 0 B |
| the same, `[MethodImpl(AggressiveInlining)]` on the wrapper | 7,928,567 ns | 0 B |
| the same, wrapper removed entirely | 8,108,683 ns | 0 B |

**+60% on the primary instrument. Task 1 was implemented, measured, and reverted; task 2 was
declined without being written.** Spec 21 also records that the interface dispatch it set out to
remove had already been devirtualised by dynamic PGO, so it was never the cost. Its closing lesson,
quoted because it governs this spec too: *"A criterion written before the measurement it governs
should not be allowed to override the measurement."*

**Ground rule for the measurement**, from spec 19 and the repo's benchmark discipline:

- **BenchmarkDotNet only.** This machine has ~40% run-to-run timing variance on identical code, and
  hand-rolled median-of-9 harnesses have produced apparent 30% wins that BenchmarkDotNet showed to
  be 0%. Hand-rolled probes are for *locating* cost and for allocation numbers, which are exact.
- **Three runs per arm, compared on the median of the per-run means.** A delta counts only if it
  exceeds the run-to-run spread of the baseline arm itself.
- **`git stash push -- XLibur/`** to A/B, so the benchmark project is byte-identical across arms.

**Decision rule — task 2 applies it, and task 2 may end this spec:**

| Median delta on any adapter-driven probe | Action |
|---|---|
| Faster, or within ±2% | Proceed to task 3. |
| +2% to +5% | Stop. Investigate the mechanism, record it, and re-measure once. Proceed only if a second measurement lands under +2%. Two consecutive readings above +2% end the spec. |
| Above +5% | **Revert task 1 and stop.** Record the numbers in a Results section and close the spec as declined-on-measurement. |

The threshold is deliberately tighter than spec 21's would have been, and the reason is stated
plainly: **this spec's benefit is defect prevention with zero measured runtime value.** A change
worth nothing on the clock cannot buy anything on the clock. Spec 21's mis-stated rule would have
kept a 60% regression to remove an 88-byte constant; the correction it published is exactly the
principle applied here.

**The existing formula benchmarks cannot measure this.** All three probes in
`FormulaEvaluationBenchmarks` evaluate `SUM`, and `SUM` is registered raw at `MathTrig.cs:109` —
`ce.RegisterFunction("SUM", 1, int.MaxValue, Sum, FunctionFlags.Range, AllowRange.All)` — with no
adapter and with `AllowRange.All`, so it exercises neither an `Adapt*` closure nor a non-trivial
intersect loop. Task 2 must add a probe. Spec 19's post-fix baselines for the existing three, kept
as a control: `UniqueSameSheet` 36.73 ms, `SharedSameSheet` 36.39 ms, `SharedCrossSheet` 32.56 ms,
8.85 MB each.

### Migration shape

411 call sites cannot land in one commit. The sweep is per-file, behind an unchanged old overload:

- **Task 1** adds `ArgSpec`, `AdaptedFunction`, `Translate`, and two new `RegisterFunction`
  overloads. The old overload stays and routes through `Translate`. From this point there is one
  runtime representation and both call forms produce it.
- **Tasks 3–7** convert one file or small group at a time. Each is independently revertible; each is
  gated by task 0's table.
- **Task 8** deletes the old overload, the `AllowRange` enum, `markedParams`, and folds optionality
  into `ArgSpec` (61 → 55 overloads). It can only compile once the last file is converted, which is
  the mechanical proof that it was.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable enabled; new code must be
  null-annotated.
- Branch per spec; never commit to main. Commit prefixes `refactor:` / `fix:` / `test:` / `perf:`.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's `sed -i`
  rewrites the file as LF, turning a one-line change into a whole-file diff. Use the Edit/Write
  tools. Verify with `git diff --numstat` — a file whose changed-line count approaches its total
  line count was rewritten, not edited. **This matters more here than in any other spec in the
  series**: tasks 3–7 make hundreds of one-line edits across twelve files, several of them over
  1,000 lines, and a stray `sed -i` would make the diff unreviewable.
- Test filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option; exit 8 = zero
  tests matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` for iteration; run without it before opening the PR.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Tests use TUnit: `await Assert.That(actual).IsEqualTo(expected)`. **Assertions are awaitable and a
  missing `await` silently passes** — CS4014 is an error here, not a warning. `[Test]`,
  `[Arguments(...)]`, `[MethodDataSource(...)]`. The suite runs serially.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 0 | Characterization: pin arity and per-argument intersection for all 411 registered functions | M | New test green on the **unmodified** tree; 411 rows |
| 1 | `ArgSpec`, `AdaptedFunction`, `Translate`, two new `RegisterFunction` overloads — old path routed through the new representation | M | Task 0 green; round-trip assertion green for all 411 |
| 2 | Benchmark the two representations; apply the decision rule | S | Medians recorded. **The spec may stop here.** |
| 3 | Convert `Logical.cs`, `Information.cs`, `Lookup.cs` — 35 registrations | S | Task 0 green |
| 4 | Convert `Statistical.cs`, `Regression.cs` — 76 registrations, 70 with `AllowRange` | M | Task 0 green |
| 5 | Convert `DateAndTime.cs`, `Financial.cs`, `Text.cs` — 94 | M | Task 0 green |
| 6 | Convert `DynamicArray.cs`, `Engineering.cs` — 72 | M | Task 0 green |
| 7 | Convert `MathTrig.cs`, `Distributions.cs` — 134 | L | Task 0 green |
| 8 | Delete the old overload, `AllowRange`, `markedParams`; fold optionality (61 → 55) | M | grep gates; suite green |
| 9 | Re-benchmark; write Results | S | Within task 2's threshold |

Task 4 is deliberately early despite its size: `Statistical.cs` and `Regression.cs` carry 70 of the
146 non-default `AllowRange` values and 24 of the 53 marked sets. If the design cannot express
something, it will fail there, while task 0's table is still fresh and only 35 sites are behind it.

---

### Task 0 — Characterization: what every registered function does today

**There is currently no direct test surface for any of this.** Nothing asserts a function's arity,
nothing asserts which of its arguments are implicitly intersected, and no test file mentions
`SignatureAdapter`. This task builds the gate before anything moves. It is mandatory.

**Files:**
- Create: `XLibur.Tests/Excel/CalcEngine/FunctionSignatureTableTests.cs`

**Interfaces:**
- Produces: `Every_function_accepts_exactly_its_declared_arity` and
  `Every_argument_intersects_or_does_not_as_recorded` — the gate for tasks 1 and 3–8.

- [ ] **Step 1: Pin arity for every registered name**

`XLFunctionLibrary` is the public path and enumerates the registry (`XLFunctionLibraryTests.cs`
already uses it this way). Arity is enforced at `XLFunctionLibrary.cs:87`, so a call with the wrong
count returns `XLError.IncompatibleValue` rather than throwing.

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Every registered function's arity and implicit-intersection behaviour, captured before spec 32
/// moves either. Nothing in the suite asserted these before: no test file referenced
/// <c>SignatureAdapter</c>, and the 411 registrations state arity and range-ness in three places
/// that were never checked against each other. This is what holds the 411-site sweep to the
/// behaviour it started from.
/// </summary>
public class FunctionSignatureTableTests
{
    /// <summary>
    /// Arity is checked before the function body runs, so an out-of-range count reports
    /// <c>#VALUE!</c> and an in-range count does something else — a result, another error, or a
    /// demand for a worksheet. The distinction is what the boundary is read from.
    /// </summary>
    private static bool AcceptsArity(XLFunctionLibrary library, string name, int count)
    {
        var args = new XLCellValue[count];
        for (var i = 0; i < count; ++i)
            args[i] = 1.0;

        try
        {
            if (!library.TryInvoke(name, args.AsSpan(), out var result))
                return false;

            return !(result.IsError && result.GetError() == XLError.IncompatibleValue
                     && IsArityRejection(library, name, count));
        }
        catch (XLNoWorksheetContextException)
        {
            // Dispatched, then asked for a grid. The arity was accepted.
            return true;
        }
    }
```

`IsArityRejection` is the one piece that cannot be inferred from the outside: a function whose body
legitimately returns `#VALUE!` for numeric `1.0` arguments is indistinguishable from an arity
rejection. Read the boundary from the registry instead, which the test assembly can reach through
`InternalsVisibleTo` (`XLibur/Properties/AssemblyInfo.cs:3`):

```csharp
    /// <summary>
    /// The declared arity window, read from the registry rather than inferred. Spec 32's whole
    /// subject is that this window is stated separately from the delegate that consumes it; this
    /// test records what it currently says so the sweep cannot change it silently.
    /// </summary>
    private static (int Min, int Max) DeclaredArity(string name)
    {
        using var wb = new XLWorkbook();
        var engine = ((XLWorkbook)wb).CalcEngine;
        return engine.Functions.TryGetFunc(name, out var min, out var max)
            ? (min, max)
            : throw new InvalidOperationException($"No function named '{name}'.");
    }
```

If `XLWorkbook.CalcEngine` is not reachable under that name, find the accessor the existing calc
tests use — `XLCalcEngine.Functions` is `internal` at `XLCalcEngine.cs:78` — and use that. **Do not
add a new public or internal accessor for this test**; if none exists, construct the engine the way
`XLFunctionLibrary`'s constructor does.

```csharp
    /// <summary>
    /// One row per registered function: its name and its declared arity window. 411 rows at
    /// 1b41cadd. The count itself is an assertion — a registration silently dropped by the sweep
    /// shows up here before it shows up as a missing function.
    /// </summary>
    public static IEnumerable<Func<string>> FunctionNames()
    {
        using var wb = new XLWorkbook();
        foreach (var name in new XLFunctionLibrary().Names.OrderBy(n => n, StringComparer.Ordinal))
        {
            var captured = name;
            yield return () => captured;
        }
    }

    [Test]
    public async Task The_registry_holds_the_expected_number_of_functions()
    {
        // 411 live RegisterFunction call sites at 1b41cadd. If this number moves, either a function
        // was added (spec 07) or the sweep dropped one. Both need a deliberate edit here.
        await Assert.That(new XLFunctionLibrary().Names.Count()).IsEqualTo(411);
    }

    [Test]
    [MethodDataSource(nameof(FunctionNames))]
    public async Task Every_function_accepts_exactly_its_declared_arity(string name)
    {
        var library = new XLFunctionLibrary();
        var (min, max) = DeclaredArity(name);

        if (min > 0)
            await Assert.That(AcceptsArity(library, name, min - 1)).IsFalse();

        await Assert.That(AcceptsArity(library, name, min)).IsTrue();

        if (max < 32)
            await Assert.That(AcceptsArity(library, name, max + 1)).IsFalse();
    }
}
```

`max < 32` skips the 60 variadic registrations, whose upper bound is 30, 254, 255 or
`int.MaxValue`. Record that skip in the test's summary; it is a real gap in the gate, not an
oversight.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FunctionSignatureTableTests/*"`
Expected: PASS, 411 rows on the arity test.

**If `The_registry_holds_the_expected_number_of_functions` reports a number other than 411**, the
count in this spec is wrong and every gate that cites it must be updated to the measured value
before proceeding. Record the number and the discrepancy.

- [ ] **Step 2: Pin implicit intersection per argument**

Arity is the easy half. Intersection is the half `markedParams` controls, and it needs a grid.

```csharp
    /// <summary>
    /// Implicit intersection is observable: a formula in row 3 passing the multi-row range
    /// <c>A1:A5</c> to a parameter that does *not* accept ranges is narrowed to <c>A3</c> before the
    /// function sees it, while a parameter that does accept ranges receives all five cells. The two
    /// give different answers for any function that reads more than the first cell.
    /// </summary>
    /// <remarks>
    /// One row per (function, argument index) pair for the 53 registrations that carry
    /// <c>AllowRange.Only</c> or <c>AllowRange.Except</c> — the only ones where a wrong index is
    /// possible. The other 358 are all-or-nothing and are covered by their existing function tests.
    /// </remarks>
    [Test]
    [MethodDataSource(nameof(MarkedRegistrations))]
    public async Task Every_argument_intersects_or_does_not_as_recorded(IntersectionCase test)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        for (var row = 1; row <= 5; ++row)
            ws.Cell(row, 1).Value = row;

        ws.Cell(3, 3).FormulaA1 = test.Formula;
        var actual = ws.Cell(3, 3).Value;

        await Assert.That(actual.ToString()).IsEqualTo(test.Expected);
    }
```

`MarkedRegistrations` is a hand-built table, one entry per marked argument. It cannot be generated
from the registry, because the registry does not expose `_markedParams` and this spec must not widen
it to do so — widening it would put a second reader on the field the sweep deletes.

**Build the table from `NETWORKDAYS` outward.** Start with the 53 marked registrations, which are
listed by:

Run: `grep -rn "AllowRange.Only\|AllowRange.Except" XLibur/Excel/CalcEngine/Functions --include=*.cs`
Expected: 53 lines across 9 files — `Statistical.cs` 20, `Financial.cs` 6, `Text.cs` 5,
`MathTrig.cs` 5, `Regression.cs` 4, `Lookup.cs` 4, `Distributions.cs` 4, `DateAndTime.cs` 4,
`Logical.cs` 1. `Engineering.cs`, `DynamicArray.cs` and `Information.cs` carry none.

Record `Expected` as **whatever the current code produces**, including errors. This is a
characterization test, not a correctness test. **Do not fix anything you find here.** If a value
looks wrong, record it, keep it, and report it — a wrong value pinned is still a gate, and finding
one would upgrade this spec from prevention to repair.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FunctionSignatureTableTests/*"`
Expected: PASS.

- [ ] **Step 3: Verify the gate bites**

In `FunctionDefinition.cs:134`, temporarily change `AllowRange.Only => !_markedParams.Contains(i)`
to `AllowRange.Only => true`. Re-run.
Expected: FAIL on multiple `Every_argument_intersects_or_does_not_as_recorded` rows. Restore the
line.

**If it does not fail**, the intersection table is not discriminating and tasks 3–7 have no gate.
Rebuild the table with ranges that span the formula's row *and* column before continuing.

- [ ] **Step 4: Commit**

```bash
git add XLibur.Tests/Excel/CalcEngine/FunctionSignatureTableTests.cs
git commit -m 'test(calc): pin arity and implicit intersection for all 411 functions (spec 32 task 0)'
```

---

### Task 1 — `ArgSpec` and the driving loop, inert alongside the existing path

**Files:**
- Create: `XLibur/Excel/CalcEngine/ArgSpec.cs`
- Modify: `XLibur/Excel/CalcEngine/FunctionRegistry.cs`
- Modify: `XLibur/Excel/CalcEngine/FunctionDefinition.cs`

**Interfaces:**
- Produces: `ArgKind`, `ArgSpec`, `AdaptedFunction`, `FunctionDefinition.Translate`,
  `FunctionRegistry.RegisterFunction(string, AdaptedFunction, FunctionFlags, int?)` and
  `FunctionRegistry.RegisterFunction(string, CalcEngineFunction, FunctionFlags, params ArgSpec[])`.

- [ ] **Step 1: Add `ArgSpec.cs`**

The `ArgKind`, `ArgSpec` and `AdaptedFunction` declarations from "The design" above, verbatim,
including their doc comments. All three are `internal`; nullable is enabled and none of them has a
reference-typed field except `AdaptedFunction.Fn` and `AdaptedFunction.Args`, both non-nullable.

- [ ] **Step 2: Replace `FunctionDefinition`'s two fields with one**

Delete `_allowRanges` (`:16`) and `_markedParams` (`:18-22`). Add:

```csharp
    /// <summary>
    /// One entry per parameter. The last entry repeats for every argument past its own index, so a
    /// variadic function such as <c>SUM</c> carries one spec rather than <c>int.MaxValue</c> of them.
    /// </summary>
    private readonly ArgSpec[] _args;
```

Replace `IntersectArguments` (`:126-141`) and `IsParameterSingleValue` (`:161-172`) with the two
bodies from "The design". Keep `CallAsArray`'s guard at `:54` working — it reads
`_allowRanges == AllowRange.All`, which becomes:

```csharp
        // Every parameter accepts a range: the function handles the whole array itself.
        if (_flags.HasFlag(FunctionFlags.ReturnsArray) && AllArgsAreRanges)
            return _function(ctx, args);
```

```csharp
    private bool AllArgsAreRanges
    {
        get
        {
            foreach (var spec in _args)
            {
                if (spec.Kind != ArgKind.Range)
                    return false;
            }

            return true;
        }
    }
```

Compute it once in the constructor into a `readonly bool` rather than per call — `CallAsArray` is on
the array-formula path, not the scalar hot path, but a loop per call for a value fixed at
registration is free to avoid.

**Do not touch `NormalizeArguments`, `EvaluateArrayElements`, `EvaluateSingleElement` or
`GetScalarArgsMaxSize`.** They are spec 30's. This task changes only what `IsParameterSingleValue`
reads, not what those four do with its answer.

- [ ] **Step 3: Add `Translate` and assert it round-trips for all 411**

Add `Translate` from "The design" as `private static` on `FunctionDefinition`, and keep the existing
constructor, now delegating:

```csharp
    /// <summary>The pre-spec-32 form. Deleted in task 8, once every call site has been converted.</summary>
    public FunctionDefinition(int minParams, int maxParams, CalcEngineFunction function, FunctionFlags flags,
        AllowRange allowRanges, IReadOnlyCollection<int> markedParams)
        : this(function, flags, Translate(minParams, maxParams, allowRanges, markedParams), maxParams)
    {
        if (allowRanges == AllowRange.None && markedParams.Count > 0)
            throw new ArgumentException("Marked params must be empty when AllowRange is None.", nameof(markedParams));
    }
```

Then prove the translation is lossless, in `FunctionSignatureTableTests`:

```csharp
    /// <summary>
    /// Spec 32 task 1's premise: (minParams, maxParams, AllowRange, markedParams) maps onto an
    /// ArgSpec[] without loss for every registration in the tree. If a single one fails to round
    /// trip, that function needs per-index specs rather than a repeating tail, and the design must
    /// change before the sweep starts.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(FunctionNames))]
    public async Task Every_registration_round_trips_through_ArgSpec(string name)
    {
        // DeclaredArity reads FunctionRegistry.TryGetFunc(name, out min, out max), which now returns
        // FunctionDefinition.MinParams/MaxParams — and those are derived from the ArgSpec[] that
        // Translate produced from the old tuple. Equality with the literal in the registration is
        // therefore the round-trip assertion.
        var (min, max) = DeclaredArity(name);
        var (declaredMin, declaredMax) = ExpectedArity[name];

        await Assert.That(min).IsEqualTo(declaredMin);
        await Assert.That(max).IsEqualTo(declaredMax);
    }
```

`ExpectedArity` is a literal table of the 411 `(name, minParams, maxParams)` triples as they read in
the registration files at `1b41cadd`. Generate it once, from the tree, with the same
nested-paren-aware scan that produced this spec's counts, and check the generated file in. It is the
one place in this spec where the old encoding is written down deliberately — as data that the new
one must reproduce.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FunctionSignatureTableTests/*"`
Expected: PASS, 411 rows.

**If any registration fails to round-trip, stop and record which.** The likely candidate is a
variadic function marking an index at or past the repeat point. That would disprove the design's
`Repeating` shortcut and is a real result worth writing down rather than working around.

- [ ] **Step 4: Add the two new `RegisterFunction` overloads**

```csharp
    /// <summary>
    /// Register a function whose argument shape comes from the adapter that built it.
    /// </summary>
    /// <param name="functionName">Name of the function in formulas.</param>
    /// <param name="adapted">The closure and the argument spec, both produced by one <c>Adapt*</c> call.</param>
    /// <param name="flags">Flags that indicate some additional info about a function.</param>
    /// <param name="maxParams">
    /// Upper bound for a variadic function that caps below <see cref="int.MaxValue"/> — the 254/255
    /// text and math functions. Omit for fixed-arity functions; the spec array length is the bound.
    /// </param>
    public void RegisterFunction(string functionName, AdaptedFunction adapted, FunctionFlags flags,
        int? maxParams = null)
    {
        _func.Add(functionName, new FunctionDefinition(adapted.Fn, flags, adapted.Args, maxParams));
    }

    /// <summary>
    /// Register a function that takes <see cref="Span{T}"/> of arguments directly, with no adapter.
    /// Its shape has no delegate signature to be derived from, so it is stated here.
    /// </summary>
    public void RegisterFunction(string functionName, CalcEngineFunction fn, FunctionFlags flags,
        params ArgSpec[] args)
    {
        _func.Add(functionName, new FunctionDefinition(fn, flags, args, maxParams: null));
    }
```

The second overload serves the 107 registrations that pass a raw `CalcEngineFunction`
(411 − 304). It is the one place an `ArgSpec[]` literal is written at a call site, and it is
unavoidable: those functions have no typed signature to derive from.

- [ ] **Step 5: Build and run the whole suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS. Nothing has been converted yet; every registration still goes through the old
overload, and every one of them now runs on the new representation.

Run: `git diff --numstat`
Expected: three files, changed-line counts well below their totals. A whole-file rewrite here means
`sed -i` was used somewhere.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/CalcEngine/ArgSpec.cs XLibur/Excel/CalcEngine/FunctionDefinition.cs XLibur/Excel/CalcEngine/FunctionRegistry.cs XLibur.Tests/Excel/CalcEngine/FunctionSignatureTableTests.cs
git commit -m 'refactor(calc): drive arity and intersection from one ArgSpec array (spec 32 task 1)'
```

---

### Task 2 — Benchmark, and decide whether this spec continues

**This task can end the spec, and that would be a real result.**

**Files:**
- Create: `XLibur.Benchmarks/FunctionAdapterBenchmarks.cs`

- [ ] **Step 1: Add a probe that actually reaches the adapter path**

The three existing probes in `FormulaEvaluationBenchmarks` all evaluate `SUM`, which is registered
raw with `AllowRange.All` (`MathTrig.cs:109`) and therefore exercises neither an `Adapt*` closure nor
a discriminating intersect loop. Add one that does, over three shapes:

```csharp
/// <summary>
/// The adapter path, which no existing benchmark reaches. <see cref="FormulaEvaluationBenchmarks"/>
/// evaluates SUM, registered without an adapter and with AllowRange.All, so it measures neither the
/// per-argument conversion closure nor an intersect loop that has to discriminate between arguments.
/// Spec 32 changes both, so both need an instrument.
/// </summary>
/// <remarks>
/// Three shapes, chosen to cover the cases spec 32 touches:
/// <list type="bullet">
/// <item>ROUND — 2 fixed numeric arguments, AllowRange.None, the commonest shape (115 of 411
/// registrations have maxParams 1 and 99 have 2).</item>
/// <item>NETWORKDAYS — 3 arguments, last optional, AllowRange.Only 2; the worked example in the
/// spec and the only shape where a marked index does any work.</item>
/// <item>SUM — unchanged control. If this moves, the measurement is noise, not the change.</item>
/// </list>
/// </remarks>
[MemoryDiagnoser]
public class FunctionAdapterBenchmarks
{
    private const int RowCount = 20_000;

    private XLWorkbook _round = null!;
    private XLWorkbook _networkDays = null!;
    private XLWorkbook _sum = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _round = Build(row => $"ROUND(D{row}, 2)");
        _networkDays = Build(row => $"NETWORKDAYS(D{row}, D{row + 1}, A1:A5)");
        _sum = Build(row => $"SUM(D{row}:D{row + 19})");
    }

    [Benchmark]
    public void Round() => _round.RecalculateAllFormulas();

    [Benchmark]
    public void NetworkDays() => _networkDays.RecalculateAllFormulas();

    [Benchmark(Baseline = true)]
    public void Sum() => _sum.RecalculateAllFormulas();
}
```

Copy `Build`, the font bootstrap call and the `[GlobalCleanup]` shape from
`FormulaEvaluationBenchmarks.cs` — `Setup` at `:61`, `Cleanup` at `:87` — rather than inventing
them. That file's remarks already document why the row count is 20,000 and why the summed-cell
count is held constant across variants (spec 19 area 5 task 5.1).

Run: `dotnet build XLibur.Benchmarks/XLibur.Benchmarks.csproj -c Release -v q`

- [ ] **Step 2: Measure the merge-base, three runs**

```bash
git stash push -- XLibur/
```

Run, three times:
```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -f net10.0 -- --filter "*FunctionAdapterBenchmarks*"
```

```bash
git stash pop
```

Record all three means per probe. This machine has ~40% run-to-run variance, so the **spread of the
baseline arm is the noise floor for the comparison** — a delta smaller than it means nothing.

- [ ] **Step 3: Measure the branch, three runs**

Same command, same fixture, task 1 applied.

- [ ] **Step 4: Apply the decision rule**

Compare medians of the per-run means. From "Runtime cost" above:

| Median delta on `Round` or `NetworkDays` | Action |
|---|---|
| Faster, or within ±2% | Proceed to task 3. |
| +2% to +5% | Investigate, record the mechanism, re-measure once. Two readings above +2% end the spec. |
| Above +5% | **Revert task 1. Close the spec as declined-on-measurement.** |

`Sum` is the control. If `Sum` moves by more than the baseline spread, the measurement is invalid —
nothing in task 1 touches a raw `AllowRange.All` registration — and the whole comparison must be
re-run before it means anything.

**The expected result is neutral or slightly faster**, because the intersect loop got cheaper, no
struct is embedded by value, and no delegate call was added. **That expectation is a premise, and
this task is what disproves it if it is wrong.** Spec 21 held an equally mechanical expectation
about interface dispatch and enumeration and was wrong by 60%.

- [ ] **Step 5: Write the numbers into this file, whichever way they go**

Add a Results section with the three-by-three table and the verdict. If the spec stops here, that
section is the deliverable and tasks 3–9 are marked declined.

```bash
git add XLibur.Benchmarks/FunctionAdapterBenchmarks.cs docs/specs/32-function-argument-spec.md
git commit -m 'perf(calc): measure the ArgSpec argument loop against the AllowRange path (spec 32 task 2)'
```

---

### Task 3 — Convert `Logical.cs`, `Information.cs`, `Lookup.cs` (35 registrations)

The pattern-setting sweep: three small files, 35 registrations, 12 with a non-default `AllowRange`
and 5 with marked indices. Small enough to review line by line, wide enough to hit every shape
family.

**Files:**
- Modify: `XLibur/Excel/CalcEngine/Functions/Logical.cs` (9)
- Modify: `XLibur/Excel/CalcEngine/Functions/Information.cs` (15)
- Modify: `XLibur/Excel/CalcEngine/Functions/Lookup.cs` (11)
- Modify: `XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs` — the overloads these files use
  change return type to `AdaptedFunction`

- [ ] **Step 1: Convert one overload, then its call sites**

For each `Adapt*` overload reached by these three files, change the return type and add the shared
spec array. Worked, for `Logical.cs:13`'s `IF`:

```csharp
    private static readonly ArgSpec[] ValueRangeRangeOpt =
        [ArgSpec.Value(), ArgSpec.Range(), ArgSpec.Range(optional: true)];

    public static AdaptedFunction AdaptLastOptional(Func<CalcContext, ScalarValue, AnyValue, AnyValue, AnyValue> f)
    {
        return new AdaptedFunction((ctx, args) => { /* body unchanged */ }, ValueRangeRangeOpt);
    }
```

Then:

```csharp
// was: ce.RegisterFunction("IF", 2, 3, AdaptLastOptional(If, false), FunctionFlags.Range, AllowRange.Only, 1, 2);
ce.RegisterFunction("IF", AdaptLastOptional(If, false), FunctionFlags.Range);
```

Check the spec array against the old `AllowRange.Only, 1, 2` **by hand, once per shape**: parameters
1 and 2 were marked as range-accepting, and the delegate declares them `AnyValue`. They agree. Where
they do not agree, **the delegate is right and the old tail was wrong** — record it, because that is
a defect this spec would have found.

- [ ] **Step 2: Name the raw registrations explicitly**

`Lookup.cs:25`, for instance:

```csharp
// was: ce.RegisterFunction("INDIRECT", 1, 2, Indirect, FunctionFlags.Range | FunctionFlags.Volatile);
ce.RegisterFunction("INDIRECT", Indirect, FunctionFlags.Range | FunctionFlags.Volatile,
    ArgSpec.Value(), ArgSpec.Value(optional: true));
```

The old form defaulted `allowRanges` to `AllowRange.None`, so both parameters were intersected. The
new form says so.

- [ ] **Step 3: Run the gate**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FunctionSignatureTableTests/*"`
Expected: PASS, all 411 rows.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/LogicalTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/InformationTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/LookupTests/*"`
Expected: PASS.

Run: `git diff --numstat`
Expected: changed-line counts near the registration counts (9, 15, 11) plus the adapter edits, not
near the file totals (153, 194, 799).

- [ ] **Step 4: Commit**

```bash
git add XLibur/Excel/CalcEngine/Functions/Logical.cs XLibur/Excel/CalcEngine/Functions/Information.cs XLibur/Excel/CalcEngine/Functions/Lookup.cs XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs
git commit -m 'refactor(calc): derive arity and range-ness for Logical, Information, Lookup (spec 32 task 3)'
```

---

### Tasks 4–7 — The remaining nine files

Identical shape to task 3. Each converts its files, runs task 0's table plus the file's own test
class, checks `git diff --numstat`, and commits.

| Task | Files | Registrations | `AllowRange` ≠ `None` | `Only`/`Except` | Own test classes |
|---|---|---:|---:|---:|---|
| 4 | `Statistical.cs`, `Regression.cs` | 76 | 70 | 24 | `StatisticalTests`, `StatisticalRankPercentileTests`, `HypothesisTestTests`, `RegressionTests` |
| 5 | `DateAndTime.cs`, `Financial.cs`, `Text.cs` | 94 | 18 | 15 | `DateAndTimeTests`, `WorkdayIntlTests`, `Financial*Tests` (5 classes), `TextTests`, `ModernTextTests` |
| 6 | `DynamicArray.cs`, `Engineering.cs` | 72 | 18 | 0 | `DynamicArrayFunctionTests`, `EngineeringTests`, `EngineeringComplexTests`, `EngineeringConvertTests` |
| 7 | `MathTrig.cs`, `Distributions.cs` | 134 | 23 | 9 | `MathTrigTests`, `XLMathTests`, `DistributionTests` |

**Task 4 is the one that would break the design.** `Statistical.cs` carries 47 non-default
`AllowRange` values and 20 marked sets in 53 registrations — more intersection configuration than
any other file, and it registers only 21 of its 53 through an adapter, so 32 of them will need
explicit `ArgSpec[]` literals. If `ArgSpec` cannot express something, it shows up there. **Do task 4
before tasks 5–7**, so a design change costs 35 converted sites rather than 277.

Task 6's `Engineering.cs` is the opposite: 54 registrations, all through `Adapt`, zero
`AllowRange`. It is the most mechanical of the nine and the best candidate for a large single
commit.

Commit messages follow task 3's:

```bash
git commit -m 'refactor(calc): derive arity and range-ness for Statistical, Regression (spec 32 task 4)'
```

---

### Task 8 — Delete the old form, fold optionality, 61 → 55

**Files:**
- Modify: `XLibur/Excel/CalcEngine/FunctionRegistry.cs` — delete `AllowRange` (`:7-20`) and the old
  `RegisterFunction` (`:61-65`)
- Modify: `XLibur/Excel/CalcEngine/FunctionDefinition.cs` — delete the old constructor and
  `Translate`
- Modify: `XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs` — fold optionality

- [ ] **Step 1: Delete the old overload and `AllowRange`**

If anything still calls the old form, the build fails and names it. That failure is the proof the
sweep is complete, so **do not stub the overload to get past it.**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Expected: clean.

Run: `grep -rn "AllowRange" XLibur --include=*.cs`
Expected: no output.

Run: `grep -rn "markedParams" XLibur --include=*.cs`
Expected: no output.

- [ ] **Step 2: Fold `AdaptLastOptional` and `AdaptLastTwoOptional` into `Adapt`**

Optionality now lives in the `ArgSpec`, so the three families collapse to one adapter per distinct
`Func<>` shape. The six duplicated shapes tabulated in "Why this spec exists" merge pairwise; the
other fifteen `AdaptLastOptional`/`AdaptLastTwoOptional` overloads become `Adapt` overloads of
shapes that had none.

`AdaptLastOptional(Func<CalcContext, double, double, ScalarValue> f, double lastDefault)`
(`SignatureAdapter.cs:457`) and `Adapt(Func<CalcContext, double, double, ScalarValue> f)` (`:71`)
have the same shape and differ only in the default. Keep one, with the default as an optional
parameter:

```csharp
    private static readonly ArgSpec[] TwoValues = [ArgSpec.Value(), ArgSpec.Value()];
    private static readonly ArgSpec[] ValueValueOpt = [ArgSpec.Value(), ArgSpec.Value(optional: true)];

    /// <summary>
    /// Two numeric arguments. Passing <paramref name="lastDefault"/> makes the second optional, which
    /// is the only thing the separate <c>AdaptLastOptional</c> family used to express.
    /// </summary>
    public static AdaptedFunction Adapt(Func<CalcContext, double, double, ScalarValue> f,
        double? lastDefault = null)
    {
        var fallback = lastDefault ?? 0d;
        return new AdaptedFunction(
            (ctx, args) => { /* body: arg1 = args.Length > 1 ? ToNumber(args[1]) : fallback */ },
            lastDefault is null ? TwoValues : ValueValueOpt);
    }
```

**Do this shape by shape, running the suite between each.** Sixteen shapes merge or move; the nine
one-offs (`AdaptIfs` ×2, `AdaptIndex`, `AdaptMatch`, `AdaptSeriesSum`, `AdaptNumberValue`,
`AdaptSubstitute`, `AdaptMultinomial`, `AdaptCoerced`) keep their names — they differ in *collection
strategy*, not shape, and merging them into `Adapt` would restore the "find it by counting
parameters" problem this spec exists to remove.

- [ ] **Step 3: Replace the `S4136` suppression comment**

The comment at `SignatureAdapter.cs:18-21` documents the discovery rule this spec removes. Replace
it, and record the generator decision next to the "hopefully generate" comment at `:28`:

```csharp
// S4136 wants the overloads grouped by name. They are ordered by arity instead, and each carries the
// ArgSpec array its own signature implies (spec 32), so a registration no longer states arity or
// range-ness separately and cannot state either one wrongly. One adapter per distinct Func<> shape:
// 55 of them, down from 61 families-plus-shapes before optionality moved into ArgSpec.
#pragma warning disable S4136
```

```csharp
    // Not source-generated, deliberately: spec 32 priced it. A generator would emit these 55 shims —
    // exactly what is here — and would be this repo's first, needing a netstandard2.0 analyzer
    // project, packaging that keeps it off consumers, and generated code that satisfies
    // TreatWarningsAsErrors. Six methods of duplication is not worth that. If the function bodies
    // are ever made uniform, generation becomes worthwhile and ArgSpec is what it should read.
```

- [ ] **Step 4: Verify the interface narrowed**

Run: `grep -c "public static AdaptedFunction" XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs`
Expected: `55`.

Run: `grep -c "public static CalcEngineFunction" XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs`
Expected: `0`.

Run: `grep -c "AdaptLastOptional\|AdaptLastTwoOptional" XLibur --include=*.cs -r`
Expected: `0`.

- [ ] **Step 5: Full suite, both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/CalcEngine/FunctionRegistry.cs XLibur/Excel/CalcEngine/FunctionDefinition.cs XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs
git commit -m 'refactor(calc): delete AllowRange and markedParams; fold optionality into ArgSpec (spec 32 task 8)'
```

---

### Task 9 — Confirm no regression, and write Results

- [ ] **Step 1: Re-run task 2's benchmark against the finished branch**

Run, three times:
```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -f net10.0 -- --filter "*FunctionAdapterBenchmarks*" "*FormulaEvaluation*"
```

Expected: `Round` and `NetworkDays` within task 2's threshold of the merge-base; `Sum`,
`UniqueSameSheet`, `SharedSameSheet` and `SharedCrossSheet` unchanged within the baseline spread.

Task 8 removed the `Translate` indirection that tasks 3–7 ran through, so this reading should be at
least as good as task 2's. If it is worse, the folding in task 8 step 2 is the suspect — it is the
only step that changed a closure body.

- [ ] **Step 2: Write the Results section**

Following spec 19's format: what was measured, on what commit, on what hardware, three runs per arm,
medians, and every premise this spec stated that turned out to be false. Specifically record:

- whether the translation round-tripped for all 411 (task 1 step 3);
- any registration whose old `AllowRange`/`markedParams` tail disagreed with its delegate
  (tasks 3–7 step 1) — **this is the finding that would upgrade the spec from prevention to
  repair**;
- the final overload count against the predicted 55;
- the benchmark deltas.

- [ ] **Step 3: Commit**

```bash
git add docs/specs/32-function-argument-spec.md
git commit -m 'docs(specs): record the argument-spec numbers and findings for spec 32'
```

---

## Acceptance criteria

1. `grep -rn "AllowRange" XLibur --include=*.cs` returns nothing.
2. `grep -rn "markedParams" XLibur --include=*.cs` returns nothing.
3. `grep -c "public static CalcEngineFunction" XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs`
   returns `0`; `grep -c "public static AdaptedFunction"` on the same file returns `55`.
4. `grep -rn "AdaptLastOptional\|AdaptLastTwoOptional" XLibur --include=*.cs` returns nothing.
5. No `RegisterFunction` call site states `minParams` or `maxParams` except the 60 variadic ones that
   cap below `int.MaxValue`. Gate:
   `grep -rnE 'RegisterFunction\("[^"]+", [0-9]+, ' XLibur --include=*.cs` returns nothing.
6. `FunctionSignatureTableTests` passes with 411 rows on each of its `[MethodDataSource]` tests, and
   `The_registry_holds_the_expected_number_of_functions` still reads 411.
7. No assertion in `FunctionSignatureTableTests` was weakened between task 0 and task 9. Gate:
   `git diff <task-0-commit>..HEAD -- XLibur.Tests/Excel/CalcEngine/FunctionSignatureTableTests.cs`
   shows additions only, except for values changed under criterion 10.
8. Full suite green on net8.0 and net10.0.
9. `Round` and `NetworkDays` within task 2's stated threshold of the merge-base, on medians of three
   runs each; `Sum` and the three `FormulaEvaluationBenchmarks` probes within the baseline spread.
10. Every registration whose old tail disagreed with its delegate is listed in Results, with the
    formula that distinguishes the two behaviours. **Zero such registrations is an acceptable and
    expected outcome** — task 0's audit already found none — and must be stated as a result rather
    than left implicit.
11. No public API change. `git diff` on `XLibur/PublicAPI.Shipped.txt` and
    `XLibur/PublicAPI.Unshipped.txt` is empty.
12. `git diff --numstat` shows no file whose changed-line count approaches its total — proof no
    `sed -i` touched a tracked file.
13. If task 2 ended the spec, criteria 1–5 and 8–11 do not apply; the Results section and the reverted
    task 1 are the deliverable, and this file's Status says declined-on-measurement.

## Conflicts

Read `docs/specs/README.md` before starting. Against the specs listed there:

- **Spec 30 (array application gets an interface, and the defect it hides gets fixed) — the live
  conflict, and it goes first.** 30 owns `FunctionDefinition.cs`'s array path: `CallAsArray`
  (`:52`), `NormalizeArguments` (`:66`), `EvaluateArrayElements` (`:92`), `EvaluateSingleElement`
  (`:106`) and `GetScalarArgsMaxSize` (`:143`). Three of those five call `IsParameterSingleValue`
  (`:161`), which reads the `_allowRanges`/`_markedParams` fields 32 replaces.

  **Two corrections to 30's Conflicts section, which was written the same day as this one.** First,
  30 states that "no file is shared" between the two specs. That is not right: **32 modifies
  `FunctionDefinition.cs`** — its task 1 replaces the two fields and rewrites `IntersectArguments`
  and `IsParameterSingleValue`. It is a file collision as well as a semantic one, which strengthens
  30's own recommendation rather than weakening it. Second, 30 corroborates "441 call sites" and
  "9 one-off named variants". Recounted here with a nested-paren-aware scan: **411 live call sites**
  (441 grep lines minus 30 commented-out placeholders), and the nine one-off overloads sit under
  **eight** names, because `AdaptIfs` has two (`SignatureAdapter.cs:652`, `:667`). The 61-overload
  headline both specs cite is correct.

  **Recommended order: 30 first**, which is also 30's own recommendation. It is three days, it is a
  correctness fix for a **confirmed** defect — `EvaluateSingleElement` builds `itemArg` and then
  calls `_function(ctx, args)` at `FunctionDefinition.cs:117`, so the per-element argument array it
  just built is discarded and every element of an array formula is evaluated against the same
  broadcast arguments — and 32 is a 411-site sweep that should land on corrected semantics rather
  than have them changed underneath it.

  If 32 lands first, 30 must rebase onto `ArgSpec`, which changes what its fix reads without changing
  what it means — recoverable, but pointless work.

- **Spec 07 (formula function coverage, waves A–F done).** 07 added a large share of the 411
  registrations; its optional remaining **wave A2** (day-count-basis financial functions) would add
  more, in `Financial.cs`, which is task 5's territory. Either **32 precedes A2**, or **A2 uses the
  new registration form** and skips being converted. Do not run them concurrently: A2 adding
  registrations mid-sweep breaks task 0's `411` count assertion and makes every gate ambiguous.

- **Spec 04 (demand-driven formula evaluation).** Overlap is real but shallow. 04 owns the
  evaluation stack and `CalcContext`; 32 touches `CalculationVisitor.cs` **not at all** and
  `CalcContext.cs` not at all. The one shared surface is `FunctionDefinition.CallFunction`
  (`FunctionDefinition.cs:41`), which 04 may re-enter differently and 32 changes the body of. They
  can run concurrently with a merge conflict in one method. **32 should not start while 04 is
  mid-flight**, because 04 is marked single-owner correctness-critical in the README and a
  411-site sweep landing under it is exactly the kind of churn that makes a correctness bug hard to
  attribute.

- **Spec 08 (LET / LAMBDA).** LAMBDA introduces user-defined functions whose argument shape is not
  known at registration time, so it needs either a runtime `ArgSpec[]` or a bypass. **32 makes 08
  easier, not harder** — a per-instance `ArgSpec[]` is the natural representation for a LAMBDA's
  parameter list, where `AllowRange` + `markedParams` is not. Sequence 32 before 08 if both are
  scheduled; there is no file conflict today because 08 is unwritten.

- **Specs 22, 23, 24, 25** (the 2026-08-23 architecture round) are all in the IO and style layers and
  touch nothing in `XLibur/Excel/CalcEngine/`. No conflict.

- **Specs 13 and 12** touch `FunctionRegistry.Names` (`FunctionRegistry.cs:29`) and
  `XLCalcEngine.Functions` (`XLCalcEngine.cs:78`). 32 changes neither member's signature or
  behaviour, and `XLFunctionLibrary`'s public contract is unchanged. No conflict.
