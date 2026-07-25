# Spec 07 — Formula Function Coverage Expansion (257 → ~420 functions)

**Area:** Feature
**Effort:** L total, but **highly parallelizable** — 6 independent waves, each a self-contained PR assignable to a different agent/model.
**Dependencies:** None (waves are independent). LET/LAMBDA are explicitly **excluded** — see Spec 08 (they need parser/scoping work, not registration).
**Status:** Proposed

## Summary

The calc engine registers 257 functions vs Excel's ~500. The registry (`XLibur/Excel/CalcEngine/FunctionRegistry.cs`) is a single clean extension point, and recent PRs (#163–#168) established the pattern: implement in the matching `CalcEngine/Functions/*.cs` file, register with a signature adapter, test with Excel-verified values. This spec batches the missing functions into waves by category.

## Pattern to follow (read these PRs first)

- #165 (SMALL, RANK, PERCENTILE, QUARTILE, MODE), #166 (PV, NPV, IRR, RATE, NPER, PPMT), #163 (AVERAGEIF/S, MAXIFS, MINIFS), #168 (dynamic-array set).
- Implementation files: `XLibur/Excel/CalcEngine/Functions/{MathTrig,Statistical,Text,DateAndTime,Lookup,Information,Database,Engineering,Logical,Financial,DynamicArray}.cs`.
- Signature marshalling: `Functions/SignatureAdapter.cs` (1026 lines) — reuse existing adapters; add new arities only when unavoidable.
- Value model: `ScalarValue`/`AnyValue` discriminated unions; error propagation must match Excel (e.g. #NUM! vs #VALUE! distinctions matter and are tested).
- Tests: mirror existing test layout in `XLibur.Tests/Excel/CalcEngine/`; every function needs Excel-verified expected values (state in the test comment how values were verified), edge cases (empty args, wrong types, error propagation), and at least one in-worksheet evaluation test.

## Waves (each = one PR, independently assignable)

### Wave A — Financial (~28 functions) — file: `Financial.cs`
`XIRR, XNPV, MIRR, FV (verify exists), IPMT, SLN, SYD, DB, DDB, VDB, ISPMT, CUMIPMT, CUMPRINC, EFFECT, NOMINAL, RRI, PDURATION, DOLLARDE, DOLLARFR, TBILLEQ, TBILLPRICE, TBILLYIELD, DISC, INTRATE, RECEIVED, FVSCHEDULE`
Notes: XIRR/IRR-family use Newton–Raphson with bisection fallback — study the existing IRR/RATE implementations from #166 for the established root-finding approach. Day-count-basis functions (PRICE, YIELD, DURATION, MDURATION, ACCRINT, COUP*) are a **separate optional wave A2** — the 30/360 & actual/actual basis engine is its own chunk of work; don't block Wave A on it.

### Wave B — Statistical, modern dotted set (~48) — file: `Statistical.cs`
`NORM.DIST, NORM.INV, NORM.S.DIST, NORM.S.INV, LOGNORM.DIST, LOGNORM.INV, CHISQ.DIST, CHISQ.DIST.RT, CHISQ.INV, CHISQ.INV.RT, CHISQ.TEST, F.DIST, F.DIST.RT, F.INV, F.INV.RT, F.TEST, T.DIST, T.DIST.2T, T.DIST.RT, T.INV, T.INV.2T, T.TEST, Z.TEST, EXPON.DIST, POISSON.DIST, WEIBULL.DIST, GAMMA.DIST, GAMMA.INV, GAMMA, GAMMALN.PRECISE, BETA.DIST, BETA.INV, HYPGEOM.DIST, NEGBINOM.DIST, BINOM.DIST (verify), BINOM.INV, CONFIDENCE.NORM, CONFIDENCE.T, PERCENTILE.EXC, PERCENTILE.INC, QUARTILE.EXC, QUARTILE.INC, RANK.AVG, RANK.EQ, MODE.SNGL, MODE.MULT, PERCENTRANK.INC, PERCENTRANK.EXC` plus legacy aliases (`NORMDIST, NORMINV, NORMSDIST, NORMSINV, CHIDIST, CHIINV, CHITEST, FDIST, FINV, FTEST, TDIST, TINV, TTEST, ZTEST, EXPONDIST, POISSON, WEIBULL, GAMMADIST, GAMMAINV, BETADIST, BETAINV, HYPGEOMDIST, NEGBINOMDIST, CONFIDENCE, CRITBINOM`) which are thin aliases of the dotted implementations.
Notes: needs an internal special-functions helper (incomplete gamma/beta, erf) — implement once in a shared internal static class with accuracy tests against published values (≥ 1e-10 relative where Excel achieves it). Check what MODE/PERCENTILE/QUARTILE from #165 map to and alias accordingly.

### Wave C — Regression & descriptive statistics (~22) — file: `Statistical.cs`
`FREQUENCY, LINEST, LOGEST, TREND, GROWTH, FORECAST, FORECAST.LINEAR, CORREL, PEARSON, COVARIANCE.P, COVARIANCE.S, COVAR, SLOPE, INTERCEPT, RSQ, STEYX, SKEW, SKEW.P, KURT, PROB, TRIMMEAN, HARMEAN, GEOMEAN (verify), AVEDEV (verify), DEVSQ (verify)`
Notes: FREQUENCY/LINEST/TREND return arrays — they must integrate with the spill engine (see `Functions/DynamicArray.cs` and `docs/dynamic-array-spill-phase-b.md` for how array-returning functions spill).

### Wave D — Engineering (~35) — file: `Engineering.cs`
`COMPLEX, IMABS, IMAGINARY, IMARGUMENT, IMCONJUGATE, IMCOS, IMCOSH, IMCOT, IMCSC, IMCSCH, IMDIV, IMEXP, IMLN, IMLOG10, IMLOG2, IMPOWER, IMPRODUCT, IMREAL, IMSEC, IMSECH, IMSIN, IMSINH, IMSQRT, IMSUB, IMSUM, IMTAN, CONVERT, BESSELI, BESSELJ, BESSELK, BESSELY, ERF, ERF.PRECISE, ERFC, ERFC.PRECISE, DELTA, GESTEP, BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT`
Notes: complex numbers are strings in Excel ("3+4i") — parse/format helper first. CONVERT needs the full unit table from Excel docs (large but mechanical). BIT* validate 48-bit ranges.

### Wave E — Modern text + array shaping (~22) — files: `Text.cs`, `DynamicArray.cs`
Text: `TEXTSPLIT, TEXTBEFORE, TEXTAFTER, VALUETOTEXT, ARRAYTOTEXT, UNICHAR, UNICODE, DBCS (stub per locale rules or document omission), ENCODEURL`
Array shaping: `VSTACK, HSTACK, TOROW, TOCOL, WRAPROWS, WRAPCOLS, CHOOSEROWS, CHOOSECOLS, TAKE, DROP, EXPAND`
Notes: array-shaping functions are pure array→array transforms on the existing `Array` type and spill like the #168 set — this wave is the natural follow-on for whoever did #168. TEXTSPLIT returns 2-D arrays and spills.

### Wave F — Date/time + misc (~8) — files: `DateAndTime.cs`, `MathTrig.cs`
`NETWORKDAYS.INTL, WORKDAY.INTL (weekend-string/number parameter engine, shared helper), DAYS (verify), ISOWEEKNUM (exists — skip), AGGREGATE (option-driven dispatch to existing functions, ignore-error/hidden modes; hidden-row awareness can defer with a documented limitation), SUBTOTAL (verify completeness of 101–111 hidden-row variants)`

## Per-wave acceptance criteria

1. Every function registered, implemented, and covered by tests with Excel-verified values including error cases; `FunctionRegistry` count increases accordingly.
2. Excel parity on error types (#NUM!, #VALUE!, #DIV/0!, #N/A) for the tested cases.
3. Array-returning functions spill correctly (Waves C, E) — at least one spill test each.
4. No regressions: full `XLibur.Tests` green; build with `TreatWarningsAsErrors` clean.
5. PR description lists functions added and any deliberate deviations/limitations (e.g. DBCS, AGGREGATE hidden-row mode) — these also go in the changelog.

## Coordination notes for parallel execution

- Waves B and C both edit `Statistical.cs` — run them sequentially or have C branch from B.
- All waves touch `FunctionRegistry.cs` — keep registrations grouped by category in the file to minimize merge conflicts; rebase order A → D → E → F → B → C is conflict-minimal.
- Shared special-functions helper (Wave B) should land before or within B; C may reuse it.
