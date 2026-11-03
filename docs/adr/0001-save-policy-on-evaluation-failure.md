# Save writes no cached value for an expected failure, and throws on a defect

On save, a dirty formula is evaluated so its cached value can be written. If that evaluation fails with a circular reference, an unsupported feature or a refused formula, save leaves the cell with no cached value, and Excel recalculates it when the file is opened. Any other failure is a defect, and the save throws. Until this decision, save swallowed every failure. That silently undid the rule from #459 that a defect always reaches the caller. Throwing on every failure was rejected too, because a save would then fail over a feature XLibur simply does not have. Decided in the round-4 architecture review (spec 56).

## Addendum: the public API draws the same line

`TryEvaluate` on a cell, a worksheet and a workbook returns a circular reference, an unsupported feature and a refused formula as an `XLFormulaIssue`, and throws every other failure, a defect above all. It is a second reader of this decision, not a second decision: save and `TryEvaluate` ask one function, `EvaluationFailure.IsExpected`, so a failure that becomes expected is added in one place and both follow. An Excel error value such as `#DIV/0!` is a result, not a failure, and is returned as a value.

## Consequences

A gap that XLibur knows it has must be raised as an unsupported feature, not as `NotImplementedException`; otherwise saves start to fail. The first such gap is the argument converter's array branch. It is reclassified as unsupported until spec 30 implements array application.
