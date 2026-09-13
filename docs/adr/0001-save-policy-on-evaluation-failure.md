# Save writes no cached value for an expected failure, and throws on a defect

On save, a dirty formula is evaluated so its cached value can be written. If that evaluation fails with a circular reference, an unsupported feature or a refused formula, save leaves the cell with no cached value, and Excel recalculates it when the file is opened. Any other failure is a defect, and the save throws. Until this decision, save swallowed every failure. That silently undid the rule from #459 that a defect always reaches the caller. Throwing on every failure was rejected too, because a save would then fail over a feature XLibur simply does not have. Decided in the round-4 architecture review (spec 56).

## Consequences

A gap that XLibur knows it has must be raised as an unsupported feature, not as `NotImplementedException`; otherwise saves start to fail. The first such gap is the argument converter's array branch. It is reclassified as unsupported until spec 30 implements array application.
