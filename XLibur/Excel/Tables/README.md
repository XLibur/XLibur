# Invariants

`table/tableColumns/tableColumn[@totalsRowLabel]` attribute must match the cell value stored in the sheet at the totals row label cell. If the text in a cell doesn't match, Excel considers the workbook to be corrupt and tries to repair it. A totals cell that the sheet leaves out opens blank, under the label.

The load does not repair a file that breaks this: it keeps the totals cell as the file has it, because Excel shows the cell, and keeps the label the table part gives. The save writes both back (#742). Setting a label, a value in the totals cell, or a totals function through the API keeps the two in step.
