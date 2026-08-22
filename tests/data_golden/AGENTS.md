# xlsxfill tests data_golden

A book is one area of the syntax, and inside it the cases form a matrix of
scenario against direction. A worksheet is one cell of it, and its tab name is
the coordinate, `<book>-<direction>-<scenario>`. Nesting is spelled outermost
first, so `rc` is a column band inside a row band.

A scenario fills every direction it can be stated in, and only those: `r` and
`c` where no nesting is needed, `rr` `rc` `cr` `cc` where two levels are.

A case about a value has no arrangement and so no matrix. It is one row of a
list sheet, its first cell naming it as a tab name would. A case whose subject
is the tab name, a header or footer, or a document property leaves the grid
empty and is written where the subject is.

A case owns the root key its data hangs from, sharing one only with cases that
differ from it in nothing but the workbook feature around them.

Sheet repetition cannot join a book: `#s` takes over the worksheet, leaving the
matrix nowhere to live. Each case is a **sheet-…** book of its own.
