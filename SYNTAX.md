# xlsxfill Placeholder Syntax

Generate an xlsx by merging a template xlsx with data.

## Principles

**Do not bring features Excel already has into the template language.**

| Goal | Means |
|---|---|
| Arithmetic | Excel formulas `=B1*C1` |
| Formatting (dates, currency, thousands separators) | Cell number formats |
| Conditionals | `=IF()`, conditional formatting, filter on the data side |
| Growing aggregation ranges | Mixed references `=SUM(D$1:D1)` |

**The data's type decides everything. No conversion.**

Type specifiers are assertions (validation), not conversions.

**Writable by an ordinary Excel user.**

Everything is expressed with strings that can be typed through the Excel UI.

---

## Notation

### Repetition

```
#{rDept}              start position of the band
#{rDept+1}            start position of the next band
```

`+1` is not an end marker. It indicates **where the second iteration starts**. The distance between the start position and the `+1` position is the band's height (or width). Just like Excel's autofill: placing the first and second determines the period.

The prefix determines the direction.

| Prefix | Meaning | What it looks at |
|---|---|---|
| `r` | Vertical (row direction) | The **row** of the cell it is placed in |
| `c` | Horizontal (column direction) | The **column** of the cell it is placed in |

`r` only looks at the row, so the column is arbitrary. `c` only looks at the column, so the row is arbitrary.

Anything may follow `r` / `c`, including Japanese.

```
#{r}        #{rDept}        #{rStaff}        #{r部署}
#{c}        #{cEmp}
```

Sheet repetition has no declaration (→ Examples > Sheets).

A start marker may share a cell with other content.

```
#{rDept}${departments#rDept.name}
```

Vertical and horizontal start markers may be written in the same cell.

```
#{r}#{c}${temperatures#r#c}
```

Whitespace is allowed.

```
#{ rDept }        #{ rDept + 1 }
```

### Value references

```
${customer.name}                          scalar
${departments#rDept.name}                 current element of a band
${departments#rDept.staff#rStaff.name}    nesting
${temperatures#r#c}                       2D intersection
${items#2}                                fixed index
${#rStaff}                                index value
${sheets#s.name}                          sheet
```

After `#`, **a name means a variable, a number means a fixed index**.

Square brackets are not used, because `[` is not allowed in sheet names.

Whitespace is allowed.

```
${ customer . name }
```

### Index

| Form | Base |
|---|---|
| `${items#2}` (as an index) | 0-based |
| `${#rStaff}` (as a value) | 1-based |

### Type assertions

```
${qty:num}
${createdAt:date}
```

| | Excel side |
|---|---|
| `:num` | Number |
| `:str` | String |
| `:bool` | Boolean TRUE/FALSE |
| `:date` | Date |
| `:time` | Time |
| `:datetime` | Datetime |

All optional. If written, a mismatch with the data's type is an error.

### Links and images

```
${[a.title](a.url)}          link
${[](a.url)}                 label is the surrounding text in the cell
${![a.alt](a.data)}          image
${![](a.data)}               no alt
```

Just like Markdown, `!` distinguishes the two.

An Excel hyperlink is an attribute attached to the cell, so **the entire cell becomes the link**.

```
Details: ${[a.title](a.url)}     →  the whole "Details: {title}" is a link
Details${[](a.url)}              →  "Details" is a link
```

How the image fits (CSS `object-fit`).

```
${![a.alt](a.data)}           contain (default)
${![a.alt](a.data)contain}
${![a.alt](a.data)cover}
${![a.alt](a.data)fill}
```

The reference is the cell's size. There is no pixel specification. Adjust the size via column width and row height.

Two or more images may be written in the same cell (an Excel image is not cell content but a shape floating on the sheet, and multiple shapes can be anchored to the same cell; how they overlap depends on what you specify and their sizes). Images and links can coexist.

To attach a link with the same URL as the cell value, write it twice.

```
${a.url}${[](a.url)}
```

---

## Examples

### Line items (vertical)

|  | A | B | C | D |
|---|---|---|---|---|
| 1 | `#{r}${items#r.name}` | `${items#r.qty}` | `${items#r.price}` | `=B1*C1` |
| 2 | `#{r+1}` | | | |

### Temperature table (2D)

|  | A | B | C |
|---|---|---|---|
| 1 | | `#{c}${employees#c.name}` | `#{c+1}` |
| 2 | `#{r}${dates#r}` | `${temperatures#r#c}` | |
| 3 | `#{r+1}` | | |

### Departments → staff (nesting)

|  | A | B | C | D |
|---|---|---|---|---|
| 1 | `#{rDept}${departments#rDept.name}` | `Name` | `Salary` | `Joined` |
| 2 | `#{rStaff}No. ${#rStaff}` | `${departments#rDept.staff#rStaff.name}` | `${departments#rDept.staff#rStaff.salary}` | `${departments#rDept.staff#rStaff.joinedAt}` |
| 3 | `#{rStaff+1}` | | | |
| 4 | `${departments#rDept.name} total` | `${departments#rDept.total}` | | |
| 5 | `#{rDept+1}` | | | |

### Staff laid out horizontally per department (horizontal inside vertical)

|  | A | B | C |
|---|---|---|---|
| 1 | `#{rDept}${departments#rDept.name}` | | |
| 2 | | `#{c}${departments#rDept.staff#c.name}` | `#{c+1}` |
| 3 | `#{rDept+1}` | | |

Staff cells are laid out to the largest department's width; a department with fewer staff leaves the extra cells empty.

### Tasks growing vertically per employee (vertical inside horizontal)

|  | A | B |
|---|---|---|
| 1 | `#{c}${employees#c.name}` | `#{c+1}` |
| 2 | `#{r}${employees#c.tasks#r}` | |
| 3 | `#{r+1}` | |

Task cells are laid out to the largest employee's height; an employee with fewer tasks leaves the extra cells empty.

An inner band may iterate over a collection unrelated to the outer one (as long as it does not extend outside). In that case the same content appears for each iteration of the outer band.

### Sheets

Sheets have no declaration. If `#s` appears anywhere in a sheet, that sheet is repeated. `s` is a fixed name.

Written in the sheet name (tab):

```
${sheets#s.name}
```

Written in a cell:

```
${sheets#s.customer}
```

---

## Where placeholders can be written

`${...}` can be written anywhere a string can be typed through the Excel UI.

- Cells
- Sheet names
- Headers/footers
- Text of shapes and text boxes
- Chart titles and axis labels
- Cell comments/notes
- Hyperlink tooltips
- Data validation input messages and error messages
- Document properties

Places other than cells are mere string containers, so type assertions, links, and images **can only be used in cells**.

`#{...}` (repetition) is also **cells only**.

---

## Behavior during expansion

### Bands

Expansion is whole-row / whole-column duplication — the same operations as inserting or deleting rows/columns in Excel.

A band nested in the same direction as its parent expands per block, with each block's own count. A band nested in the crossing direction shares its rows/columns with every block of the parent, so it expands once, to the largest count among the blocks; each block fills in its own values and shorter blocks leave the remaining cells empty.

| | |
|---|---|
| Collections of different lengths in the same band | min |
| 0 items | The whole band disappears (including header rows inside it). For a crossing-direction inner band this applies only when every block has 0 items; a single empty block just leaves its cells empty |
| The row/column where `#{r+1}` is placed | Removed (writing content there is meaningless) |
| Formulas inside a band | Updated to stay consistent |
| Tables, defined names, conditional formatting, print areas, filter ranges | Follow the expansion |
| Column widths when expanding horizontally | Duplicated |
| Shapes anchored to cells inside a band | Repeated |

A shape set to "don't move or size with cells" is not repeated (controllable with standard Excel functionality).

### Sheets

| | |
|---|---|
| Added sheets | Inserted consecutively at the original position |
| Tab name without `#s` | Made unique as `{original tab name} (n)` (n is a 1-based sequential number) |
| 0 items | The sheet disappears |

### Values

| | |
|---|---|
| null at the leaf (`customer.name` is null) | Blank cell |
| null inside a concatenation (`a` in `${a} ${b}`) | Empty string |
| `${first} ${last}` | Concatenated into a string |
| Boolean | Excel TRUE/FALSE |
| String containing newlines | In-cell newlines |
| url is null / empty string | No link |
| data is null / empty bytes | No image |

Nothing is done about numeric precision (Excel has 15 digits) or limits on the number of rows/columns after expansion.

---

## Errors

Errors are reported in one of two ways.

### Reported in the output (processing continues)

A message replaces the offending construct only; everything else — surrounding text, other valid placeholders, the rest of the sheet — is processed as usual. The rule is uniform over every place placeholders can be written. Errors involving multiple constructs (marker pairs, overlapping bands, etc.) put the message on each of them. A band whose declaration is invalid is treated as undeclared and does not expand.

Syntax errors use `#SYNTAX! <construct>: <reason>`. The list is written for `r`; for `c`, swap rows and columns:

- `${}` `#{}` (empty)
- `#{foo}` (name not starting with `r` or `c`)
- `#{r+2}` (anything other than `+1`)
- `#{s}` `#{s+1}` (sheets have no repetition declaration)
- `#{r}` written without `#{r+1}`
- `#{rDept}` and `#{rDept+1}` in different columns
- `#{r}` and `#{r+1}` in the same row (band height 0)
- Overlapping bands in the same direction
- `#{rDept}` declared but `#rDept` never used anywhere
- `${x#rFoo}` written but no `#{rFoo}` exists
- Using a band's name outside the band (headers/footers, sheet names, etc.)
- A merged cell straddling a band boundary
- Using type assertions, links, or images outside a cell
- Writing `#{...}` outside a cell
- Two or more links in the same cell
- A type assertion attached to a link or image (e.g. `${[a.title](a.url):str}`)

Data errors local to a value use `#DATA! <construct>: <reason>`:

- Mismatch with a type assertion
- Referencing a non-scalar value (dict / list)
- Fixed index out of range (`${items#5}` when there are only 3)
- tz-aware datetime (the caller makes it naive)
- null in the middle of a path (`customer.name` when `customer` is null)

### Raised (processing stops)

The input as a whole is unusable; substitution never starts.

- Root is not an object
- Property name contains `.` `#` `:` `(` `)` `[` `]` `!` `$` `{` `}`

If a sheet name violates Excel's constraints (31 characters, `\ / ? * [ ] :`, duplicates), Excel raises the error.

---

## Out of scope

| | |
|---|---|
| Escaping | Outputting a literal `${` is not allowed |
| Placeholders inside formulas | Excel's formula bar rejects `${...}`. Abandoned |
| Hidden sheets, sheet protection | Not considered |
| Fetching images from URLs | The caller converts them to bytes |
