`google-sheet workbook`
=======================

Manage local Excel (XLSX) workbooks

* [`google-sheet workbook:find`](#google-sheet-workbookfind)
* [`google-sheet workbook:inspect`](#google-sheet-workbookinspect)
* [`google-sheet workbook:names [ACTION]`](#google-sheet-workbooknames-action)
* [`google-sheet workbook:read`](#google-sheet-workbookread)
* [`google-sheet workbook:write [DATA]`](#google-sheet-workbookwrite-data)

## `google-sheet workbook:find`

Locate cells matching a condition in a local XLSX workbook and return their A1 coordinates. Same match semantics and result shape as data:find, so follow-up writes can target exact cells without reading the whole sheet. Pure read: no network, no auth, no writes.

```
USAGE
  $ google-sheet workbook:find -f <value> [-h] [--columns <value> | -x] [--sort <value>] [--filter <value>]
    [--output csv|json|yaml |  | [--csv | --no-truncate]] [--no-header | ] [-t <value>] [--range <value>] [--equals
    <value> | --contains <value> | --regex <value>] [--column <value> | --header <value>] [--ignoreCase] [--limit
    <value>] [--first] [--byRow] [-r]

FLAGS
  -f, --file=<value>            (required) Path to the local XLSX file to scan
  -h, --help                    Show CLI help.
  -r, --rawOutput               Get the raw output as a JSON string
  -t, --worksheetTitle=<value>  [env: WORKSHEET_TITLE] Title of the worksheet to scan (default: first sheet)
  -x, --extended                show extra columns
      --byRow                   Collapse matches to unique rows and include full row values
      --column=<value>          Restrict the scan to one column letter (e.g. "B")
      --columns=<value>         only show provided columns (comma-separated)
      --contains=<value>        Match cells whose rendered value contains this substring
      --csv                     output is csv format [alias: --output=csv]
      --equals=<value>          Match cells whose rendered value equals this string
      --filter=<value>          filter property by partial string matching, ex: name=foo
      --first                   Return only the first match (alias for --limit=1)
      --header=<value>          Restrict the scan to the column whose first-row header matches this value
      --[no-]ignoreCase         Case-insensitive matching
      --limit=<value>           [default: 100] Maximum matches to return (matchCount still reports the true total)
      --no-header               hide table header from output
      --no-truncate             do not truncate output to fit screen
      --output=<option>         output in a more machine friendly format
                                <options: csv|json|yaml>
      --range=<value>           The A1 range bounding the scan (default: used range of the sheet)
      --regex=<value>           Match cells whose rendered value matches this regular expression
      --sort=<value>            property to sort by (prepend '-' for descending)

DESCRIPTION
  Locate cells matching a condition in a local XLSX workbook and return their A1 coordinates. Same match semantics and
  result shape as data:find, so follow-up writes can target exact cells without reading the whole sheet. Pure read: no
  network, no auth, no writes.

EXAMPLES
  $ gsheet workbook:find --file=template.xlsx --equals="Tổng cộng" --first --rawOutput

  $ gsheet workbook:find --file=template.xlsx --worksheetTitle="Functional effort" --contains="MODULE" --byRow

  $ gsheet workbook:find --file=report.xlsx --regex="^ERR-" --range="Data!A1:K200"
```

_See code: [src/commands/workbook/find.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/workbook/find.ts)_

## `google-sheet workbook:inspect`

Inspect metadata, sheets, defined names, and capabilities of a local XLSX workbook

```
USAGE
  $ google-sheet workbook:inspect -f <value> [-h] [-r] [--includeFormulaCells]

FLAGS
  -f, --file=<value>         (required) Path to the local XLSX file to inspect
  -h, --help                 Show CLI help.
  -r, --rawOutput            Get the raw output as a JSON string
      --includeFormulaCells  Include the full formula-cell address list per sheet (bounded output otherwise)

DESCRIPTION
  Inspect metadata, sheets, defined names, and capabilities of a local XLSX workbook

EXAMPLES
  $ gsheet workbook:inspect --file=report.xlsx

  $ gsheet workbook:inspect --file=report.xlsx --rawOutput
```

_See code: [src/commands/workbook/inspect.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/workbook/inspect.ts)_

## `google-sheet workbook:names [ACTION]`

List, add or remove defined names (named ranges) in a local XLSX workbook. Run without an action to list every defined name, including the managed ranges the CLI records. add requires --name and --refersTo (a sheet-qualified A1 range); remove requires --name.

```
USAGE
  $ google-sheet workbook:names [ACTION] -f <value> [-h] [-o <value>] [--inPlace] [--discardUnsupported] [--name
    <value>] [--refersTo <value>] [--dryRun] [-r]

ARGUMENTS
  [ACTION]  (list|add|remove) [default: list] list (default), add or remove

FLAGS
  -f, --file=<value>        (required) Path to the local XLSX workbook
  -h, --help                Show CLI help.
  -o, --output=<value>      Destination path for the modified XLSX file (without it and without --inPlace the mutation
                            is refused unless --dryRun)
  -r, --rawOutput           Get the raw output as a JSON string
      --discardUnsupported  Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would be
                            dropped by the local engine
      --dryRun              Preview the mutation without applying it
      --inPlace             Modify the --file workbook in place (a .bak backup is written first)
      --name=<value>        Defined name to add or remove (add/remove only)
      --refersTo=<value>    Sheet-qualified A1 range the name refers to, e.g. "Data!$B$2:$B$4" (add only)

DESCRIPTION
  List, add or remove defined names (named ranges) in a local XLSX workbook. Run without an action to list every defined
  name, including the managed ranges the CLI records. add requires --name and --refersTo (a sheet-qualified A1 range);
  remove requires --name.

EXAMPLES
  $ gsheet workbook:names --file=book.xlsx

  $ gsheet workbook:names --file=book.xlsx add --name Amounts --refersTo "Data!$B$2:$B$4" --inPlace

  $ gsheet workbook:names --file=book.xlsx remove --name Amounts --inPlace
```

_See code: [src/commands/workbook/names.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/workbook/names.ts)_

## `google-sheet workbook:read`

Read cell values and formula freshness from a local XLSX workbook range

```
USAGE
  $ google-sheet workbook:read -f <value> --range <value> [-h] [--mode unformatted|formatted|formula] [--maxRows
    <value>] [--maxCols <value>] [--includeEmpty] [-r]

FLAGS
  -f, --file=<value>     (required) Path to the local XLSX file to read
  -h, --help             Show CLI help.
  -r, --rawOutput        Get the raw output as a JSON string
      --includeEmpty     Include trailing empty cells in output rows
      --maxCols=<value>  Maximum columns to read
      --maxRows=<value>  Maximum rows to read
      --mode=<option>    [default: unformatted] Value extraction mode ("unformatted", "formatted", or "formula")
                         <options: unformatted|formatted|formula>
      --range=<value>    (required) A1 range to read (e.g. "Sheet1!A1:D10")

DESCRIPTION
  Read cell values and formula freshness from a local XLSX workbook range

EXAMPLES
  $ gsheet workbook:read --file=report.xlsx --range='Sheet1!A1:D10'

  $ gsheet workbook:read --file=report.xlsx --range='Sheet1!A1:D10' --mode=formatted

  $ gsheet workbook:read --file=report.xlsx --range='Sheet1!A1:D10' --rawOutput
```

_See code: [src/commands/workbook/read.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/workbook/read.ts)_

## `google-sheet workbook:write [DATA]`

Write tabular data or a ReportDocument into a local XLSX workbook with atomic persistence

```
USAGE
  $ google-sheet workbook:write [DATA] [-h] [-f <value>] [-o <value>] [--inPlace] [-i <value>] [--inputFormat
    json|csv] [-t <value>] [--startCell <value>] [--dryRun] [--overwrite] [--overwriteFormulas] [--cells <value>]
    [--discardUnsupported] [-r]

ARGUMENTS
  [DATA]  Data as a JSON string (nested 2D array or ReportDocument object)

FLAGS
  -f, --file=<value>            Path to existing template XLSX file to load and modify
  -h, --help                    Show CLI help.
  -i, --input=<value>           Path to input data file (JSON or CSV) or "-" for stdin
  -o, --output=<value>          Destination path for the output XLSX file
  -r, --rawOutput               Get the raw output as a JSON string
  -t, --worksheetTitle=<value>  [default: Sheet1] Target worksheet name (default: "Sheet1")
      --cells=<value>           Sparse cell writes as JSON: {"B5":"value","J8":"GD_WEB1"} or [{"a1":"B5","value":"x"}].
                                Each entry writes exactly one cell; every other cell is untouched. Mutually exclusive
                                with positional data and --input
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview changes without modifying or saving files
      --inPlace                 Modify the --file workbook in place (creates automatic .bak backup)
      --inputFormat=<option>    Format of input file ("json" or "csv")
                                <options: json|csv>
      --overwrite               Allow overwriting existing destination output file
      --overwriteFormulas       Allow overwriting existing formula cells in the workbook
      --startCell=<value>       [default: A1] Top-left cell address to place data (e.g. "A1", "B2")

DESCRIPTION
  Write tabular data or a ReportDocument into a local XLSX workbook with atomic persistence

EXAMPLES
  $ gsheet workbook:write --output=new.xlsx '[["Name", "Amount"], ["Alice", 100], ["Bob", 200]]'

  $ gsheet workbook:write --file=template.xlsx --output=filled.xlsx --input=data.csv --worksheetTitle="Summary"

  $ gsheet workbook:write --file=existing.xlsx --inPlace --input=rows.json --startCell="B5" --overwriteFormulas
```

_See code: [src/commands/workbook/write.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/workbook/write.ts)_
