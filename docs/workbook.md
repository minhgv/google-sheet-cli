`google-sheet workbook`
=======================

Manage local Excel (XLSX) workbooks

* [`google-sheet workbook:inspect`](#google-sheet-workbookinspect)
* [`google-sheet workbook:read`](#google-sheet-workbookread)
* [`google-sheet workbook:write [DATA]`](#google-sheet-workbookwrite-data)

## `google-sheet workbook:inspect`

Inspect metadata, sheets, defined names, and capabilities of a local XLSX workbook

```
USAGE
  $ google-sheet workbook:inspect -f <value> [-h] [-r]

FLAGS
  -f, --file=<value>  (required) Path to the local XLSX file to inspect
  -h, --help          Show CLI help.
  -r, --rawOutput     Get the raw output as a JSON string

DESCRIPTION
  Inspect metadata, sheets, defined names, and capabilities of a local XLSX workbook

EXAMPLES
  $ gsheet workbook:inspect --file=report.xlsx

  $ gsheet workbook:inspect --file=report.xlsx --rawOutput
```

_See code: [src/commands/workbook/inspect.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/workbook/inspect.ts)_

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

_See code: [src/commands/workbook/read.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/workbook/read.ts)_

## `google-sheet workbook:write [DATA]`

Write tabular data or a ReportDocument into a local XLSX workbook with atomic persistence

```
USAGE
  $ google-sheet workbook:write [DATA] [-h] [-f <value>] [-o <value>] [--inPlace] [-i <value>] [--inputFormat
    json|csv] [-t <value>] [--startCell <value>] [--dryRun] [--overwrite] [--overwriteFormulas] [-r]

ARGUMENTS
  [DATA]  Data as a JSON string (nested 2D array or ReportDocument object)

FLAGS
  -f, --file=<value>            Path to existing template XLSX file to load and modify
  -h, --help                    Show CLI help.
  -i, --input=<value>           Path to input data file (JSON or CSV) or "-" for stdin
  -o, --output=<value>          Destination path for the output XLSX file
  -r, --rawOutput               Get the raw output as a JSON string
  -t, --worksheetTitle=<value>  [default: Sheet1] Target worksheet name (default: "Sheet1")
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

_See code: [src/commands/workbook/write.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/workbook/write.ts)_
