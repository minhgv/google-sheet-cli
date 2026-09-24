`google-sheet data`
===================

Manage data in worksheet

* [`google-sheet data:append [DATA]`](#google-sheet-dataappend-data)
* [`google-sheet data:append-table [DATA]`](#google-sheet-dataappend-table-data)
* [`google-sheet data:batch-get`](#google-sheet-databatch-get)
* [`google-sheet data:batch-update [DATA]`](#google-sheet-databatch-update-data)
* [`google-sheet data:clear`](#google-sheet-dataclear)
* [`google-sheet data:export-csv`](#google-sheet-dataexport-csv)
* [`google-sheet data:find`](#google-sheet-datafind)
* [`google-sheet data:get`](#google-sheet-dataget)
* [`google-sheet data:schema`](#google-sheet-dataschema)
* [`google-sheet data:update [DATA]`](#google-sheet-dataupdate-data)
* [`google-sheet data:upsert [DATA]`](#google-sheet-dataupsert-data)
* [`google-sheet data:validate`](#google-sheet-datavalidate)

## `google-sheet data:append [DATA]`

Append cells with the specified data after the last row in starting col (legacy append semantics; for native table appends use data:append-table)

```
USAGE
  $ google-sheet data:append [DATA] -t <value> -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>]
    [-f <value>] [--useOauth] [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [--minCol <value>] [-i <value>]
    [--inputFormat json|csv] [--dryRun]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
  -j, --json                       Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                   failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>     (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -v, --valueInputOption=<option>  [default: RAW, env: VALUE_INPUT_OPTION] The style of the input ("RAW" or
                                   "USER_ENTERED")
                                   <options: RAW|USER_ENTERED>
      --dryRun                     Preview data without modifying worksheet
      --inputFormat=<option>       Format of input file ("json" or "csv")
                                   <options: json|csv>
      --minCol=<value>             [default: 1] The optional starting col of the operation
      --redacted                   [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                   from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                   counts, statuses and outcome states are kept.

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Append cells with the specified data after the last row in starting col (legacy append semantics; for native table
  appends use data:append-table)

EXAMPLES
  $ gsheet data:append --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> '[["1", "2", "3"]]'
  Data successfully appended to "<worksheetTitle>"

  $ gsheet data:append --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=rows.csv
  Data successfully appended to "<worksheetTitle>"

  $ gsheet data:append --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=- --inputFormat=json < rows.json
  Data successfully appended to "<worksheetTitle>"
```

_See code: [src/commands/data/append.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/append.ts)_

## `google-sheet data:append-table [DATA]`

Append table data using native Google Sheets values.append API (prevents client-side race conditions)

```
USAGE
  $ google-sheet data:append-table [DATA] -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [-t <value>] [--range <value>] [-i <value>] [--inputFormat json|csv] [-v
    RAW|USER_ENTERED] [--insertDataOption OVERWRITE|INSERT_ROWS]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
  -j, --json                       Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                   failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>     Title of the worksheet to append table into
  -v, --valueInputOption=<option>  [default: RAW, env: VALUE_INPUT_OPTION] The style of the input ("RAW" or
                                   "USER_ENTERED")
                                   <options: RAW|USER_ENTERED>
      --inputFormat=<option>       Format of input file ("json" or "csv")
                                   <options: json|csv>
      --insertDataOption=<option>  How the input data should be inserted: "OVERWRITE" (overwrites blank cells) or
                                   "INSERT_ROWS" (inserts new rows)
                                   <options: OVERWRITE|INSERT_ROWS>
      --range=<value>              Specific table range to search for existing data table (e.g. "Sheet1!A1")
      --redacted                   [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                   from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                   counts, statuses and outcome states are kept.

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Append table data using native Google Sheets values.append API (prevents client-side race conditions)

EXAMPLES
  $ gsheet data:append-table --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> '[["1", "2", "3"]]'

  $ gsheet data:append-table --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=table.csv

  $ gsheet data:append-table --spreadsheetId=<spreadsheetId> --range='Sheet1!A1' --input=- --inputFormat=json < table.json
```

_See code: [src/commands/data/append-table.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/append-table.ts)_

## `google-sheet data:batch-get`

Fetch cell data across multiple ranges in a single batch request

```
USAGE
  $ google-sheet data:batch-get -s <value> --ranges <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>]
    [-f <value>] [--useOauth] [--clientSecretFile <value>] [--valueRenderOption
    FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA] [--dateTimeRenderOption FORMATTED_STRING|SERIAL_NUMBER] [--chunkSize
    <value>]

FLAGS
  -h, --help                           Show CLI help.
  -j, --json                           Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                       failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                      Get the raw output as a JSON string
  -s, --spreadsheetId=<value>          (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --chunkSize=<value>              Maximum number of ranges per batch chunk (default: 50)
      --dateTimeRenderOption=<option>  [default: FORMATTED_STRING] Determines how dates, times, and durations should be
                                       rendered
                                       <options: FORMATTED_STRING|SERIAL_NUMBER>
      --ranges=<value>                 (required) JSON array of A1 range strings to query (e.g. '["Sheet1!A1:B10",
                                       "Sheet2!C1:D5"]')
      --redacted                       [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and
                                       credentials from error envelopes and dry-run diagnostics before they are written.
                                       Coordinates, counts, statuses and outcome states are kept.
      --valueRenderOption=<option>     [default: FORMATTED_VALUE] Determines how values should be rendered in the output
                                       <options: FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA>

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Fetch cell data across multiple ranges in a single batch request

EXAMPLES
  $ gsheet data:batch-get --spreadsheetId=<spreadsheetId> --ranges='["Sheet1!A1:B10", "Sheet2!C1:D5"]'

  $ gsheet data:batch-get --spreadsheetId=<spreadsheetId> --ranges='["Sheet1!A1:B10"]' --valueRenderOption=UNFORMATTED_VALUE
```

_See code: [src/commands/data/batch-get.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/batch-get.ts)_

## `google-sheet data:batch-update [DATA]`

Update cell data across multiple ranges in a single batch request

```
USAGE
  $ google-sheet data:batch-update [DATA] -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [-i <value>] [--inputFormat json] [--dryRun]
    [--overwriteFormulas] [--chunkByteSize <value>] [--maxRowsPerChunk <value>]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input JSON file or "-" for stdin
  -j, --json                       Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                   failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -v, --valueInputOption=<option>  [default: RAW, env: VALUE_INPUT_OPTION] The style of the input ("RAW" or
                                   "USER_ENTERED")
                                   <options: RAW|USER_ENTERED>
      --chunkByteSize=<value>      Maximum byte size per payload chunk
      --dryRun                     Preview batch update changes without modifying spreadsheet
      --inputFormat=<option>       [default: json] Format of input file (currently "json")
                                   <options: json>
      --maxRowsPerChunk=<value>    Maximum rows per payload chunk
      --overwriteFormulas          Allow overwriting existing formula cells
      --redacted                   [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                   from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                   counts, statuses and outcome states are kept.

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Update cell data across multiple ranges in a single batch request

EXAMPLES
  $ gsheet data:batch-update --spreadsheetId=<spreadsheetId> '[{"range": "Sheet1!A1:B2", "values": [["1", "2"], ["3", "4"]]}]'

  $ gsheet data:batch-update --spreadsheetId=<spreadsheetId> --input=updates.json

  $ gsheet data:batch-update --spreadsheetId=<spreadsheetId> --input=- < updates.json
```

_See code: [src/commands/data/batch-update.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/batch-update.ts)_

## `google-sheet data:clear`

Clear cell values in an explicitly bounded range, keeping every other cell property (values-only clear). With --workbook the clear runs on a local XLSX file instead of Google Sheets.

```
USAGE
  $ google-sheet data:clear --range <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [-t <value>] [-s <value>] [--workbook <value>] [-o <value>] [--inPlace]
    [--discardUnsupported] [--dryRun] [--overwriteFormulas]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -o, --output=<value>          Destination path for the modified XLSX file (without it and without --inPlace the
                                mutation is refused unless --dryRun)
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   [env: SPREADSHEET_ID] ID of the spreadsheet to use (Google Sheets target)
  -t, --worksheetTitle=<value>  [env: WORKSHEET_TITLE] Title of the worksheet to use (Google Sheets target; also selects
                                the sheet inside --workbook)
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview the clear without modifying the worksheet
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --overwriteFormulas       Allow the clear to overwrite cells that contain formulas
      --range=<value>           (required) A1 range to clear; must be bounded on both axes (e.g. "Sheet1!A1:D20")
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --workbook=<value>        Path to a local .xlsx workbook to mutate instead of the Google Sheets target. Long name
                                only: the shared short -f belongs to --credentialsFile

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Clear cell values in an explicitly bounded range, keeping every other cell property (values-only clear). With
  --workbook the clear runs on a local XLSX file instead of Google Sheets.

EXAMPLES
  $ gsheet data:clear --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range="Sheet1!A2:D20"
  Successfully cleared 76 cells in "Sheet1" ("Sheet1!A2:D20")

  $ gsheet data:clear --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range="Sheet1!A2:D20" --dryRun
  Clear preview (dry run): 76 cells in "Sheet1!A2:D20" would be cleared

  $ gsheet data:clear --workbook=template.xlsx --range="Functional effort!A8:J30" --inPlace
```

_See code: [src/commands/data/clear.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/clear.ts)_

## `google-sheet data:export-csv`

Export worksheet data as deterministic RFC 4180 CSV, from a Google Sheets worksheet or a local XLSX workbook. Cell rendering is selectable (raw values, formatted display strings, or formula text where the backend stores it), and the injection policy for string cells is explicit: safe (default) prefixes dangerous leading characters with an apostrophe, preserve emits them byte-faithfully and says so in the receipt. Local XLSX formula results are the cached values from the last save by the producing application - they are never recalculated.

```
USAGE
  $ google-sheet data:export-csv [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [-s <value>] [-t <value>] [--range <value>] [--workbook <value>] [--mode
    raw|formatted|formula] [--injection safe|preserve] [-o <value>] [--overwrite]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -o, --output=<value>          Path of the CSV file to write atomically (an existing file needs --overwrite). Without
                                it, the CSV goes to stdout and no receipt is printed
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   [env: SPREADSHEET_ID] ID of the spreadsheet to export (Google Sheets source)
  -t, --worksheetTitle=<value>  [env: WORKSHEET_TITLE] Title of the worksheet to export (Google Sheets source)
      --injection=<option>      [default: safe] Formula-injection policy for string cells: "safe" prefixes dangerous
                                leading characters (=, +, -, @, tab, CR) with an apostrophe; "preserve" emits them
                                byte-faithfully and warns in the receipt
                                <options: safe|preserve>
      --mode=<option>           [default: raw] Cell rendering: "raw" machine values (default), "formatted" display
                                strings, "formula" formula text where the backend stores it (XLSX always reports cached
                                formula results)
                                <options: raw|formatted|formula>
      --overwrite               Replace an existing output file (default: refuse)
      --range=<value>           A1 range to export, e.g. "Sheet1!A1:D10" (optional for the Google Sheets source, where
                                it bounds the used range of --worksheetTitle; required with --workbook)
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --workbook=<value>        Path to a local .xlsx workbook to export instead of the Google Sheets source. Long name
                                only: the shared short -f belongs to --credentialsFile

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Export worksheet data as deterministic RFC 4180 CSV, from a Google Sheets worksheet or a local XLSX workbook. Cell
  rendering is selectable (raw values, formatted display strings, or formula text where the backend stores it), and the
  injection policy for string cells is explicit: safe (default) prefixes dangerous leading characters with an
  apostrophe, preserve emits them byte-faithfully and says so in the receipt. Local XLSX formula results are the cached
  values from the last save by the producing application - they are never recalculated.

EXAMPLES
  $ gsheet data:export-csv --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --output=./data.csv
  Exported 3 row(s) x 2 column(s) from <spreadsheetId>/Sheet1 (Sheet1!A1:B3) to ./data.csv (48 bytes, mode: raw, injection: safe, transformed cells: 0)

  $ gsheet data:export-csv --workbook=./report.xlsx --range='Sheet1!A1:D10' --mode=formatted --output=./report.csv --overwrite
  Exported 10 row(s) x 4 column(s) from ./report.xlsx (Sheet1!A1:D10) to ./report.csv (312 bytes, mode: formatted, injection: safe, transformed cells: 1)

  $ gsheet data:export-csv --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range=A1:B10 --mode=formula --injection=preserve
  =B2*2
  =SUM(A1:A10)
```

_See code: [src/commands/data/export-csv.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/export-csv.ts)_

## `google-sheet data:find`

Locate cells matching a condition and return their A1 coordinates. Pure read: scans the range once and reports where matches live, so follow-up writes can target exact cells without downloading the whole sheet.

```
USAGE
  $ google-sheet data:find -s <value> -t <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f
    <value>] [--useOauth] [--clientSecretFile <value>] [--columns <value> | -x] [--sort <value>] [--filter <value>]
    [--output csv|json|yaml |  | [--csv | --no-truncate]] [--no-header | ] [--range <value>] [--equals <value> |
    --contains <value> | --regex <value>] [--column <value> | --header <value>] [--ignoreCase] [--valueRenderOption
    FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA] [--dateTimeRenderOption FORMATTED_STRING|SERIAL_NUMBER] [--limit <value>]
    [--first] [--byRow]

FLAGS
  -h, --help                           Show CLI help.
  -j, --json                           Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                       failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                      Get the raw output as a JSON string
  -s, --spreadsheetId=<value>          (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>         (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -x, --extended                       show extra columns
      --byRow                          Collapse matches to unique rows and include full row values
      --column=<value>                 Restrict matches to one column by A1 letter (e.g. "B")
      --columns=<value>                only show provided columns (comma-separated)
      --contains=<value>               Match cells containing this substring
      --csv                            output is csv format [alias: --output=csv]
      --dateTimeRenderOption=<option>  [default: FORMATTED_STRING] Determines how dates, times, and durations are
                                       rendered before matching
                                       <options: FORMATTED_STRING|SERIAL_NUMBER>
      --equals=<value>                 Match cells exactly equal to this value
      --filter=<value>                 filter property by partial string matching, ex: name=foo
      --first                          Return only the first match (alias for --limit=1)
      --header=<value>                 Restrict matches to the column whose first scanned row equals this header text
      --[no-]ignoreCase                Case-insensitive matching
      --limit=<value>                  [default: 100] Maximum matches to return (matchCount still reports the true
                                       total)
      --no-header                      hide table header from output
      --no-truncate                    do not truncate output to fit screen
      --output=<option>                output in a more machine friendly format
                                       <options: csv|json|yaml>
      --range=<value>                  The A1 range bounding the scan (default: whole worksheet)
      --redacted                       [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and
                                       credentials from error envelopes and dry-run diagnostics before they are written.
                                       Coordinates, counts, statuses and outcome states are kept.
      --regex=<value>                  Match cells against this regular expression
      --sort=<value>                   property to sort by (prepend '-' for descending)
      --valueRenderOption=<option>     [default: FORMATTED_VALUE] Determines how cell values are rendered before
                                       matching
                                       <options: FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA>

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Locate cells matching a condition and return their A1 coordinates. Pure read: scans the range once and reports where
  matches live, so follow-up writes can target exact cells without downloading the whole sheet.

EXAMPLES
  $ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --equals="NV0123" --column=B --first --rawOutput

  $ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --contains="invoice" --byRow

  $ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --regex="^ERR-" --range=A1:K200

  $ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --equals="Paid" --header="Status"
```

_See code: [src/commands/data/find.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/find.ts)_

## `google-sheet data:get`

Returns cell data

```
USAGE
  $ google-sheet data:get -s <value> -t <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f
    <value>] [--useOauth] [--clientSecretFile <value>] [--columns <value> | -x] [--sort <value>] [--filter <value>]
    [--output csv|json|yaml |  | [--csv | --no-truncate]] [--no-header | ] [-w] [--range <value>] [--minRow <value>]
    [--minCol <value>] [--maxRow <value>] [--maxCol <value>] [--valueRenderOption
    FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA] [--dateTimeRenderOption FORMATTED_STRING|SERIAL_NUMBER]

FLAGS
  -h, --help                           Show CLI help.
  -j, --json                           Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                       failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                      Get the raw output as a JSON string
  -s, --spreadsheetId=<value>          (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>         (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -w, --hasHeaderRow                   If the first row should be treated as header row
  -x, --extended                       show extra columns
      --columns=<value>                only show provided columns (comma-separated)
      --csv                            output is csv format [alias: --output=csv]
      --dateTimeRenderOption=<option>  [default: FORMATTED_STRING] Determines how dates, times, and durations should be
                                       rendered
                                       <options: FORMATTED_STRING|SERIAL_NUMBER>
      --filter=<value>                 filter property by partial string matching, ex: name=foo
      --maxCol=<value>                 The optional ending col of the operation
      --maxRow=<value>                 The optional ending row of the operation
      --minCol=<value>                 [default: 1] The optional starting col of the operation
      --minRow=<value>                 [default: 1] The optional starting row of the operation
      --no-header                      hide table header from output
      --no-truncate                    do not truncate output to fit screen
      --output=<option>                output in a more machine friendly format
                                       <options: csv|json|yaml>
      --range=<value>                  The range to use to query the cells
      --redacted                       [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and
                                       credentials from error envelopes and dry-run diagnostics before they are written.
                                       Coordinates, counts, statuses and outcome states are kept.
      --sort=<value>                   property to sort by (prepend '-' for descending)
      --valueRenderOption=<option>     [default: FORMATTED_VALUE] Determines how values should be rendered in the output
                                       <options: FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA>

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Returns cell data

EXAMPLES
  $ gsheet data:get --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>
  (a)  (b)  (c)
  A1   B1   C1
  A2   B2   C2
  A3   B3   C3
```

_See code: [src/commands/data/get.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/get.ts)_

## `google-sheet data:schema`

Discover the column schema of a worksheet (read-only)

```
USAGE
  $ google-sheet data:schema -s <value> -t <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f
    <value>] [--useOauth] [--clientSecretFile <value>] [--columns <value> | -x] [--sort <value>] [--filter <value>]
    [--output csv|json|yaml |  | [--csv | --no-truncate]] [--no-header | ] [--minRow <value>] [--minCol <value>]
    [--maxRow <value>] [--maxCol <value>]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -x, --extended                show extra columns
      --columns=<value>         only show provided columns (comma-separated)
      --csv                     output is csv format [alias: --output=csv]
      --filter=<value>          filter property by partial string matching, ex: name=foo
      --maxCol=<value>          [default: 26] The last column of the sampled range
      --maxRow=<value>          [default: 100] The last row of the sampled range
      --minCol=<value>          [default: 1] The first column of the sampled range
      --minRow=<value>          [default: 1] The first row of the sampled range (holds the header row)
      --no-header               hide table header from output
      --no-truncate             do not truncate output to fit screen
      --output=<option>         output in a more machine friendly format
                                <options: csv|json|yaml>
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --sort=<value>            property to sort by (prepend '-' for descending)

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Discover the column schema of a worksheet (read-only)

  Reports the actual header row with its absolute A1 coordinates, inferred cell types,
  formula presence, data-validation rules and named ranges over a bounded sample. The
  header row is the first row of the sampled range. Inferred types describe the sample,
  not an authoritative schema - Google stores dates as numbers and formatting decides
  what they look like.

EXAMPLES
  $ gsheet data:schema --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>
  HEADER  COLUMN  TYPES               FORMULAS  VALIDATIONS
  Name    A       string (3)          0
  Amount  B       number (2), string  0
  Flag    C       boolean (3)         0         ONE_OF_LIST (3)
```

_See code: [src/commands/data/schema.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/schema.ts)_

## `google-sheet data:update [DATA]`

Updates cells with the specified data

```
USAGE
  $ google-sheet data:update [DATA] -t <value> -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>]
    [-f <value>] [--useOauth] [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [--minRow <value>] [--minCol <value>]
    [-i <value>] [--inputFormat json|csv] [--dryRun]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
  -j, --json                       Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                   failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>     (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -v, --valueInputOption=<option>  [default: RAW, env: VALUE_INPUT_OPTION] The style of the input ("RAW" or
                                   "USER_ENTERED")
                                   <options: RAW|USER_ENTERED>
      --dryRun                     Preview update without modifying worksheet
      --inputFormat=<option>       Format of input file ("json" or "csv")
                                   <options: json|csv>
      --minCol=<value>             [default: 1] The optional starting col of the operation
      --minRow=<value>             [default: 1] The optional starting row of the operation
      --redacted                   [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                   from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                   counts, statuses and outcome states are kept.

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Updates cells with the specified data

EXAMPLES
  $ gsheet data:update --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> '[["1", "2", "3"]]'
  Data successfully updated in "<worksheetTitle>"

  $ gsheet data:update --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=update.csv
  Data successfully updated in "<worksheetTitle>"

  $ gsheet data:update --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=- --inputFormat=json < update.json
  Data successfully updated in "<worksheetTitle>"
```

_See code: [src/commands/data/update.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/update.ts)_

## `google-sheet data:upsert [DATA]`

Upsert rows keyed by one column: existing keys update the supplied cells only, new keys append below the table

```
USAGE
  $ google-sheet data:upsert [DATA] -t <value> -s <value> --key <value> [-h] [-r] [-j] [--redacted] [-c <value>]
    [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [--range <value>] [-i
    <value>] [--inputFormat json|csv] [--dryRun] [--overwriteFormulas]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
  -j, --json                       Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                   failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>     (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -v, --valueInputOption=<option>  [default: RAW, env: VALUE_INPUT_OPTION] The style of the input ("RAW" or
                                   "USER_ENTERED")
                                   <options: RAW|USER_ENTERED>
      --dryRun                     Report added/updated/unchanged rows and planned ranges without writing
      --inputFormat=<option>       Format of input file ("json" or "csv")
                                   <options: json|csv>
      --key=<value>                (required) Header name of the single key column used to match input rows against
                                   existing rows
      --overwriteFormulas          Allow updates to overwrite existing formulas in changed cells
      --range=<value>              A1 range of the existing table including its header row (e.g. "Sheet1!A1:F100");
                                   defaults to the whole worksheet grid
      --redacted                   [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                   from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                   counts, statuses and outcome states are kept.

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Upsert rows keyed by one column: existing keys update the supplied cells only, new keys append below the table

EXAMPLES
  $ gsheet data:upsert --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --key=id '[["id", "score"], ["001", 99]]'
  Upserted into "<worksheetTitle>": 0 added, 1 updated, 0 unchanged

  $ gsheet data:upsert --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --key=id --input=rows.csv
  Upserted into "<worksheetTitle>": 3 added, 0 updated, 2 unchanged

  $ gsheet data:upsert --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --key=id --input=- --inputFormat=json < rows.json
```

_See code: [src/commands/data/upsert.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/upsert.ts)_

## `google-sheet data:validate`

Validates worksheet data against a TableSchema (read-only)

```
USAGE
  $ google-sheet data:validate -s <value> -t <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f
    <value>] [--useOauth] [--clientSecretFile <value>] [--columns <value> | -x] [--sort <value>] [--filter <value>]
    [--output csv|json|yaml |  | [--csv | --no-truncate]] [--no-header | ] [--schema <value> | --schemaFile <value>]
    [--minRow <value>] [--minCol <value>] [--maxRow <value>] [--maxCol <value>]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
  -x, --extended                show extra columns
      --columns=<value>         only show provided columns (comma-separated)
      --csv                     output is csv format [alias: --output=csv]
      --filter=<value>          filter property by partial string matching, ex: name=foo
      --maxCol=<value>          [default: 26] The last column of the validated range
      --maxRow=<value>          [default: 100] The last row of the validated range
      --minCol=<value>          [default: 1] The first column of the validated range
      --minRow=<value>          [default: 1] The first row of the validated range (holds the header row)
      --no-header               hide table header from output
      --no-truncate             do not truncate output to fit screen
      --output=<option>         output in a more machine friendly format
                                <options: csv|json|yaml>
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --schema=<value>          The TableSchema as a JSON string
      --schemaFile=<value>      Path to a TableSchema JSON file, or "-" for stdin
      --sort=<value>            property to sort by (prepend '-' for descending)

AUTHENTICATION FLAGS
  -c, --clientEmail=<value>       [env: GSHEET_CLIENT_EMAIL] The client email to use for authentication. Uses the
                                  GSHEET_CLIENT_EMAIL env variable if not provided.
  -f, --credentialsFile=<value>   [env: GSHEET_CREDENTIALS_FILE] Path to the service account JSON file to read the
                                  credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The
                                  clientEmail and privateKey flags take precedence.
  -p, --privateKey=<value>        [env: GSHEET_PRIVATE_KEY] The private key to use for authentication. Uses the
                                  GSHEET_PRIVATE_KEY env variable if not provided.
      --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App
                                  type)
      --useOauth                  [env: GSHEET_USE_OAUTH] Use OAuth 2.0 user authentication instead of service account

DESCRIPTION
  Validates worksheet data against a TableSchema (read-only)

  Reuses the report validation semantics (required, type, enum, unique, min/max, ...),
  reads the sheet with RAW cell values, and reports every violation at its exact sheet
  row and A1 coordinate. Nothing is ever written. The run exits nonzero when the data
  violates a valid schema; a malformed schema is reported as SCHEMA_INVALID instead.
  With --json or --rawOutput the issues arrive inside the error envelope on stderr.

EXAMPLES
  $ gsheet data:validate --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --schema='{"fields":[{"name":"Name","type":"string","required":true},{"name":"Amount","type":"decimal","min":0},{"name":"Code","type":"string","unique":true}]}'
```

_See code: [src/commands/data/validate.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/data/validate.ts)_
