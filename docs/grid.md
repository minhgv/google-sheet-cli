`google-sheet grid`
===================

Delete rows or columns from a worksheet. Data after the deleted range shifts up/left. --dryRun previews the values about to be removed. With --workbook the delete runs on a local XLSX file instead of Google Sheets.

* [`google-sheet grid:delete`](#google-sheet-griddelete)
* [`google-sheet grid:freeze`](#google-sheet-gridfreeze)
* [`google-sheet grid:hide`](#google-sheet-gridhide)
* [`google-sheet grid:insert`](#google-sheet-gridinsert)
* [`google-sheet grid:resize`](#google-sheet-gridresize)

## `google-sheet grid:delete`

Delete rows or columns from a worksheet. Data after the deleted range shifts up/left. --dryRun previews the values about to be removed. With --workbook the delete runs on a local XLSX file instead of Google Sheets.

```
USAGE
  $ google-sheet grid:delete --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [--redacted] [-c <value>]
    [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>]
    [-o <value>] [--inPlace] [--discardUnsupported] [--count <value>] [--dryRun] [--force] [--updateRefs]

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
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview the mutation without applying it
      --force                   Local backend only: adjust merged ranges that intersect the deleted span instead of
                                refusing
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --start=<value>           (required) The 1-based first row/column index to affect
      --updateRefs              Local backend only: rewrite same-sheet formula references and defined names affected by
                                the delete (refs fully inside the deleted span become #REF!); cross-sheet references are
                                left untouched
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
  Delete rows or columns from a worksheet. Data after the deleted range shifts up/left. --dryRun previews the values
  about to be removed. With --workbook the delete runs on a local XLSX file instead of Google Sheets.

EXAMPLES
  $ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=10 --count=3

  $ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=5 --dryRun

  $ gsheet grid:delete --workbook=template.xlsx -t "Functional effort" --dimension=ROWS --start=8 --count=5 --inPlace

  $ gsheet grid:delete --workbook=template.xlsx --dimension=ROWS --start=8 --count=2 --updateRefs --inPlace
```

_See code: [src/commands/grid/delete.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/grid/delete.ts)_

## `google-sheet grid:freeze`

Freeze or unfreeze rows and columns on a worksheet. At least one of --rows/--columns is required; 0 unfreezes that axis. With --workbook the freeze runs on a local XLSX file instead of Google Sheets.

```
USAGE
  $ google-sheet grid:freeze [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>] [-o <value>] [--inPlace]
    [--discardUnsupported] [--rows <value>] [--columns <value>] [--dryRun]

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
      --columns=<value>         Number of columns to freeze (0 unfreezes)
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview the mutation without applying it
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --rows=<value>            Number of rows to freeze (0 unfreezes)
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
  Freeze or unfreeze rows and columns on a worksheet. At least one of --rows/--columns is required; 0 unfreezes that
  axis. With --workbook the freeze runs on a local XLSX file instead of Google Sheets.

EXAMPLES
  $ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1

  $ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1 --columns=2

  $ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=0 --columns=0

  $ gsheet grid:freeze --workbook=template.xlsx -t Sheet1 --rows=1 --columns=2 --inPlace
```

_See code: [src/commands/grid/freeze.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/grid/freeze.ts)_

## `google-sheet grid:hide`

Hide or unhide rows or columns on a worksheet. With --workbook the hide runs on a local XLSX file instead of Google Sheets.

```
USAGE
  $ google-sheet grid:hide --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [--redacted] [-c <value>]
    [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>]
    [-o <value>] [--inPlace] [--discardUnsupported] [--count <value>] [--dryRun] [--unhide]

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
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview the mutation without applying it
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --start=<value>           (required) The 1-based first row/column index to affect
      --unhide                  Unhide instead of hide
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
  Hide or unhide rows or columns on a worksheet. With --workbook the hide runs on a local XLSX file instead of Google
  Sheets.

EXAMPLES
  $ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=8 --count=3

  $ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=2 --unhide

  $ gsheet grid:hide --workbook=template.xlsx -t Sheet1 --dimension=COLUMNS --start=2 --inPlace
```

_See code: [src/commands/grid/hide.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/grid/hide.ts)_

## `google-sheet grid:insert`

Insert rows or columns into a worksheet at a position. Existing data at and after the position shifts down/right. With --workbook the insert runs on a local XLSX file instead of Google Sheets.

```
USAGE
  $ google-sheet grid:insert --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [--redacted] [-c <value>]
    [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>]
    [-o <value>] [--inPlace] [--discardUnsupported] [--count <value>] [--dryRun] [--inheritFromBefore] [--force]
    [--updateRefs]

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
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview the mutation without applying it
      --force                   Local backend only: adjust merged ranges that intersect the insertion boundary instead
                                of refusing
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --inheritFromBefore       Inherit formatting from the row/column before instead of after (matches "insert
                                above/left")
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --start=<value>           (required) The 1-based first row/column index to affect
      --updateRefs              Local backend only: rewrite same-sheet formula references and defined names affected by
                                the insert (refs inside a deleted range would become #REF!); cross-sheet references are
                                left untouched
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
  Insert rows or columns into a worksheet at a position. Existing data at and after the position shifts down/right. With
  --workbook the insert runs on a local XLSX file instead of Google Sheets.

EXAMPLES
  $ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=5 --count=2

  $ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=3 --inheritFromBefore

  $ gsheet grid:insert --workbook=template.xlsx --worksheetTitle="Functional effort" --dimension=ROWS --start=8 --count=3 --inPlace

  $ gsheet grid:insert --workbook=template.xlsx --dimension=ROWS --start=8 --count=2 --updateRefs --inPlace

  $ gsheet grid:insert --workbook=template.xlsx -t Sheet1 --dimension=ROWS --start=8 --count=3 --dryRun
```

_See code: [src/commands/grid/insert.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/grid/insert.ts)_

## `google-sheet grid:resize`

Resize rows or columns to an explicit pixel size, or auto-size them to their content. With --workbook the resize runs on a local XLSX file: pixels are converted into Excel units (columns: width characters; rows: height points) and --auto sizes columns to their longest cell text (capped) / resets rows to the default height.

```
USAGE
  $ google-sheet grid:resize --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [--redacted] [-c <value>]
    [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>]
    [-o <value>] [--inPlace] [--discardUnsupported] [--count <value>] [--dryRun] [--pixels <value> | --auto]

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
      --auto                    Auto-size to content
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --discardUnsupported      Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would
                                be dropped by the local engine
      --dryRun                  Preview the mutation without applying it
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --pixels=<value>          Explicit pixel size
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --start=<value>           (required) The 1-based first row/column index to affect
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
  Resize rows or columns to an explicit pixel size, or auto-size them to their content. With --workbook the resize runs
  on a local XLSX file: pixels are converted into Excel units (columns: width characters; rows: height points) and
  --auto sizes columns to their longest cell text (capped) / resets rows to the default height.

EXAMPLES
  $ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --count=4 --pixels=120

  $ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --auto

  $ gsheet grid:resize --workbook=template.xlsx -t Sheet1 --dimension=COLUMNS --start=1 --count=4 --auto --inPlace
```

_See code: [src/commands/grid/resize.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/grid/resize.ts)_
