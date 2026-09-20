`google-sheet grid`
===================

Delete rows or columns from a worksheet. Data after the deleted range shifts up/left. --dryRun previews the values about to be removed.

* [`google-sheet grid:delete`](#google-sheet-griddelete)
* [`google-sheet grid:freeze`](#google-sheet-gridfreeze)
* [`google-sheet grid:hide`](#google-sheet-gridhide)
* [`google-sheet grid:insert`](#google-sheet-gridinsert)
* [`google-sheet grid:resize`](#google-sheet-gridresize)

## `google-sheet grid:delete`

Delete rows or columns from a worksheet. Data after the deleted range shifts up/left. --dryRun previews the values about to be removed.

```
USAGE
  $ google-sheet grid:delete -s <value> -t <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [-c
    <value>] [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [--count <value>] [--dryRun]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --dryRun                  Preview the batchUpdate request without applying it
      --start=<value>           (required) The 1-based first row/column index to affect

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
  about to be removed.

EXAMPLES
  $ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=10 --count=3

  $ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=5 --dryRun
```

_See code: [src/commands/grid/delete.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/grid/delete.ts)_

## `google-sheet grid:freeze`

Freeze or unfreeze rows and columns on a worksheet. At least one of --rows/--columns is required; 0 unfreezes that axis.

```
USAGE
  $ google-sheet grid:freeze -s <value> -t <value> [-h] [-r] [-j] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [--rows <value>] [--columns <value>] [--dryRun]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --columns=<value>         Number of columns to freeze (0 unfreezes)
      --dryRun                  Preview the batchUpdate request without applying it
      --rows=<value>            Number of rows to freeze (0 unfreezes)

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
  axis.

EXAMPLES
  $ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1

  $ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1 --columns=2

  $ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=0 --columns=0
```

_See code: [src/commands/grid/freeze.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/grid/freeze.ts)_

## `google-sheet grid:hide`

Hide or unhide rows or columns on a worksheet.

```
USAGE
  $ google-sheet grid:hide -s <value> -t <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [-c
    <value>] [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [--count <value>] [--dryRun] [--unhide]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --dryRun                  Preview the batchUpdate request without applying it
      --start=<value>           (required) The 1-based first row/column index to affect
      --unhide                  Unhide instead of hide

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
  Hide or unhide rows or columns on a worksheet.

EXAMPLES
  $ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=8 --count=3

  $ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=2 --unhide
```

_See code: [src/commands/grid/hide.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/grid/hide.ts)_

## `google-sheet grid:insert`

Insert rows or columns into a worksheet at a position. Existing data at and after the position shifts down/right.

```
USAGE
  $ google-sheet grid:insert -s <value> -t <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [-c
    <value>] [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [--count <value>] [--dryRun]
    [--inheritFromBefore]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --dryRun                  Preview the batchUpdate request without applying it
      --inheritFromBefore       Inherit formatting from the row/column before instead of after (matches "insert
                                above/left")
      --start=<value>           (required) The 1-based first row/column index to affect

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
  Insert rows or columns into a worksheet at a position. Existing data at and after the position shifts down/right.

EXAMPLES
  $ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=5 --count=2

  $ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=3 --inheritFromBefore
```

_See code: [src/commands/grid/insert.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/grid/insert.ts)_

## `google-sheet grid:resize`

Resize rows or columns to an explicit pixel size, or auto-size them to their content.

```
USAGE
  $ google-sheet grid:resize -s <value> -t <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r] [-j] [-c
    <value>] [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [--count <value>] [--dryRun] [--pixels
    <value> | --auto]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --auto                    Auto-size to content
      --count=<value>           [default: 1] How many rows/columns the operation covers
      --dimension=<option>      (required) The dimension to mutate
                                <options: ROWS|COLUMNS>
      --dryRun                  Preview the batchUpdate request without applying it
      --pixels=<value>          Explicit pixel size
      --start=<value>           (required) The 1-based first row/column index to affect

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
  Resize rows or columns to an explicit pixel size, or auto-size them to their content.

EXAMPLES
  $ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --count=4 --pixels=120

  $ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --auto
```

_See code: [src/commands/grid/resize.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/grid/resize.ts)_
