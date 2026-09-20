`google-sheet worksheet`
========================

Manage worksheets

* [`google-sheet worksheet:add`](#google-sheet-worksheetadd)
* [`google-sheet worksheet:copy`](#google-sheet-worksheetcopy)
* [`google-sheet worksheet:get`](#google-sheet-worksheetget)
* [`google-sheet worksheet:remove`](#google-sheet-worksheetremove)
* [`google-sheet worksheet:rename`](#google-sheet-worksheetrename)

## `google-sheet worksheet:add`

Add a worksheet with the specified title to the spreadsheet

```
USAGE
  $ google-sheet worksheet:add -t <value> -s <value> [-h] [-r] [-j] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use

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
  Add a worksheet with the specified title to the spreadsheet

EXAMPLES
  $ gsheet worksheet:add --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>
  Worksheet "<worksheetTitle>" (<id>) successfully created
```

_See code: [src/commands/worksheet/add.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/worksheet/add.ts)_

## `google-sheet worksheet:copy`

Copy a worksheet to an explicit destination spreadsheet through the Sheets API. The copy carries the source title; when the destination already has a worksheet with that title, the server assigns a unique one (e.g. "Sheet1 Copy") and the receipt reports it. Values and formulas are carried over; the source worksheet is left untouched.

```
USAGE
  $ google-sheet worksheet:copy -s <value> -t <value> --destinationSpreadsheetId <value> [-h] [-r] [-j] [-c <value>]
    [-p <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                              Show CLI help.
  -j, --json                              Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                          failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                         Get the raw output as a JSON string
  -s, --spreadsheetId=<value>             (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>            (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --destinationSpreadsheetId=<value>  (required) ID of the spreadsheet to copy the worksheet into

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
  Copy a worksheet to an explicit destination spreadsheet through the Sheets API. The copy carries the source title;
  when the destination already has a worksheet with that title, the server assigns a unique one (e.g. "Sheet1 Copy") and
  the receipt reports it. Values and formulas are carried over; the source worksheet is left untouched.

EXAMPLES
  $ gsheet worksheet:copy --spreadsheetId=<id> --worksheetTitle=<worksheetTitle> --destinationSpreadsheetId=<destId>
  Worksheet "<worksheetTitle>" copied to spreadsheet <destId> as sheet <sheetId>
```

_See code: [src/commands/worksheet/copy.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/worksheet/copy.ts)_

## `google-sheet worksheet:get`

Get info for a specific worksheet

```
USAGE
  $ google-sheet worksheet:get -t <value> -s <value> [-h] [-r] [-j] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use

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
  Get info for a specific worksheet

EXAMPLES
  $ gsheet worksheet:get --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>
  Fetched "<worksheetTitle>" (<id>)
```

_See code: [src/commands/worksheet/get.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/worksheet/get.ts)_

## `google-sheet worksheet:remove`

Remove a worksheet with the specified title from the spreadsheet

```
USAGE
  $ google-sheet worksheet:remove -t <value> -s <value> [-h] [-r] [-j] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use

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
  Remove a worksheet with the specified title from the spreadsheet

EXAMPLES
  $ gsheet worksheet:remove --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>
  Worksheet "<worksheetTitle>" successfully removed
```

_See code: [src/commands/worksheet/remove.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/worksheet/remove.ts)_

## `google-sheet worksheet:rename`

Add a worksheet with the specified title to the spreadsheet

```
USAGE
  $ google-sheet worksheet:rename -t <value> --newWorksheetTitle <value> -s <value> [-h] [-r] [-j] [-c <value>] [-p
    <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                       Show CLI help.
  -j, --json                       Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                   failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>     (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --newWorksheetTitle=<value>  (required) New title of the worksheet to use

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
  Add a worksheet with the specified title to the spreadsheet

EXAMPLES
  $ gsheet worksheet:rename --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --newWorksheetTitle=<newWorksheetTitle>
  Worksheet "<worksheetTitle>" successfully renamed to "<newWorksheetTitle>"
```

_See code: [src/commands/worksheet/rename.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/worksheet/rename.ts)_
