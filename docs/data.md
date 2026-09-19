`google-sheet data`
===================

Manage data in worksheet

* [`google-sheet data:append [DATA]`](#google-sheet-dataappend-data)
* [`google-sheet data:append-table [DATA]`](#google-sheet-dataappend-table-data)
* [`google-sheet data:batch-get`](#google-sheet-databatch-get)
* [`google-sheet data:batch-update [DATA]`](#google-sheet-databatch-update-data)
* [`google-sheet data:find`](#google-sheet-datafind)
* [`google-sheet data:get`](#google-sheet-dataget)
* [`google-sheet data:update [DATA]`](#google-sheet-dataupdate-data)

## `google-sheet data:append [DATA]`

Append cells with the specified data after the last row in starting col (legacy append semantics; for native table appends use data:append-table)

```
USAGE
  $ google-sheet data:append [DATA] -t <value> -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [--minCol <value>] [-i <value>] [--inputFormat
    json|csv] [--dryRun]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
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
  $ google-sheet data:append-table [DATA] -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [-t <value>] [--range <value>] [-i <value>] [--inputFormat json|csv] [-v
    RAW|USER_ENTERED] [--insertDataOption OVERWRITE|INSERT_ROWS]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
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
  $ google-sheet data:batch-get -s <value> --ranges <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [--valueRenderOption FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA]
    [--dateTimeRenderOption FORMATTED_STRING|SERIAL_NUMBER] [--chunkSize <value>]

FLAGS
  -h, --help                           Show CLI help.
  -r, --rawOutput                      Get the raw output as a JSON string
  -s, --spreadsheetId=<value>          (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --chunkSize=<value>              Maximum number of ranges per batch chunk (default: 50)
      --dateTimeRenderOption=<option>  [default: FORMATTED_STRING] Determines how dates, times, and durations should be
                                       rendered
                                       <options: FORMATTED_STRING|SERIAL_NUMBER>
      --ranges=<value>                 (required) JSON array of A1 range strings to query (e.g. '["Sheet1!A1:B10",
                                       "Sheet2!C1:D5"]')
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
  $ google-sheet data:batch-update [DATA] -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [-i <value>] [--inputFormat json] [--dryRun]
    [--overwriteFormulas] [--chunkByteSize <value>] [--maxRowsPerChunk <value>]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input JSON file or "-" for stdin
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

## `google-sheet data:find`

Locate cells matching a condition and return their A1 coordinates. Pure read: scans the range once and reports where matches live, so follow-up writes can target exact cells without downloading the whole sheet.

```
USAGE
  $ google-sheet data:find -s <value> -t <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [--columns <value> | -x] [--sort <value>] [--filter <value>] [--output csv|json|yaml |
    | [--csv | --no-truncate]] [--no-header | ] [--range <value>] [--equals <value> | --contains <value> | --regex
    <value>] [--column <value> | --header <value>] [--ignoreCase] [--valueRenderOption
    FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA] [--dateTimeRenderOption FORMATTED_STRING|SERIAL_NUMBER] [--limit <value>]
    [--first] [--byRow]

FLAGS
  -h, --help                           Show CLI help.
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
  $ google-sheet data:get -s <value> -t <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [--columns <value> | -x] [--sort <value>] [--filter <value>] [--output csv|json|yaml |
    | [--csv | --no-truncate]] [--no-header | ] [-w] [--range <value>] [--minRow <value>] [--minCol <value>] [--maxRow
    <value>] [--maxCol <value>] [--valueRenderOption FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA] [--dateTimeRenderOption
    FORMATTED_STRING|SERIAL_NUMBER]

FLAGS
  -h, --help                           Show CLI help.
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

## `google-sheet data:update [DATA]`

Updates cells with the specified data

```
USAGE
  $ google-sheet data:update [DATA] -t <value> -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [-v RAW|USER_ENTERED] [--minRow <value>] [--minCol <value>] [-i <value>]
    [--inputFormat json|csv] [--dryRun]

ARGUMENTS
  [DATA]  The data to be used as a JSON string - nested array [["1", "2", "3"]]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
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
