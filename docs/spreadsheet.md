`google-sheet spreadsheet`
==========================

Manage spreadsheets

* [`google-sheet spreadsheet:add`](#google-sheet-spreadsheetadd)
* [`google-sheet spreadsheet:copy`](#google-sheet-spreadsheetcopy)
* [`google-sheet spreadsheet:export`](#google-sheet-spreadsheetexport)
* [`google-sheet spreadsheet:get`](#google-sheet-spreadsheetget)
* [`google-sheet spreadsheet:list`](#google-sheet-spreadsheetlist)
* [`google-sheet spreadsheet:permissions`](#google-sheet-spreadsheetpermissions)
* [`google-sheet spreadsheet:share`](#google-sheet-spreadsheetshare)
* [`google-sheet spreadsheet:unshare`](#google-sheet-spreadsheetunshare)

## `google-sheet spreadsheet:add`

Add a worksheet with the specified title to the spreadsheet

```
USAGE
  $ google-sheet spreadsheet:add --spreadsheetTitle <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f
    <value>] [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                      Show CLI help.
  -j, --json                      Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                  failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                 Get the raw output as a JSON string
      --redacted                  [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                  from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                  counts, statuses and outcome states are kept.
      --spreadsheetTitle=<value>  (required) Title of the spreadsheet

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
  $ gsheet worksheet:add --spreadsheetTitle=<spreadsheetTitle>
  Spreadsheet "<spreadsheetTitle>" (<id>) successfully created > https://docs.google.com/spreadsheets/d/<id>/edit
```

_See code: [src/commands/spreadsheet/add.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/add.ts)_

## `google-sheet spreadsheet:copy`

Copy a spreadsheet into a new one through the Drive API (drive.file scope). Under drive.file only files this app created or has opened are visible. The source spreadsheet is left untouched; sharing settings are not duplicated.

```
USAGE
  $ google-sheet spreadsheet:copy -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [--title <value>]

FLAGS
  -h, --help                   Show CLI help.
  -j, --json                   Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                               Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --redacted               [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                               from error envelopes and dry-run diagnostics before they are written. Coordinates,
                               counts, statuses and outcome states are kept.
      --title=<value>          Title of the new spreadsheet (the source title is inherited when omitted)

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
  Copy a spreadsheet into a new one through the Drive API (drive.file scope). Under drive.file only files this app
  created or has opened are visible. The source spreadsheet is left untouched; sharing settings are not duplicated.

EXAMPLES
  $ gsheet spreadsheet:copy --spreadsheetId=<id> --title="Copy of <title>"
  Spreadsheet "<id>" copied to "<title>" (<newId>) > https://docs.google.com/spreadsheets/d/<newId>/edit

  $ gsheet spreadsheet:copy --spreadsheetId=<id>
  Spreadsheet "<id>" copied to "<sourceTitle>" (<newId>) > https://docs.google.com/spreadsheets/d/<newId>/edit
```

_See code: [src/commands/spreadsheet/copy.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/copy.ts)_

## `google-sheet spreadsheet:export`

Export a spreadsheet to a local PDF or XLSX file through the Drive API (drive.file scope). Under drive.file only files this app created or has opened are visible. The Drive API caps exports at 10 MB; larger spreadsheets fail with Google's own error. The output file is written atomically (temporary file in the same directory, then renamed) and an existing file is only replaced with --overwrite.

```
USAGE
  $ google-sheet spreadsheet:export -s <value> --format pdf|xlsx -o <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p
    <value>] [-f <value>] [--useOauth] [--clientSecretFile <value>] [--overwrite]

FLAGS
  -h, --help                   Show CLI help.
  -j, --json                   Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                               Success output is unchanged - use --rawOutput for JSON success.
  -o, --output=<value>         (required) Path of the local file to write
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --format=<option>        (required) Export format
                               <options: pdf|xlsx>
      --overwrite              Replace an existing output file (default: refuse)
      --redacted               [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
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
  Export a spreadsheet to a local PDF or XLSX file through the Drive API (drive.file scope). Under drive.file only files
  this app created or has opened are visible. The Drive API caps exports at 10 MB; larger spreadsheets fail with
  Google's own error. The output file is written atomically (temporary file in the same directory, then renamed) and an
  existing file is only replaced with --overwrite.

EXAMPLES
  $ gsheet spreadsheet:export --spreadsheetId=<id> --format=pdf --output=./report.pdf
  Exported spreadsheet <id> to /abs/path/report.pdf (application/pdf, 12345 bytes)

  $ gsheet spreadsheet:export --spreadsheetId=<id> --format=xlsx --output=./report.xlsx --overwrite
  Exported spreadsheet <id> to /abs/path/report.xlsx (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, 23456 bytes)
```

_See code: [src/commands/spreadsheet/export.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/export.ts)_

## `google-sheet spreadsheet:get`

Get info for a specific spreadsheet

```
USAGE
  $ google-sheet spreadsheet:get -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                   Show CLI help.
  -j, --json                   Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                               Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --redacted               [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
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
  Get info for a specific spreadsheet

EXAMPLES
  $ gsheet spreadsheet:get --spreadsheetId=<spreadsheetId>
  Fetched "<spreadsheetTitle>" (<id>) > https://docs.google.com/spreadsheets/d/<id>/edit
```

_See code: [src/commands/spreadsheet/get.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/get.ts)_

## `google-sheet spreadsheet:list`

List spreadsheets visible to this app via the Drive API (discovery). Uses the drive.file scope: only files this app created or has opened are listed — an empty result does not prove a spreadsheet is absent. Duplicate titles stay separate rows; nothing is ever selected automatically, pick an id yourself. OAuth tokens issued before drive.file was added must re-run `gsheet auth:login`.

```
USAGE
  $ google-sheet spreadsheet:list [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [--exact --name <value>] [--pageSize <value>] [--pageToken <value> | --all]

FLAGS
  -h, --help               Show CLI help.
  -j, --json               Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                           Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput          Get the raw output as a JSON string
      --all                Page through every result, aggregated into one listing (safety bound: 1000 files)
      --exact              Match the whole title exactly instead of a substring (requires --name)
      --name=<value>       Title fragment to match (partial match); without it every visible spreadsheet is listed
      --pageSize=<value>   Files per page (default 50, bounded to 1-100)
      --pageToken=<value>  Continuation token from a previous listing (nextPageToken)
      --redacted           [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials from
                           error envelopes and dry-run diagnostics before they are written. Coordinates, counts,
                           statuses and outcome states are kept.

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
  List spreadsheets visible to this app via the Drive API (discovery). Uses the drive.file scope: only files this app
  created or has opened are listed — an empty result does not prove a spreadsheet is absent. Duplicate titles stay
  separate rows; nothing is ever selected automatically, pick an id yourself. OAuth tokens issued before drive.file was
  added must re-run `gsheet auth:login`.

EXAMPLES
  $ gsheet spreadsheet:list

  $ gsheet spreadsheet:list --name "Report 2026" --rawOutput

  $ gsheet spreadsheet:list --name "Report" --pageSize 100 --all

  $ gsheet spreadsheet:list --name "Report 2026" --exact
```

_See code: [src/commands/spreadsheet/list.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/list.ts)_

## `google-sheet spreadsheet:permissions`

List the sharing permissions on the spreadsheet (Drive API, drive.file scope).

```
USAGE
  $ google-sheet spreadsheet:permissions -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                   Show CLI help.
  -j, --json                   Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                               Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --redacted               [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
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
  List the sharing permissions on the spreadsheet (Drive API, drive.file scope).

EXAMPLES
  $ gsheet spreadsheet:permissions --spreadsheetId=<id>

  $ gsheet spreadsheet:permissions --spreadsheetId=<id> --rawOutput
```

_See code: [src/commands/spreadsheet/permissions.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/permissions.ts)_

## `google-sheet spreadsheet:share`

Share the spreadsheet with users, a domain, or anyone with the link. Uses the Drive API (drive.file scope) — OAuth tokens issued before that scope was added must re-run `gsheet auth:login`. Under drive.file only files this app created or has opened are visible.

```
USAGE
  $ google-sheet spreadsheet:share -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [--email <value>... | --domain <value> | --anyone] [--type
    user|group|domain|anyone] [--role reader|commenter|writer] [--message <value> --notify]

FLAGS
  -h, --help                   Show CLI help.
  -j, --json                   Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                               Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --anyone                 Grant access to anyone with the link
      --domain=<value>         Google Workspace domain to grant access
      --email=<value>...       Email address to grant access (repeatable)
      --message=<value>        Message attached to the notification email (requires --notify)
      --notify                 Send Google notification email (default: off — agents should not spam)
      --redacted               [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                               from error envelopes and dry-run diagnostics before they are written. Coordinates,
                               counts, statuses and outcome states are kept.
      --role=<option>          [default: reader] Access role
                               <options: reader|commenter|writer>
      --type=<option>          Grantee type (inferred from --email/--domain/--anyone when omitted)
                               <options: user|group|domain|anyone>

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
  Share the spreadsheet with users, a domain, or anyone with the link. Uses the Drive API (drive.file scope) — OAuth
  tokens issued before that scope was added must re-run `gsheet auth:login`. Under drive.file only files this app
  created or has opened are visible.

EXAMPLES
  $ gsheet spreadsheet:share --spreadsheetId=<id> --email user@example.com --role writer

  $ gsheet spreadsheet:share --spreadsheetId=<id> --email a@x.com --email b@x.com --role reader

  $ gsheet spreadsheet:share --spreadsheetId=<id> --domain example.com --role commenter

  $ gsheet spreadsheet:share --spreadsheetId=<id> --anyone --role reader
```

_See code: [src/commands/spreadsheet/share.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/share.ts)_

## `google-sheet spreadsheet:unshare`

Remove a sharing permission from the spreadsheet, by permission id or grantee email (Drive API, drive.file scope).

```
USAGE
  $ google-sheet spreadsheet:unshare -s <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [--permissionId <value> | --email <value>]

FLAGS
  -h, --help                   Show CLI help.
  -j, --json                   Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                               Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --email=<value>          Grantee email to remove (resolved via the permission list)
      --permissionId=<value>   Permission id from spreadsheet:permissions
      --redacted               [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
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
  Remove a sharing permission from the spreadsheet, by permission id or grantee email (Drive API, drive.file scope).

EXAMPLES
  $ gsheet spreadsheet:unshare --spreadsheetId=<id> --permissionId <id>

  $ gsheet spreadsheet:unshare --spreadsheetId=<id> --email user@example.com
```

_See code: [src/commands/spreadsheet/unshare.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/spreadsheet/unshare.ts)_
