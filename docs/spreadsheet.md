`google-sheet spreadsheet`
==========================

Manage spreadsheets

* [`google-sheet spreadsheet:add`](#google-sheet-spreadsheetadd)
* [`google-sheet spreadsheet:get`](#google-sheet-spreadsheetget)
* [`google-sheet spreadsheet:permissions`](#google-sheet-spreadsheetpermissions)
* [`google-sheet spreadsheet:share`](#google-sheet-spreadsheetshare)
* [`google-sheet spreadsheet:unshare`](#google-sheet-spreadsheetunshare)

## `google-sheet spreadsheet:add`

Add a worksheet with the specified title to the spreadsheet

```
USAGE
  $ google-sheet spreadsheet:add --spreadsheetTitle <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                      Show CLI help.
  -r, --rawOutput                 Get the raw output as a JSON string
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

_See code: [src/commands/spreadsheet/add.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/spreadsheet/add.ts)_

## `google-sheet spreadsheet:get`

Get info for a specific spreadsheet

```
USAGE
  $ google-sheet spreadsheet:get -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>]

FLAGS
  -h, --help                   Show CLI help.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use

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

_See code: [src/commands/spreadsheet/get.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/spreadsheet/get.ts)_

## `google-sheet spreadsheet:permissions`

List the sharing permissions on the spreadsheet (Drive API, drive.file scope).

```
USAGE
  $ google-sheet spreadsheet:permissions -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>]

FLAGS
  -h, --help                   Show CLI help.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use

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

_See code: [src/commands/spreadsheet/permissions.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/spreadsheet/permissions.ts)_

## `google-sheet spreadsheet:share`

Share the spreadsheet with users, a domain, or anyone with the link. Uses the Drive API (drive.file scope) — OAuth tokens issued before that scope was added must re-run `gsheet auth:login`. Under drive.file only files this app created or has opened are visible.

```
USAGE
  $ google-sheet spreadsheet:share -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [--email <value>... | --domain <value> | --anyone] [--type user|group|domain|anyone]
    [--role reader|commenter|writer] [--message <value> --notify]

FLAGS
  -h, --help                   Show CLI help.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --anyone                 Grant access to anyone with the link
      --domain=<value>         Google Workspace domain to grant access
      --email=<value>...       Email address to grant access (repeatable)
      --message=<value>        Message attached to the notification email (requires --notify)
      --notify                 Send Google notification email (default: off — agents should not spam)
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

_See code: [src/commands/spreadsheet/share.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/spreadsheet/share.ts)_

## `google-sheet spreadsheet:unshare`

Remove a sharing permission from the spreadsheet, by permission id or grantee email (Drive API, drive.file scope).

```
USAGE
  $ google-sheet spreadsheet:unshare -s <value> [-h] [-r] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [--permissionId <value> | --email <value>]

FLAGS
  -h, --help                   Show CLI help.
  -r, --rawOutput              Get the raw output as a JSON string
  -s, --spreadsheetId=<value>  (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
      --email=<value>          Grantee email to remove (resolved via the permission list)
      --permissionId=<value>   Permission id from spreadsheet:permissions

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

_See code: [src/commands/spreadsheet/unshare.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/spreadsheet/unshare.ts)_
