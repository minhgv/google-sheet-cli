`google-sheet report`
=====================

Report automation runner and template generation

* [`google-sheet report:run`](#google-sheet-reportrun)

## `google-sheet report:run`

Generate and publish business reports (finance, manpower, table) from JSON/CSV/XLSX/Sheets sources into XLSX or Google Sheets

```
USAGE
  $ google-sheet report:run --template <value> [-h] [-r] [-i <value>] [--inputFormat json|csv] [--sourceWorkbook
    <value>] [--sourceSpreadsheet <value>] [--ranges <value>] [-o <value>] [-s <value>] [--workbookTemplate <value>]
    [--allowCachedFormulaValues] [--dryRun] [--overwrite] [--overwriteFormulas] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>]

FLAGS
  -h, --help                       Show CLI help.
  -i, --input=<value>              Path to input data file (JSON or CSV) or "-" for stdin
  -o, --output=<value>             Path for output destination XLSX file
  -r, --rawOutput                  Get the raw output as a JSON string
  -s, --spreadsheetId=<value>      ID of target Google Spreadsheet to write report into
      --allowCachedFormulaValues   Explicitly allow using cached formula values from source workbook (required when
                                   source contains formulas)
      --dryRun                     Preview report generation without writing to disk or Google Sheets
      --inputFormat=<option>       Format of input file ("json" or "csv")
                                   <options: json|csv>
      --overwrite                  Allow overwriting existing output XLSX file
      --overwriteFormulas          Allow overwriting existing formula cells in target destination
      --ranges=<value>             JSON array of A1 range strings for source workbook or spreadsheet (e.g.
                                   '["Sheet1!A1:D50"]')
      --sourceSpreadsheet=<value>  ID of source Google Spreadsheet to read data from
      --sourceWorkbook=<value>     Path to source local XLSX file
      --template=<value>           (required) Path to report template JSON file
      --workbookTemplate=<value>   Path to existing XLSX template file to preserve formatting and styles in output file

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
  Generate and publish business reports (finance, manpower, table) from JSON/CSV/XLSX/Sheets sources into XLSX or Google
  Sheets

EXAMPLES
  $ gsheet report:run --template=templates/finance.json --input=data/transactions.csv --output=out/finance-report.xlsx

  $ gsheet report:run --template=templates/manpower.json --sourceWorkbook=data/tasks.xlsx --ranges='["Tasks!A1:H50"]' --allowCachedFormulaValues --output=out/manpower.xlsx --workbookTemplate=templates/styled-base.xlsx

  $ gsheet report:run --template=templates/finance.json --input=data/tx.json --spreadsheetId=<spreadsheetId> --dryRun

  $ gsheet report:run --template=templates/finance.json --sourceSpreadsheet=<sourceId> --ranges='["Sheet1!A1:Z100"]' --spreadsheetId=<targetId> --overwriteFormulas
```

_See code: [src/commands/report/run.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/report/run.ts)_
