`google-sheet format`
=====================

Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. Only formatting is touched - cell values and formulas are never overwritten. With --workbook the format runs on a local XLSX file instead of Google Sheets, using the subset of flags ExcelJS models (wrapStrategy only as WRAP; --numberFormatType has no local equivalent).

* [`google-sheet format:cells`](#google-sheet-formatcells)
* [`google-sheet format:merge`](#google-sheet-formatmerge)

## `google-sheet format:cells`

Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. Only formatting is touched - cell values and formulas are never overwritten. With --workbook the format runs on a local XLSX file instead of Google Sheets, using the subset of flags ExcelJS models (wrapStrategy only as WRAP; --numberFormatType has no local equivalent).

```
USAGE
  $ google-sheet format:cells [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>] [--useOauth]
    [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>] [-o <value>] [--inPlace]
    [--discardUnsupported] [--range <value>] [--bold] [--italic] [--underline] [--strikethrough] [--fontSize <value>]
    [--fontFamily <value>] [--textColor <value>] [--backgroundColor <value>] [--horizontalAlignment LEFT|CENTER|RIGHT]
    [--verticalAlignment TOP|MIDDLE|BOTTOM] [--wrapStrategy OVERFLOW_CELL|CLIP|WRAP] [--wrapText] [--numberFormatType
    TEXT|NUMBER|PERCENT|CURRENCY|DATE|TIME|DATE_TIME|SCIENTIFIC --numberFormat <value>] [--borders <value>]
    [--borderStyle DOTTED|DASHED|SOLID|SOLID_MEDIUM|SOLID_THICK|DOUBLE|NONE] [--borderColor <value>] [--clear] [-i
    <value>] [--dryRun]

FLAGS
  -h, --help                          Show CLI help.
  -i, --input=<value>                 Path to a JSON format spec file (or "-" for stdin)
  -j, --json                          Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                      failure). Success output is unchanged - use --rawOutput for JSON success.
  -o, --output=<value>                Destination path for the modified XLSX file (without it and without --inPlace the
                                      mutation is refused unless --dryRun)
  -r, --rawOutput                     Get the raw output as a JSON string
  -s, --spreadsheetId=<value>         [env: SPREADSHEET_ID] ID of the spreadsheet to use (Google Sheets target)
  -t, --worksheetTitle=<value>        [env: WORKSHEET_TITLE] Title of the worksheet to use (Google Sheets target; also
                                      selects the sheet inside --workbook)
      --backgroundColor=<value>       Cell background color as #RRGGBB
      --bold                          Bold text
      --borderColor=<value>           [default: #000000] Border color as #RRGGBB (only used with --borders)
      --borderStyle=<option>          [default: SOLID] Border line style (only used with --borders)
                                      <options: DOTTED|DASHED|SOLID|SOLID_MEDIUM|SOLID_THICK|DOUBLE|NONE>
      --borders=<value>               Border sides: comma list of top,bottom,left,right,innerHorizontal,innerVertical or
                                      "all"/"inner"
      --clear                         Reset all formatting on the range (cannot be combined with style flags)
      --discardUnsupported            Allow saving a workbook whose unsupported features (charts, pivot tables, macros)
                                      would be dropped by the local engine
      --dryRun                        Preview the batchUpdate requests without applying them
      --fontFamily=<value>            Font family name (e.g. "Roboto")
      --fontSize=<value>              Font size in points
      --horizontalAlignment=<option>  Horizontal alignment
                                      <options: LEFT|CENTER|RIGHT>
      --inPlace                       Modify the --workbook file in place (a .bak backup is written first)
      --italic                        Italic text
      --numberFormat=<value>          Number format pattern (e.g. "#,##0.00", "0.0%", "YYYY-MM-DD")
      --numberFormatType=<option>     Explicit number format type (inferred from the pattern when omitted)
                                      <options: TEXT|NUMBER|PERCENT|CURRENCY|DATE|TIME|DATE_TIME|SCIENTIFIC>
      --range=<value>                 The A1 range to format (e.g. "A1:J1"); required unless --input carries "ranges"
      --redacted                      [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and
                                      credentials from error envelopes and dry-run diagnostics before they are written.
                                      Coordinates, counts, statuses and outcome states are kept.
      --strikethrough                 Strikethrough text
      --textColor=<value>             Text color as #RRGGBB
      --underline                     Underline text
      --verticalAlignment=<option>    Vertical alignment
                                      <options: TOP|MIDDLE|BOTTOM>
      --workbook=<value>              Path to a local .xlsx workbook to mutate instead of the Google Sheets target. Long
                                      name only: the shared short -f belongs to --credentialsFile
      --wrapStrategy=<option>         Text wrap strategy (locally only WRAP is supported, mapped to wrap text)
                                      <options: OVERFLOW_CELL|CLIP|WRAP>
      --wrapText                      Local backend only: wrap long text onto multiple lines within the cell

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
  Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. Only
  formatting is touched - cell values and formulas are never overwritten. With --workbook the format runs on a local
  XLSX file instead of Google Sheets, using the subset of flags ExcelJS models (wrapStrategy only as WRAP;
  --numberFormatType has no local equivalent).

EXAMPLES
  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --bold --backgroundColor="#1a73e8" --textColor="#ffffff"

  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A2:J50 --numberFormat="#,##0.00" --borders=all

  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J50 --clear

  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --input=style.json --dryRun
```

_See code: [src/commands/format/cells.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/format/cells.ts)_

## `google-sheet format:merge`

Merge or unmerge cells over a bounded range. Merging keeps the top-left value; other values in the range are hidden. With --workbook the merge runs on a local XLSX file instead of Google Sheets.

```
USAGE
  $ google-sheet format:merge --range <value> [-h] [-r] [-j] [--redacted] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [-s <value>] [-t <value>] [--workbook <value>] [-o <value>] [--inPlace]
    [--discardUnsupported] [--type MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS | --unmerge] [--dryRun]

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
      --dryRun                  Preview the mutation without applying it
      --inPlace                 Modify the --workbook file in place (a .bak backup is written first)
      --range=<value>           (required) The A1 range to merge or unmerge (e.g. "A1:J1")
      --redacted                [env: GSHEET_REDACTED] Strip cell contents, formulas, incoming values and credentials
                                from error envelopes and dry-run diagnostics before they are written. Coordinates,
                                counts, statuses and outcome states are kept.
      --type=<option>           [default: MERGE_ALL] Merge type
                                <options: MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS>
      --unmerge                 Unmerge previously merged cells in the range
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
  Merge or unmerge cells over a bounded range. Merging keeps the top-left value; other values in the range are hidden.
  With --workbook the merge runs on a local XLSX file instead of Google Sheets.

EXAMPLES
  $ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1

  $ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:C3 --type=MERGE_ROWS

  $ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --unmerge

  $ gsheet format:merge --workbook=template.xlsx -t "Functional effort" --range=B8:B12 --inPlace
```

_See code: [src/commands/format/merge.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/format/merge.ts)_
