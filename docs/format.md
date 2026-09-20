`google-sheet format`
=====================

Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. Only formatting is touched - cell values and formulas are never overwritten.

* [`google-sheet format:cells`](#google-sheet-formatcells)
* [`google-sheet format:merge`](#google-sheet-formatmerge)

## `google-sheet format:cells`

Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. Only formatting is touched - cell values and formulas are never overwritten.

```
USAGE
  $ google-sheet format:cells -s <value> -t <value> [-h] [-r] [-j] [-c <value>] [-p <value>] [-f <value>]
    [--useOauth] [--clientSecretFile <value>] [--range <value>] [--bold] [--italic] [--underline] [--strikethrough]
    [--fontSize <value>] [--fontFamily <value>] [--textColor <value>] [--backgroundColor <value>] [--horizontalAlignment
    LEFT|CENTER|RIGHT] [--verticalAlignment TOP|MIDDLE|BOTTOM] [--wrapStrategy OVERFLOW_CELL|CLIP|WRAP]
    [--numberFormatType TEXT|NUMBER|PERCENT|CURRENCY|DATE|TIME|DATE_TIME|SCIENTIFIC --numberFormat <value>] [--borders
    <value>] [--borderStyle DOTTED|DASHED|SOLID|SOLID_MEDIUM|SOLID_THICK|DOUBLE|NONE] [--borderColor <value>] [--clear]
    [-i <value>] [--dryRun]

FLAGS
  -h, --help                          Show CLI help.
  -i, --input=<value>                 Path to a JSON format spec file (or "-" for stdin)
  -j, --json                          Report failures as a machine-readable JSON envelope on stderr (exit code 1 on
                                      failure). Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput                     Get the raw output as a JSON string
  -s, --spreadsheetId=<value>         (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>        (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --backgroundColor=<value>       Cell background color as #RRGGBB
      --bold                          Bold text
      --borderColor=<value>           [default: #000000] Border color as #RRGGBB (only used with --borders)
      --borderStyle=<option>          [default: SOLID] Border line style (only used with --borders)
                                      <options: DOTTED|DASHED|SOLID|SOLID_MEDIUM|SOLID_THICK|DOUBLE|NONE>
      --borders=<value>               Border sides: comma list of top,bottom,left,right,innerHorizontal,innerVertical or
                                      "all"/"inner"
      --clear                         Reset all formatting on the range (cannot be combined with style flags)
      --dryRun                        Preview the batchUpdate requests without applying them
      --fontFamily=<value>            Font family name (e.g. "Roboto")
      --fontSize=<value>              Font size in points
      --horizontalAlignment=<option>  Horizontal alignment
                                      <options: LEFT|CENTER|RIGHT>
      --italic                        Italic text
      --numberFormat=<value>          Number format pattern (e.g. "#,##0.00", "0.0%", "YYYY-MM-DD")
      --numberFormatType=<option>     Explicit number format type (inferred from the pattern when omitted)
                                      <options: TEXT|NUMBER|PERCENT|CURRENCY|DATE|TIME|DATE_TIME|SCIENTIFIC>
      --range=<value>                 The A1 range to format (e.g. "A1:J1"); required unless --input carries "ranges"
      --strikethrough                 Strikethrough text
      --textColor=<value>             Text color as #RRGGBB
      --underline                     Underline text
      --verticalAlignment=<option>    Vertical alignment
                                      <options: TOP|MIDDLE|BOTTOM>
      --wrapStrategy=<option>         Text wrap strategy
                                      <options: OVERFLOW_CELL|CLIP|WRAP>

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
  formatting is touched - cell values and formulas are never overwritten.

EXAMPLES
  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --bold --backgroundColor="#1a73e8" --textColor="#ffffff"

  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A2:J50 --numberFormat="#,##0.00" --borders=all

  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J50 --clear

  $ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --input=style.json --dryRun
```

_See code: [src/commands/format/cells.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/format/cells.ts)_

## `google-sheet format:merge`

Merge or unmerge cells over a bounded range. Merging keeps the top-left value; other values in the range are hidden by Google.

```
USAGE
  $ google-sheet format:merge -s <value> -t <value> --range <value> [-h] [-r] [-j] [-c <value>] [-p <value>] [-f
    <value>] [--useOauth] [--clientSecretFile <value>] [--type MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS | --unmerge]
    [--dryRun]

FLAGS
  -h, --help                    Show CLI help.
  -j, --json                    Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure).
                                Success output is unchanged - use --rawOutput for JSON success.
  -r, --rawOutput               Get the raw output as a JSON string
  -s, --spreadsheetId=<value>   (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use
  -t, --worksheetTitle=<value>  (required) [env: WORKSHEET_TITLE] Title of the worksheet to use
      --dryRun                  Preview the batchUpdate request without applying it
      --range=<value>           (required) The A1 range to merge or unmerge (e.g. "A1:J1")
      --type=<option>           [default: MERGE_ALL] Merge type
                                <options: MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS>
      --unmerge                 Unmerge previously merged cells in the range

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
  Merge or unmerge cells over a bounded range. Merging keeps the top-left value; other values in the range are hidden by
  Google.

EXAMPLES
  $ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1

  $ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:C3 --type=MERGE_ROWS

  $ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --unmerge
```

_See code: [src/commands/format/merge.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/format/merge.ts)_
