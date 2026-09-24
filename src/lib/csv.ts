/**
 * Deterministic CSV serializer for spreadsheet export (T-07, AC-06/AC-07).
 *
 * Pure module: no I/O, no imports, no locale-dependent behavior, no clock. The same
 * `rows` and `options` always produce byte-identical output, so exported files are
 * diffable and reproducible.
 *
 * Wire format (RFC 4180):
 * - Fields containing the delimiter, a double quote, CR or LF are wrapped in double
 *   quotes; internal double quotes are doubled. Nothing else is quoted - in particular
 *   numbers and booleans never carry quotes.
 * - Record separator defaults to CRLF (`\r\n`) per RFC 4180. Every record is terminated,
 *   including the last one (the RFC permits both; emitting the terminator keeps the file
 *   friendly to line-oriented tools and is deterministic either way).
 * - Cell content is byte-faithful: a cell's own LF or CR is never translated to the
 *   configured record separator, only quoted.
 * - Empty string and `null` both serialize as an empty field; there are no invented
 *   headers - the caller passes the header row as the first row if it wants one.
 * - Encoding is UTF-8 without BOM by default; `bom: true` prepends U+FEFF. This function
 *   returns a string - writing bytes (UTF-8) is the file writer's job.
 *
 * Cell rendering:
 * - Values arrive already typed from the read adapter; there is no reformatting. Numbers
 *   and booleans render as `String(value)` with no rounding and no locale grouping
 *   (`true` -> `true`, not `TRUE`). Non-finite numbers (NaN, Infinity) are refused
 *   deterministically: their text would silently change type on re-import. Pass such
 *   text explicitly as a string if it is really intended.
 * - Ragged input is refused, never padded: every row must match the first row's width.
 * - Render mode (raw | formatted | formula text) is the CALLER's choice: this serializer
 *   receives already-rendered cells and treats formula strings as data. Under the `safe`
 *   injection policy those formula strings get the `'` prefix - which is correct, because
 *   formula text in a CSV is exactly the dangerous case (a spreadsheet application will
 *   evaluate `=...` on import). `csvEscapeFormulaText` exists for callers that render
 *   formula text outside `serializeCsv` and want the same treatment.
 *
 * Injection policy (AC-07):
 * - `safe` (default): string cells whose first character is `=`, `+`, `-`, `@`, tab or CR
 *   get a leading `'` (apostrophe) prefix, the spreadsheet-neutralization convention.
 *   The policy applies to STRING type only - typed negative numbers stay numeric and are
 *   never prefixed or quoted.
 * - `preserve`: byte-faithful output; dangerous cells pass through unmodified and the
 *   receipt carries a warning with the count.
 * - Receipts record the selected policy and how many cells the safe policy transformed.
 *   Warnings carry counts only, never cell contents, so receipts are safe for the JSON
 *   channel while data bytes stay in the CSV.
 */

/** One cell of spreadsheet data, already rendered by the read adapter. */
export type CsvCell = string | number | boolean | null;

/** Injection policy for string cells with a formula-dangerous leading character. */
export type CsvInjectionPolicy = 'safe' | 'preserve';

/** All policies, locked order for tests and docs. */
export const CSV_INJECTION_POLICIES = ['safe', 'preserve'] as const;

/** Allowed record separators. RFC 4180 specifies CRLF; LF is accepted for LF-only pipelines. */
export type CsvNewline = '\r\n' | '\n' | '\r';

/** Resolved option defaults. */
export const CSV_DEFAULTS: {
  readonly delimiter: string;
  readonly newline: CsvNewline;
  readonly bom: boolean;
  readonly injectionPolicy: CsvInjectionPolicy;
} = {
  delimiter: ',',
  newline: '\r\n',
  bom: false,
  injectionPolicy: 'safe',
};

/** Serializer options; everything optional, defaults resolved and reported in the receipt. */
export interface CsvOptions {
  /** Single-character field separator. Default `','`. Tab (`'\t'`) is allowed for TSV. */
  delimiter?: string;
  /** Record separator. Default `'\r\n'` per RFC 4180. */
  newline?: CsvNewline;
  /** Prepend a UTF-8 BOM (U+FEFF) for BOM-hungry consumers. Default `false`. */
  bom?: boolean;
  /** Formula-injection policy. Default `'safe'`. */
  injectionPolicy?: CsvInjectionPolicy;
}

/** Machine-readable report of one serialization. Data bytes and this JSON receipt are distinct channels. */
export interface CsvReceipt {
  /** Number of data rows serialized (0 for empty input). */
  rows: number;
  /** Column count; all rows are rectangular, so this is the width of every row. */
  columns: number;
  /** The injection policy actually applied (resolved, not echoed). */
  policy: CsvInjectionPolicy;
  /** Cells the `safe` policy modified by prefixing `'`. Always 0 under `preserve`. */
  transformedCells: number;
  /** Counts and policy notes only - never cell contents. */
  warnings: string[];
}

/** Result of one serialization: the CSV text plus its receipt. Immutable. */
export interface CsvResult {
  readonly csv: string;
  readonly receipt: CsvReceipt;
}

/** First characters that spreadsheet applications may evaluate on CSV import. */
const FORMULA_DANGEROUS_PREFIXES = '=+-@\t\r';

const isFormulaDangerous = (text: string): boolean => text.length > 0 && FORMULA_DANGEROUS_PREFIXES.includes(text[0]);

const throwInvalid = (what: string, value: unknown): never => {
  throw new TypeError(`Invalid CSV ${what}: ${String(value)}`);
};

/** Freeze a result deeply: the returned contract is immutable, callers build on it, never into it. */
const deepFreeze = <T>(value: T): T => {
  if (Array.isArray(value)) {
    value.forEach(deepFreeze);
  } else if (typeof value === 'object' && value !== null) {
    Object.values(value).forEach(deepFreeze);
  }
  return Object.isFrozen(value) ? value : Object.freeze(value);
};

/**
 * Apply the `safe` policy to one already-rendered formula TEXT (e.g. a formula-text
 * render mode export). Serialization applies the identical rule to formula strings it
 * receives as data, so this helper is only needed by callers that escape formula text
 * outside `serializeCsv`. Dangerous text gets the `'` prefix; everything else is
 * returned unchanged.
 */
export function csvEscapeFormulaText(text: string): string {
  if (typeof text !== 'string') throwInvalid('formula text', text);
  return isFormulaDangerous(text) ? `'${text}` : text;
}

interface ResolvedCsvOptions {
  delimiter: string;
  newline: CsvNewline;
  bom: boolean;
  injectionPolicy: CsvInjectionPolicy;
}

const resolveOptions = (options: CsvOptions): ResolvedCsvOptions => {
  const delimiter = options.delimiter === undefined ? CSV_DEFAULTS.delimiter : options.delimiter;
  if (typeof delimiter !== 'string' || delimiter.length !== 1) {
    throwInvalid('delimiter (must be exactly one character)', delimiter);
  }
  if (delimiter === '"' || delimiter === '\r' || delimiter === '\n') {
    throwInvalid('delimiter (cannot be a quote, CR or LF)', delimiter);
  }

  const newline = options.newline === undefined ? CSV_DEFAULTS.newline : options.newline;
  if (newline !== '\r\n' && newline !== '\n' && newline !== '\r') {
    throwInvalid('newline (expected "\\r\\n", "\\n" or "\\r")', newline);
  }

  const bom = options.bom === undefined ? CSV_DEFAULTS.bom : options.bom;
  if (typeof bom !== 'boolean') throwInvalid('bom flag (expected a boolean)', bom);

  const injectionPolicy =
    options.injectionPolicy === undefined ? CSV_DEFAULTS.injectionPolicy : options.injectionPolicy;
  if (!CSV_INJECTION_POLICIES.includes(injectionPolicy)) {
    throwInvalid('injection policy (expected one of: safe, preserve)', injectionPolicy);
  }

  return { delimiter, newline, bom, injectionPolicy };
};

/** Quote per RFC 4180: only when the field contains the delimiter, a quote, CR or LF. */
const encodeField = (text: string, delimiter: string): string => {
  const needsQuote = text.includes(delimiter) || text.includes('"') || text.includes('\r') || text.includes('\n');
  return needsQuote ? `"${text.replace(/"/g, '""')}"` : text;
};

/**
 * Serialize spreadsheet rows to a deterministic RFC 4180 CSV string plus a receipt.
 *
 * Deterministic refusals (TypeError, thrown before any byte is produced): ragged rows,
 * non-finite numbers, unsupported cell types, and invalid options. Refusal beats
 * guessing - nothing is padded, rounded or silently substituted.
 */
export function serializeCsv(rows: readonly CsvCell[][], options: CsvOptions = {}): CsvResult {
  if (!Array.isArray(rows)) throwInvalid('rows (expected an array of arrays)', rows);
  if (options === null || typeof options !== 'object' || Array.isArray(options)) {
    throwInvalid('options (expected an object)', options);
  }
  const { delimiter, newline, bom, injectionPolicy } = resolveOptions(options);

  let width = -1;
  let dangerousCount = 0;
  let transformedCells = 0;

  const lines: string[] = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!Array.isArray(row)) throwInvalid(`row ${r} (expected an array of cells)`, row);
    if (width === -1) {
      width = row.length; // first row anchors the width; every later row must match it
    } else if (row.length !== width) {
      throw new TypeError(
        `Ragged CSV input: row ${r} has ${row.length} cell(s), expected ${width} ` +
          `(every row must match the first row's width; nothing is padded)`
      );
    }

    const fields: string[] = [];
    for (let c = 0; c < row.length; c++) {
      const cell: CsvCell = row[c];
      let text: string;
      if (typeof cell === 'string') {
        if (isFormulaDangerous(cell)) {
          dangerousCount++;
          if (injectionPolicy === 'safe') {
            text = `'${cell}`;
            transformedCells++;
          } else {
            text = cell;
          }
        } else {
          text = cell;
        }
      } else if (typeof cell === 'number') {
        if (!Number.isFinite(cell)) {
          throw new TypeError(
            `Non-finite number at CSV row ${r} column ${c}: ${String(cell)} cannot be serialized ` +
              `unambiguously; pass it as an explicit string if that text is intended`
          );
        }
        text = String(cell);
      } else if (typeof cell === 'boolean') {
        text = String(cell);
      } else if (cell === null) {
        text = '';
      } else {
        throw new TypeError(
          `Unsupported CSV cell at row ${r} column ${c}: expected string | number | boolean | null, ` +
            `got ${cell === undefined ? 'undefined (use null for an empty field)' : typeof cell}`
        );
      }
      fields.push(encodeField(text, delimiter));
    }
    lines.push(fields.join(delimiter));
  }

  const warnings: string[] = [];
  if (injectionPolicy === 'preserve' && dangerousCount > 0) {
    warnings.push(
      `injection-policy=preserve: ${dangerousCount} string cell(s) with a formula-dangerous leading ` +
        `character (=, +, -, @, tab, CR) were emitted unmodified`
    );
  }

  const csv = (bom ? '\uFEFF' : '') + (lines.length > 0 ? `${lines.join(newline)}${newline}` : '');
  const receipt: CsvReceipt = {
    rows: rows.length,
    columns: width === -1 ? 0 : width,
    policy: injectionPolicy,
    transformedCells,
    warnings,
  };
  return deepFreeze({ csv, receipt });
}
