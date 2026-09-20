import { GoogleSheetCli } from './google-sheet';
import { colToA1, formatA1Cell, formatBoundedA1Range } from './sheet-batch';
import { ValidationError } from './validation-error';

/**
 * Pure planning for the key-based upsert: frame the existing table, validate the input
 * against it, and compute the deterministic fixed-range write plan. No I/O happens in this
 * module — every rejection it can produce fires before GoogleSheet.upsert asks the transport
 * for a single write, which is what keeps a refused upsert from leaving partial updates
 * behind.
 */

type Scalar = string | number | boolean | null;

/**
 * The existing table as read from the sheet: header row at startRow, data rows below it, all
 * coordinates absolute against the worksheet.
 */
export interface TableFrame {
  worksheetTitle: string;
  /** header row (1-based) */
  startRow: number;
  /** first table column (1-based) */
  startCol: number;
  header: string[];
  width: number;
  /** data rows, each normalized to exactly `width` cells */
  rows: GoogleSheetCli.RawData;
  /** index of the key column inside `header` */
  keyIndex: number;
  /** typed scalar key -> data row index */
  keyToRow: Map<string, number>;
}

/**
 * The input matrix mapped onto the table: which input column feeds which table column, which
 * input column carries the key, and the span of table columns the input supplies.
 */
export interface UpsertInput {
  /** data rows below the input header row, cells normalized to scalars */
  rows: GoogleSheetCli.RawData;
  /** input column index -> table column index, in input header order */
  colMap: Map<number, number>;
  /** index of the key column inside the input header */
  keyInputColumn: number;
  /** min table column index the input supplies */
  spanStart: number;
  /** max table column index the input supplies */
  spanEnd: number;
}

export interface UpsertPlan {
  /** write ranges in execution order: one entry per input row, in input order */
  updates: { range: string; values: GoogleSheetCli.RawData }[];
  addedRows: number;
  updatedRows: number;
  unchangedRows: number;
  updatedRanges: string[];
  addedRanges: string[];
  /** every write range, in execution order */
  plannedRanges: string[];
}

/**
 * Human-readable rendering of a cell value for error messages.
 */
const describeValue = (value: unknown): string => {
  if (value === null) return 'null';
  if (typeof value === 'string') return `"${value}"`;
  if (typeof value === 'object') return Array.isArray(value) ? 'an array' : 'an object';
  return String(value);
};

/**
 * Normalize one cell to a scalar. Null and undefined collapse to the empty string, which is
 * what an unset cell reads back as; anything non-scalar is refused rather than coerced.
 */
export const normalizeCell = (value: unknown): Scalar => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  throw new ValidationError(`Upsert cells must be scalars (string, number, boolean or null); got ${describeValue(value)}`);
};

/**
 * Type-aware cell equality for change detection. A number never equals the string spelling of
 * itself: under the RAW default an upsert must not silently retype cells, so `7` and `"7"`
 * count as a change, while unset, null and "" all count as the same empty cell.
 */
export const cellsEqual = (existing: unknown, incoming: unknown): boolean => {
  const left = normalizeCell(existing);
  const right = normalizeCell(incoming);
  if (typeof left !== typeof right) return false;
  return left === right;
};

/**
 * Type-aware key token. `"001"` stays a string token distinct from the number `1`, so
 * leading-zero identifiers keep matching string cells and never collapse into numbers.
 * Returns null for an empty key, which callers refuse.
 */
export const scalarKey = (value: unknown): string | null => {
  const cell = normalizeCell(value);
  if (cell === '') return null;
  if (typeof cell === 'number') return `n:${cell}`;
  if (typeof cell === 'boolean') return `b:${cell}`;
  return `s:${cell}`;
};

/**
 * Frame the existing table from a raw values read and validate it: contiguous non-empty
 * header, unique headers, rectangular data, and keys that exist and are unique. The read is
 * expected to be clamped and trailing-empty-trimmed the way values.batchGet trims, so the
 * header width comes from the last non-empty header cell.
 *
 * @param {GoogleSheetCli.RawData} [values] - raw grid rows starting at the header row
 * @param {object} options - worksheetTitle/startRow/startCol locate the table, keyColumn names the key
 * @returns {TableFrame}
 */
export const buildTableFrame = (
  values: GoogleSheetCli.RawData | undefined,
  options: { worksheetTitle: string; startRow: number; startCol: number; keyColumn: string }
): TableFrame => {
  const grid = values || [];
  const headerRow = grid.length > 0 ? grid[0] || [] : [];

  // A grid-wide read pads out to the sheet's whole width with empty cells; the table is only
  // as wide as its header row, so trailing empty header cells are trimmed, not rejected.
  let width = headerRow.length;
  while (width > 0 && normalizeCell(headerRow[width - 1]) === '') width--;
  if (width === 0) {
    throw new ValidationError(
      `No table header found at ${formatA1Cell(options.startCol, options.startRow)} in "${options.worksheetTitle}"; upsert requires an existing table with a header row`
    );
  }

  const header: string[] = [];
  for (let c = 0; c < width; c++) {
    const cell = normalizeCell(headerRow[c]);
    if (cell === '') {
      throw new ValidationError(
        `Empty header at ${formatA1Cell(options.startCol + c, options.startRow)} in "${options.worksheetTitle}"; upsert requires an unambiguous contiguous header row`
      );
    }
    const name = String(cell);
    const firstAt = header.indexOf(name);
    if (firstAt >= 0) {
      throw new ValidationError(
        `Duplicate header "${name}" in "${options.worksheetTitle}" (columns ${colToA1(options.startCol + firstAt)} and ${colToA1(options.startCol + c)}); upsert requires unambiguous headers`
      );
    }
    header.push(name);
  }

  const keyIndex = header.indexOf(options.keyColumn);
  if (keyIndex < 0) {
    throw new ValidationError(`Key column "${options.keyColumn}" not found in the existing table header [${header.join(', ')}]`);
  }

  const rows: GoogleSheetCli.RawData = [];
  for (let r = 1; r < grid.length; r++) {
    const raw = grid[r] || [];
    // Data outside the header width would be silently unreachable through the header map,
    // which would make upsert lie about preservation; refuse the ragged table instead.
    for (let c = width; c < raw.length; c++) {
      if (normalizeCell(raw[c]) !== '') {
        throw new ValidationError(
          `Ragged table: ${formatA1Cell(options.startCol + c, options.startRow + r)} holds data outside the ${width}-column header of "${options.worksheetTitle}"`
        );
      }
    }
    rows.push(Array.from({ length: width }, (_, c) => normalizeCell(raw[c])));
  }

  const keyToRow = new Map<string, number>();
  for (let r = 0; r < rows.length; r++) {
    const key = scalarKey(rows[r][keyIndex]);
    const rowNumber = options.startRow + 1 + r;
    if (key === null) {
      throw new ValidationError(
        `Empty key in "${options.worksheetTitle}" at row ${rowNumber} (key column "${options.keyColumn}"); upsert cannot match rows without a key`
      );
    }
    const previous = keyToRow.get(key);
    if (previous !== undefined) {
      throw new ValidationError(
        `Duplicate key ${describeValue(rows[r][keyIndex])} in "${options.worksheetTitle}" at rows ${options.startRow + 1 + previous} and ${rowNumber}; fix the table before upserting`
      );
    }
    keyToRow.set(key, r);
  }

  return {
    worksheetTitle: options.worksheetTitle,
    startRow: options.startRow,
    startCol: options.startCol,
    header,
    width,
    rows,
    keyIndex,
    keyToRow,
  };
};

/**
 * Validate the input matrix against a framed table and map input columns onto table columns.
 * Refuses empty or duplicate input headers, headers the table does not have, a missing key
 * column, and empty or duplicate input keys.
 *
 * @param {GoogleSheetCli.RawData} rows - header-first input matrix
 * @param {TableFrame} frame - the framed existing table
 * @param {string} keyColumn - header name of the key column, must be supplied by the input
 * @returns {UpsertInput}
 */
export const parseUpsertInput = (rows: GoogleSheetCli.RawData, frame: TableFrame, keyColumn: string): UpsertInput => {
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new ValidationError('Upsert input must be a non-empty 2D array with a header first row');
  }

  const inputHeaderRow = rows[0] || [];
  const inputHeader: string[] = [];
  for (let c = 0; c < inputHeaderRow.length; c++) {
    const cell = normalizeCell(inputHeaderRow[c]);
    if (cell === '') {
      throw new ValidationError(`Empty header in the input at position ${c + 1}; upsert requires named columns`);
    }
    const name = String(cell);
    if (inputHeader.includes(name)) {
      throw new ValidationError(`Duplicate header "${name}" in the input; upsert requires unambiguous columns`);
    }
    inputHeader.push(name);
  }

  const colMap = new Map<number, number>();
  for (let c = 0; c < inputHeader.length; c++) {
    const tableColumn = frame.header.indexOf(inputHeader[c]);
    if (tableColumn < 0) {
      throw new ValidationError(
        `Unknown input column "${inputHeader[c]}": the existing table header is [${frame.header.join(', ')}]; upsert cannot place new columns`
      );
    }
    colMap.set(c, tableColumn);
  }

  const keyInputColumn = inputHeader.indexOf(keyColumn);
  if (keyInputColumn < 0) {
    throw new ValidationError(`The input does not supply the key column "${keyColumn}"`);
  }

  const inputRows: GoogleSheetCli.RawData = rows.slice(1).map((row, r) => {
    if (!Array.isArray(row)) {
      throw new ValidationError(`Upsert input row ${r + 2} is not an array`);
    }
    return row.map((cell) => normalizeCell(cell));
  });

  const keys = new Map<string, number>();
  for (let r = 0; r < inputRows.length; r++) {
    const key = scalarKey(inputRows[r][keyInputColumn]);
    const inputRowNumber = r + 2; // the header is row 1 of the input
    if (key === null) {
      throw new ValidationError(
        `Empty key in the input at row ${inputRowNumber} (key column "${keyColumn}"); upsert cannot match rows without a key`
      );
    }
    const previous = keys.get(key);
    if (previous !== undefined) {
      throw new ValidationError(
        `Duplicate key ${describeValue(inputRows[r][keyInputColumn])} in the input at rows ${previous + 2} and ${inputRowNumber}; keys must be unique within one upsert`
      );
    }
    keys.set(key, r);
  }

  const supplied = [...colMap.values()];
  return {
    rows: inputRows,
    colMap,
    keyInputColumn,
    spanStart: Math.min(...supplied),
    spanEnd: Math.max(...supplied),
  };
};

/**
 * Compute the write plan for one upsert run. Matched rows are written per changed cell so
 * untouched cells - formulas included - are never part of any write range; new keys append as
 * one contiguous row-span write below the last existing data row. The plan is deterministic:
 * the same input against the same table always produces the same ranges in the same order.
 *
 * @param {TableFrame} frame - the framed existing table
 * @param {UpsertInput} input - validated, column-mapped input
 * @returns {UpsertPlan}
 */
export const planUpsert = (frame: TableFrame, input: UpsertInput): UpsertPlan => {
  const updates: { range: string; values: GoogleSheetCli.RawData }[] = [];
  const updatedRanges: string[] = [];
  const addedRanges: string[] = [];
  let addedRows = 0;
  let updatedRows = 0;
  let unchangedRows = 0;

  for (const row of input.rows) {
    const key = scalarKey(row[input.keyInputColumn]);
    if (key === null) {
      // Unreachable through parseUpsertInput, which refuses empty keys first; guarded so the
      // planner stays correct on its own.
      throw new ValidationError('Upsert plan reached an empty key');
    }
    const existingIndex = frame.keyToRow.get(key);

    if (existingIndex === undefined) {
      // Appended row: the read trimmed trailing empty rows, so this row is empty and a
      // contiguous write across the supplied span cannot clobber anything.
      const rowNumber = frame.startRow + 1 + frame.rows.length + addedRows;
      const span = Array.from({ length: input.spanEnd - input.spanStart + 1 }, () => '' as Scalar);
      for (const [inputColumn, tableColumn] of input.colMap) {
        span[tableColumn - input.spanStart] = normalizeCell(row[inputColumn]);
      }
      const range = formatBoundedA1Range(
        frame.worksheetTitle,
        frame.startCol + input.spanStart,
        rowNumber,
        frame.startCol + input.spanEnd,
        rowNumber
      );
      updates.push({ range, values: [span] });
      addedRanges.push(range);
      addedRows++;
      continue;
    }

    let changed = false;
    for (const [inputColumn, tableColumn] of input.colMap) {
      // A short input row simply did not supply the trailing cells; never-sent cells are not
      // write targets. Only cells the input actually carries can change anything.
      if (inputColumn >= row.length) continue;
      const incoming = normalizeCell(row[inputColumn]);
      const existing = frame.rows[existingIndex][tableColumn];
      if (cellsEqual(existing, incoming)) continue;
      const rowNumber = frame.startRow + 1 + existingIndex;
      const column = frame.startCol + tableColumn;
      const range = formatBoundedA1Range(frame.worksheetTitle, column, rowNumber, column, rowNumber);
      updates.push({ range, values: [[incoming]] });
      updatedRanges.push(range);
      changed = true;
    }
    if (changed) {
      updatedRows++;
    } else {
      unchangedRows++;
    }
  }

  return {
    updates,
    addedRows,
    updatedRows,
    unchangedRows,
    updatedRanges,
    addedRanges,
    // Execution order, not grouped by kind: updates iterates input rows in order, so a mixed
    // run interleaves changed-cell writes with appended rows exactly as they will be sent.
    plannedRanges: updates.map((update) => update.range),
  };
};

/**
 * The A1 address of the first cell with content in a below-bound read, or null when the
 * region is empty. The read must be rendered as formulas: that way a cell refuses the upsert
 * both when it holds data and when it holds a formula - even one evaluating to empty, which
 * a plain values read would report as an empty cell and wave through. Only the address is
 * reported, never the cell's value, so refusal errors stay free of spreadsheet contents.
 *
 * @param {GoogleSheetCli.RawData} [values] - grid rows read starting at startRow
 * @param {object} options - startRow/startCol locate the read region's top-left corner
 * @returns {string | null}
 */
export const firstBelowBoundCell = (
  values: GoogleSheetCli.RawData | undefined,
  options: { startRow: number; startCol: number }
): string | null => {
  const grid = values || [];
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell !== null && cell !== undefined && cell !== '') {
        return formatA1Cell(options.startCol + c, options.startRow + r);
      }
    }
  }
  return null;
};
