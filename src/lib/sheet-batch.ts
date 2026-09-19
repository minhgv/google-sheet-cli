import { sheets_v4 } from '@googleapis/sheets';
import { GoogleSheetCli } from './google-sheet';
import type { ReportCell, ReportDocument, ReportFormula, ReportNumberFormat, ReportSheet } from './report/types';

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
export const DEFAULT_CHUNK_BYTE_SIZE = 2 * 1024 * 1024; // 2MB safe boundary
export const DEFAULT_BATCH_GET_CHUNK_SIZE = 50; // 50 ranges per batchGet request

/**
 * Convert 1-based column number to A1 notation column letters (e.g. 1 -> A, 27 -> AA)
 */
export const colToA1 = (col: number): string => {
  if (col < 1) throw new Error('col has to be greater than 0');
  let div = col;
  let label = '';
  while (div > 0) {
    const mod = (div - 1) % ALPHABET.length;
    label = `${ALPHABET[mod]}${label}`;
    div = Math.floor((div - mod - 1) / ALPHABET.length);
  }
  return label;
};

/**
 * Convert A1 notation column letters to 1-based column number (e.g. A -> 1, AA -> 27)
 */
export const a1ToCol = (label: string): number => {
  const upper = label.toUpperCase();
  if (!/^[A-Z]+$/.test(upper)) {
    throw new Error(`Invalid column label "${label}"`);
  }
  let col = 0;
  for (let i = 0; i < upper.length; i++) {
    col = col * ALPHABET.length + (upper.charCodeAt(i) - 64);
  }
  return col;
};

/**
 * Parse a single A1 cell reference (e.g. "B3", "$C$12")
 */
export const parseA1Cell = (cellRef: string): { col: number; row: number } => {
  const match = cellRef.trim().match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!match) throw new Error(`Invalid cell reference "${cellRef}"`);
  return {
    col: a1ToCol(match[1]),
    row: parseInt(match[2], 10),
  };
};

/**
 * Format a cell coordinate into A1 string (e.g. col 2, row 3 -> "B3")
 */
export const formatA1Cell = (col: number, row: number): string => {
  return `${colToA1(col)}${row}`;
};

/**
 * Check if a worksheet title needs single-quote wrapping for A1 notation
 */
export const needsQuoting = (title: string): boolean => {
  if (!title) return false;
  // If already single-quoted
  if (title.startsWith("'") && title.endsWith("'")) {
    return false;
  }
  // Title starting with digit or non-letter/underscore, or containing non-alphanumeric/underscore chars or spaces
  return /^[^A-Za-z_]|[^A-Za-z0-9_]/.test(title);
};

/**
 * Format worksheet title safely with quotes if needed
 */
export const escapeWorksheetTitle = (title: string): string => {
  if (!title) return '';
  let unquoted = title;
  if (unquoted.startsWith('"') && unquoted.endsWith('"') && unquoted.length >= 2) {
    unquoted = unquoted.slice(1, -1);
  }
  if (unquoted.startsWith("'") && unquoted.endsWith("'") && unquoted.length >= 2) {
    return unquoted;
  }
  if (needsQuoting(unquoted)) {
    return `'${unquoted.replace(/'/g, "''")}'`;
  }
  return unquoted;
};

export interface ParsedA1Range {
  worksheetTitle?: string;
  startCol?: number;
  startRow?: number;
  endCol?: number;
  endRow?: number;
  isBounded: boolean;
  rawRange: string;
}

/**
 * Strict A1 range parser supporting formats:
 * - 'Sheet 1'!A1:B10
 * - Sheet1!A1:B10
 * - Sheet1!A1
 * - A1:B10
 * - A1
 * - Sheet1!A:B (unbounded rows)
 * - Sheet1!1:10 (unbounded cols)
 */
export const parseStrictA1Range = (rangeStr: string): ParsedA1Range => {
  if (typeof rangeStr !== 'string' || !rangeStr.trim()) {
    throw new Error(`Invalid range "${rangeStr}": range must be a non-empty string`);
  }

  const trimmed = rangeStr.trim();
  let worksheetTitle: string | undefined;
  let a1Part = trimmed;

  // Split worksheet title if present
  if (trimmed.startsWith("'")) {
    let idx = 1;
    let title = '';
    while (idx < trimmed.length) {
      if (trimmed[idx] === "'") {
        if (trimmed[idx + 1] === "'") {
          title += "'";
          idx += 2;
          continue;
        }
        break;
      }
      title += trimmed[idx];
      idx++;
    }
    if (idx >= trimmed.length || trimmed[idx] !== "'" || trimmed[idx + 1] !== '!') {
      throw new Error(`Invalid range "${rangeStr}": malformed quoted worksheet title`);
    }
    worksheetTitle = title;
    a1Part = trimmed.slice(idx + 2);
  } else if (trimmed.startsWith('"')) {
    throw new Error(`Invalid range "${rangeStr}": double-quoted worksheet titles are not valid A1 notation; use single quotes`);
  } else {
    const bangIdx = trimmed.lastIndexOf('!');
    if (bangIdx >= 0) {
      worksheetTitle = trimmed.slice(0, bangIdx);
      a1Part = trimmed.slice(bangIdx + 1);
    }
  }

  if (!a1Part) {
    if (worksheetTitle) {
      return { worksheetTitle, isBounded: false, rawRange: trimmed };
    }
    throw new Error(`Invalid range "${rangeStr}": empty cell coordinates`);
  }

  // Parse a1Part: e.g. "A1:B10", "A1", "A:B", "1:10"
  const colonIdx = a1Part.indexOf(':');
  if (colonIdx < 0) {
    // Single cell or column or row
    const cellMatch = a1Part.match(/^\$?([A-Za-z]+)?\$?(\d+)?$/);
    if (!cellMatch || (!cellMatch[1] && !cellMatch[2])) {
      throw new Error(`Invalid range "${rangeStr}": unparseable cell "${a1Part}"`);
    }
    const col = cellMatch[1] ? a1ToCol(cellMatch[1]) : undefined;
    const row = cellMatch[2] ? parseInt(cellMatch[2], 10) : undefined;
    return {
      worksheetTitle,
      startCol: col,
      startRow: row,
      endCol: col,
      endRow: row,
      isBounded: false,
      rawRange: trimmed,
    };
  }

  const startPart = a1Part.slice(0, colonIdx);
  const endPart = a1Part.slice(colonIdx + 1);

  const startMatch = startPart.match(/^\$?([A-Za-z]+)?\$?(\d+)?$/);
  const endMatch = endPart.match(/^\$?([A-Za-z]+)?\$?(\d+)?$/);

  if (!startMatch || !endMatch || (!startMatch[1] && !startMatch[2]) || (!endMatch[1] && !endMatch[2])) {
    throw new Error(`Invalid range "${rangeStr}": unparseable bounding parts`);
  }

  const startCol = startMatch[1] ? a1ToCol(startMatch[1]) : undefined;
  const startRow = startMatch[2] ? parseInt(startMatch[2], 10) : undefined;
  const endCol = endMatch[1] ? a1ToCol(endMatch[1]) : undefined;
  const endRow = endMatch[2] ? parseInt(endMatch[2], 10) : undefined;

  return {
    worksheetTitle,
    startCol,
    startRow,
    endCol: endCol ?? startCol,
    endRow: endRow ?? startRow,
    isBounded: true,
    rawRange: trimmed,
  };
};

/**
 * Format a bounding box back to strict A1 range
 */
export const formatBoundedA1Range = (
  worksheetTitle: string | undefined,
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number
): string => {
  const prefix = worksheetTitle ? `${escapeWorksheetTitle(worksheetTitle)}!` : '';
  const start = `${colToA1(startCol)}${startRow}`;
  const end = `${colToA1(endCol)}${endRow}`;
  return start === end ? `${prefix}${start}` : `${prefix}${start}:${end}`;
};

/**
 * Rough estimation of JSON byte size for a value range update
 */
export const estimateUpdateBytes = (range: string, values: GoogleSheetCli.RawData): number => {
  let bytes = range.length + 32; // envelope
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    if (!row) continue;
    bytes += 8; // array brackets & comma
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell === null || cell === undefined) {
        bytes += 4; // null
      } else if (typeof cell === 'string') {
        bytes += cell.length + 4; // quotes + escapes
      } else if (typeof cell === 'number') {
        bytes += 16;
      } else if (typeof cell === 'boolean') {
        bytes += 5;
      } else {
        bytes += 32;
      }
    }
  }
  return bytes;
};

/**
 * Split an update that spans many rows into multiple smaller row-chunked updates
 * ensuring correct A1 notation coordinates and worksheet quoting.
 */
export const splitUpdateByRows = (
  update: { range: string; values: GoogleSheetCli.RawData },
  maxBytesPerChunk: number = DEFAULT_CHUNK_BYTE_SIZE,
  maxRowsPerChunk: number = 5000
): { range: string; values: GoogleSheetCli.RawData }[] => {
  const parsed = parseStrictA1Range(update.range);
  const startCol = parsed.startCol || 1;
  const startRow = parsed.startRow || 1;
  const values = update.values;

  if (values.length <= 1) {
    return [update];
  }

  const longestRow = values.reduce((max, row) => Math.max(max, row?.length || 0), 0);
  const endCol = parsed.endCol || (startCol + longestRow - 1);

  // If already small enough, return as is
  const totalBytes = estimateUpdateBytes(update.range, values);
  if (totalBytes <= maxBytesPerChunk && values.length <= maxRowsPerChunk) {
    return [update];
  }

  // Calculate safe row batch size
  const bytesPerRow = Math.max(1, Math.ceil(totalBytes / values.length));
  const calculatedRowChunk = Math.max(1, Math.min(maxRowsPerChunk, Math.floor((maxBytesPerChunk * 0.8) / bytesPerRow)));

  const chunks: { range: string; values: GoogleSheetCli.RawData }[] = [];
  for (let offset = 0; offset < values.length; offset += calculatedRowChunk) {
    const slice = values.slice(offset, offset + calculatedRowChunk);
    const chunkStartRow = startRow + offset;
    const chunkEndRow = chunkStartRow + slice.length - 1;
    const chunkRange = formatBoundedA1Range(parsed.worksheetTitle, startCol, chunkStartRow, endCol, chunkEndRow);
    chunks.push({
      range: chunkRange,
      values: slice,
    });
  }

  return chunks;
};

/**
 * Pack updates into batches such that each batch payload remains within maxBytesPerChunk
 */
export const packUpdateBatches = (
  updates: { range: string; values: GoogleSheetCli.RawData }[],
  maxBytesPerChunk: number = DEFAULT_CHUNK_BYTE_SIZE
): { range: string; values: GoogleSheetCli.RawData }[][] => {
  const flatSplitUpdates: { range: string; values: GoogleSheetCli.RawData }[] = [];
  for (const upd of updates) {
    flatSplitUpdates.push(...splitUpdateByRows(upd, maxBytesPerChunk));
  }

  const batches: { range: string; values: GoogleSheetCli.RawData }[][] = [];
  let currentBatch: { range: string; values: GoogleSheetCli.RawData }[] = [];
  let currentBatchBytes = 0;

  for (const item of flatSplitUpdates) {
    const itemBytes = estimateUpdateBytes(item.range, item.values);
    if (currentBatch.length > 0 && currentBatchBytes + itemBytes > maxBytesPerChunk) {
      batches.push(currentBatch);
      currentBatch = [];
      currentBatchBytes = 0;
    }
    currentBatch.push(item);
    currentBatchBytes += itemBytes;
  }

  if (currentBatch.length > 0) {
    batches.push(currentBatch);
  }

  return batches;
};

/**
 * Check if a cell is an intentional formula or starts with '='
 */
export const isFormula = (val: unknown): boolean => {
  if (typeof val === 'string' && val.startsWith('=')) return true;
  if (typeof val === 'object' && val !== null && 'formula' in val) return true;
  return false;
};

/**
 * Extract formula string normalized with leading '='
 */
export const extractFormulaText = (val: unknown): string | null => {
  if (typeof val === 'string' && val.startsWith('=')) return val;
  if (typeof val === 'object' && val !== null && 'formula' in val) {
    const f = (val as ReportFormula).formula;
    return f.startsWith('=') ? f : `=${f}`;
  }
  return null;
};

export interface FormulaConflict {
  cell: string;
  existingFormula: string;
  incomingValue: unknown;
  updateIndex?: number;
}

/**
 * Inspect target ranges against existing formula data to find any unauthorized formula overwrites
 */
export const findFormulaOverwrites = (
  updates: { range: string; values: GoogleSheetCli.RawData }[],
  existingData: { range: string; values: GoogleSheetCli.RawData }[]
): FormulaConflict[] => {
  const conflicts: FormulaConflict[] = [];

  // Map existing formulas by "Sheet!Col:Row"
  const existingFormulas = new Map<string, { formula: string; sheetTitle?: string; col: number; row: number }>();

  for (const existing of existingData) {
    if (!existing.values || existing.values.length === 0) continue;
    const parsed = parseStrictA1Range(existing.range);
    const startCol = parsed.startCol || 1;
    const startRow = parsed.startRow || 1;

    for (let r = 0; r < existing.values.length; r++) {
      const row = existing.values[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        const formula = extractFormulaText(cell);
        if (formula) {
          const colNum = startCol + c;
          const rowNum = startRow + r;
          const key = `${(parsed.worksheetTitle || '').toLowerCase()}!${colNum}:${rowNum}`;
          existingFormulas.set(key, {
            formula,
            sheetTitle: parsed.worksheetTitle,
            col: colNum,
            row: rowNum,
          });
        }
      }
    }
  }

  if (existingFormulas.size === 0) {
    return conflicts;
  }

  // Check incoming updates against existing formulas
  for (let uIdx = 0; uIdx < updates.length; uIdx++) {
    const upd = updates[uIdx];
    const parsed = parseStrictA1Range(upd.range);
    const startCol = parsed.startCol || 1;
    const startRow = parsed.startRow || 1;

    for (let r = 0; r < upd.values.length; r++) {
      const row = upd.values[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const incoming = row[c];
        const colNum = startCol + c;
        const rowNum = startRow + r;
        const key = `${(parsed.worksheetTitle || '').toLowerCase()}!${colNum}:${rowNum}`;

        const existing = existingFormulas.get(key);
        if (existing) {
          const incomingFormula = extractFormulaText(incoming);
          // If incoming is NOT the identical formula (e.g. overwriting formula with a number/string/null/blank/different formula)
          if (incomingFormula !== existing.formula) {
            const cellAddr = formatBoundedA1Range(existing.sheetTitle, colNum, rowNum, colNum, rowNum);
            conflicts.push({
              cell: cellAddr,
              existingFormula: existing.formula,
              incomingValue: incoming,
              updateIndex: uIdx,
            });
          }
        }
      }
    }
  }

  return conflicts;
};

/**
 * Convert a ReportCell to Google Sheets API Schema$ExtendedValue
 */
export const toGoogleExtendedValue = (cell: ReportCell): sheets_v4.Schema$ExtendedValue => {
  if (cell === null || cell === undefined || cell === '') {
    return {};
  }
  if (typeof cell === 'boolean') {
    return { boolValue: cell };
  }
  if (typeof cell === 'number') {
    return { numberValue: cell };
  }
  if (typeof cell === 'string') {
    // Literal string (even if it starts with '='), preserves literal text without formula evaluation
    return { stringValue: cell };
  }
  if (typeof cell === 'object' && 'formula' in cell) {
    const formulaStr = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`;
    return { formulaValue: formulaStr };
  }
  return { stringValue: String(cell) };
};

/**
 * Convert number format pattern to Google Sheets API Schema$NumberFormat
 */
export const toGoogleNumberFormat = (formatStr: string): sheets_v4.Schema$NumberFormat => {
  const upper = formatStr.toUpperCase();
  if (upper.includes('%')) {
    return { type: 'PERCENT', pattern: formatStr };
  }
  if (upper.includes('$') || upper.includes('₫') || upper.includes('€') || upper.includes('£')) {
    return { type: 'CURRENCY', pattern: formatStr };
  }
  if (upper.includes('YYYY') || upper.includes('YY') || upper.includes('MM') || upper.includes('DD')) {
    return { type: 'DATE', pattern: formatStr };
  }
  if (upper.includes('#') || upper.includes('0')) {
    return { type: 'NUMBER', pattern: formatStr };
  }
  return { type: 'TEXT', pattern: formatStr };
};

export interface SheetManagedMetadata {
  templateId: string;
  templateVersion?: string | number;
  sourceHash?: string;
  startRow: number;
  startCol: number;
  rowCount: number;
  colCount: number;
  updatedAt: string;
}

export const METADATA_KEY_REPORT_MANAGED = 'gsheet_report_managed';
