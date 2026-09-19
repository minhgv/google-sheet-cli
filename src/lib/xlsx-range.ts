import { ReportSheet } from './xlsx-types';

export const MAX_EXCEL_ROWS = 1048576;
export const MAX_EXCEL_COLS = 16384; // XFD

export interface ParsedA1Range {
  raw: string;
  sheetName?: string;
  startCol: number; // 1-based index
  startRow: number; // 1-based index
  endCol: number; // 1-based index
  endRow: number; // 1-based index
  isSingleCell: boolean;
  isFullRow: boolean;
  isFullCol: boolean;
}

/**
 * Converts a 1-based column index to Excel column letters (e.g. 1 -> "A", 27 -> "AA", 16384 -> "XFD").
 */
export function indexToColLetter(colIdx: number): string {
  if (!Number.isInteger(colIdx) || colIdx < 1 || colIdx > MAX_EXCEL_COLS) {
    throw new Error(
      `Column index ${colIdx} is out of bounds (must be between 1 and ${MAX_EXCEL_COLS}).`
    );
  }

  let temp = colIdx;
  let letter = '';
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

/**
 * Converts Excel column letters to a 1-based column index (e.g. "A" -> 1, "AA" -> 27, "XFD" -> 16384).
 */
export function colLetterToIndex(colStr: string): number {
  if (!colStr || typeof colStr !== 'string') {
    throw new Error(`Invalid column string: "${colStr}".`);
  }

  const upper = colStr.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(upper)) {
    throw new Error(`Invalid column letters: "${colStr}".`);
  }

  let index = 0;
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }

  if (index < 1 || index > MAX_EXCEL_COLS) {
    throw new Error(
      `Column "${colStr}" (index ${index}) exceeds maximum Excel column limit (${MAX_EXCEL_COLS}).`
    );
  }

  return index;
}

/**
 * Converts row and column (1-based) to an A1 cell address (e.g. 1, 1 -> "A1").
 */
export function cellAddress(row: number, col: number): string {
  return `${indexToColLetter(col)}${row}`;
}

/**
 * Parses a single cell address like "A1" or "AA100".
 */
export function parseCellAddress(address: string): { row: number; col: number } {
  if (!address || typeof address !== 'string') {
    throw new Error(`Invalid cell address: "${address}".`);
  }

  const match = address.trim().toUpperCase().match(/^([A-Z]+)([1-9][0-9]*)$/);
  if (!match) {
    throw new Error(`Malformed cell address: "${address}". Expected format like "A1" or "B12".`);
  }

  const col = colLetterToIndex(match[1]);
  const row = parseInt(match[2], 10);

  if (row < 1 || row > MAX_EXCEL_ROWS) {
    throw new Error(`Row ${row} in address "${address}" exceeds Excel limit (1..${MAX_EXCEL_ROWS}).`);
  }

  return { row, col };
}

/**
 * Parses a sheet title and range string. Handles:
 * - 'Sheet 1'!A1:C10
 * - 'O''Reilly'!A1:B2 (escaped single quotes)
 * - Sheet1!A1
 * - A1:C10
 * - A1
 * - A:C (full column)
 * - 1:10 (full row)
 * - Inverted ranges like C10:A1 -> normalized to A1:C10
 */
export function parseA1Range(
  rangeStr: string,
  options?: { maxRows?: number; maxCols?: number; defaultSheet?: string }
): ParsedA1Range {
  if (!rangeStr || typeof rangeStr !== 'string' || !rangeStr.trim()) {
    throw new Error(`Empty or invalid range string provided.`);
  }

  const trimmed = rangeStr.trim();
  let sheetName: string | undefined = options?.defaultSheet;
  let coordinatePart = trimmed;

  // Extract sheet name if present
  if (trimmed.startsWith("'")) {
    // Quoted sheet name: find closing quote not followed by another quote
    let i = 1;
    let foundClosing = false;
    let extractedSheet = '';

    while (i < trimmed.length) {
      if (trimmed[i] === "'") {
        if (i + 1 < trimmed.length && trimmed[i + 1] === "'") {
          // Escaped single quote
          extractedSheet += "'";
          i += 2;
        } else {
          // Closing quote
          foundClosing = true;
          i++;
          break;
        }
      } else {
        extractedSheet += trimmed[i];
        i++;
      }
    }

    if (!foundClosing || i >= trimmed.length || trimmed[i] !== '!') {
      throw new Error(
        `Malformed sheet-qualified range "${trimmed}". Expected format: 'Sheet Name'!A1:B10`
      );
    }

    sheetName = extractedSheet;
    coordinatePart = trimmed.slice(i + 1).trim();
  } else if (trimmed.includes('!')) {
    const exclamationIdx = trimmed.indexOf('!');
    const rawSheet = trimmed.slice(0, exclamationIdx).trim();
    if (!rawSheet) {
      throw new Error(`Malformed range "${trimmed}": empty sheet name before '!'.`);
    }
    sheetName = rawSheet;
    coordinatePart = trimmed.slice(exclamationIdx + 1).trim();
  }

  if (!coordinatePart) {
    throw new Error(`Malformed range "${trimmed}": missing coordinate part after sheet name.`);
  }

  const upperCoord = coordinatePart.toUpperCase();
  const maxRows = Math.min(options?.maxRows ?? MAX_EXCEL_ROWS, MAX_EXCEL_ROWS);
  const maxCols = Math.min(options?.maxCols ?? MAX_EXCEL_COLS, MAX_EXCEL_COLS);

  // Check 1: Single cell (e.g. "A1")
  const singleCellMatch = upperCoord.match(/^([A-Z]+)([1-9][0-9]*)$/);
  if (singleCellMatch) {
    const col = colLetterToIndex(singleCellMatch[1]);
    const row = parseInt(singleCellMatch[2], 10);
    if (row > maxRows) {
      throw new Error(`Row ${row} exceeds maximum rows limit (${maxRows}).`);
    }
    if (col > maxCols) {
      throw new Error(`Column ${singleCellMatch[1]} exceeds maximum columns limit (${maxCols}).`);
    }
    return {
      raw: rangeStr,
      sheetName,
      startCol: col,
      startRow: row,
      endCol: col,
      endRow: row,
      isSingleCell: true,
      isFullRow: false,
      isFullCol: false,
    };
  }

  // Check 2: Bounded rectangle (e.g. "A1:C10")
  const rectMatch = upperCoord.match(/^([A-Z]+)([1-9][0-9]*):([A-Z]+)([1-9][0-9]*)$/);
  if (rectMatch) {
    const c1 = colLetterToIndex(rectMatch[1]);
    const r1 = parseInt(rectMatch[2], 10);
    const c2 = colLetterToIndex(rectMatch[3]);
    const r2 = parseInt(rectMatch[4], 10);

    const startCol = Math.min(c1, c2);
    const endCol = Math.max(c1, c2);
    const startRow = Math.min(r1, r2);
    const endRow = Math.max(r1, r2);

    if (endRow > maxRows) {
      throw new Error(`End row ${endRow} exceeds maximum rows limit (${maxRows}).`);
    }
    if (endCol > maxCols) {
      throw new Error(`End column exceeds maximum columns limit (${maxCols}).`);
    }

    return {
      raw: rangeStr,
      sheetName,
      startCol,
      startRow,
      endCol,
      endRow,
      isSingleCell: startCol === endCol && startRow === endRow,
      isFullRow: false,
      isFullCol: false,
    };
  }

  // Check 3: Full column range (e.g. "A:C")
  const fullColMatch = upperCoord.match(/^([A-Z]+):([A-Z]+)$/);
  if (fullColMatch) {
    const c1 = colLetterToIndex(fullColMatch[1]);
    const c2 = colLetterToIndex(fullColMatch[2]);
    const startCol = Math.min(c1, c2);
    const endCol = Math.max(c1, c2);

    return {
      raw: rangeStr,
      sheetName,
      startCol,
      startRow: 1,
      endCol,
      endRow: maxRows,
      isSingleCell: false,
      isFullRow: false,
      isFullCol: true,
    };
  }

  // Check 4: Full row range (e.g. "1:10")
  const fullRowMatch = upperCoord.match(/^([1-9][0-9]*):([1-9][0-9]*)$/);
  if (fullRowMatch) {
    const r1 = parseInt(fullRowMatch[1], 10);
    const r2 = parseInt(fullRowMatch[2], 10);
    const startRow = Math.min(r1, r2);
    const endRow = Math.max(r1, r2);

    if (endRow > maxRows) {
      throw new Error(`End row ${endRow} exceeds maximum rows limit (${maxRows}).`);
    }

    return {
      raw: rangeStr,
      sheetName,
      startCol: 1,
      startRow,
      endCol: maxCols,
      endRow,
      isSingleCell: false,
      isFullRow: true,
      isFullCol: false,
    };
  }

  throw new Error(
    `Invalid A1 range syntax: "${coordinatePart}" in "${rangeStr}". Expected format like "A1", "A1:C10", "A:C", or "1:10".`
  );
}

/**
 * Formats a sheet name for A1 notation, adding single quotes and escaping if necessary.
 */
export function formatSheetTitle(sheetName: string): string {
  if (!sheetName) return '';
  // Needs quotes if contains spaces, punctuation, quotes, or starts with a digit
  const needsQuotes = /[^a-zA-Z0-9_]/.test(sheetName) || /^[0-9]/.test(sheetName);
  if (needsQuotes) {
    const escaped = sheetName.replace(/'/g, "''");
    return `'${escaped}'`;
  }
  return sheetName;
}

/**
 * Formats range coordinates into an A1 notation string.
 */
export function formatA1Range(params: {
  sheetName?: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}): string {
  const { sheetName, startCol, startRow, endCol, endRow } = params;
  const c1 = cellAddress(startRow, startCol);
  const c2 = cellAddress(endRow, endCol);
  const coords = c1 === c2 ? c1 : `${c1}:${c2}`;

  if (sheetName) {
    return `${formatSheetTitle(sheetName)}!${coords}`;
  }
  return coords;
}

/**
 * Calculates the bounding box for placing a ReportSheet's rows.
 */
export function calculateTargetRange(
  sheet: ReportSheet,
  defaultStartCell: string = 'A1'
): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  rangeString: string;
} {
  const startAddress = sheet.startCell || defaultStartCell;
  const { row: startRow, col: startCol } = parseCellAddress(startAddress);

  const rowCount = sheet.rows.length;
  let maxColsInRows = 0;
  for (const row of sheet.rows) {
    if (row.length > maxColsInRows) {
      maxColsInRows = row.length;
    }
  }

  const endRow = rowCount > 0 ? startRow + rowCount - 1 : startRow;
  const endCol = maxColsInRows > 0 ? startCol + maxColsInRows - 1 : startCol;

  const rangeString = formatA1Range({
    sheetName: sheet.name,
    startCol,
    startRow,
    endCol,
    endRow,
  });

  return {
    startRow,
    startCol,
    endRow,
    endCol,
    rangeString,
  };
}
