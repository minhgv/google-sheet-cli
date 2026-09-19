import { sheets_v4 } from '@googleapis/sheets';
import { parseStrictA1Range, toGoogleNumberFormat } from './sheet-batch';

/**
 * Pure builders that translate the CLI's cell-style spec into Sheets API
 * `spreadsheets.batchUpdate` requests. No I/O happens here: the same functions
 * produce the `--dryRun` preview and the payload that is actually sent, so what
 * the user previews is byte-for-byte what Google receives.
 *
 * The field mask is computed from exactly the properties supplied - never `*` -
 * and is confined to `userEnteredFormat.*`, so a format request cannot overwrite
 * cell values by construction.
 */

export type BorderSide = 'top' | 'bottom' | 'left' | 'right' | 'innerHorizontal' | 'innerVertical';
export type BorderStyle =
  | 'DOTTED'
  | 'DASHED'
  | 'SOLID'
  | 'SOLID_MEDIUM'
  | 'SOLID_THICK'
  | 'DOUBLE'
  | 'NONE';
export type HorizontalAlignment = 'LEFT' | 'CENTER' | 'RIGHT';
export type VerticalAlignment = 'TOP' | 'MIDDLE' | 'BOTTOM';
export type WrapStrategy = 'OVERFLOW_CELL' | 'CLIP' | 'WRAP';
export type NumberFormatType =
  | 'TEXT'
  | 'NUMBER'
  | 'PERCENT'
  | 'CURRENCY'
  | 'DATE'
  | 'TIME'
  | 'DATE_TIME'
  | 'SCIENTIFIC';
export type MergeType = 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS';

export interface BorderSpecInput {
  /** sides to draw; "all" expands to the four outer sides, "inner" to both inner sides */
  sides: BorderSide[] | 'all' | 'inner';
  style?: BorderStyle;
  /** #RRGGBB; defaults to black */
  color?: string;
}

export interface CellStyleSpec {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontSize?: number;
  fontFamily?: string;
  /** #RRGGBB */
  textColor?: string;
  /** #RRGGBB */
  backgroundColor?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  wrapStrategy?: WrapStrategy;
  /** e.g. "#,##0.00", "0.0%", "YYYY-MM-DD" */
  numberFormat?: string;
  /** explicit Google type; inferred from the pattern when omitted */
  numberFormatType?: NumberFormatType;
  borders?: BorderSpecInput;
}

export interface FormatSpec {
  /** A1 ranges; a quoted title inside a range overrides `worksheetTitle` */
  ranges: string[];
  /** fallback worksheet for ranges that carry no title */
  worksheetTitle?: string;
  /** reset all formatting on the ranges; mutually exclusive with `style` */
  clear?: boolean;
  style?: CellStyleSpec;
}

export interface BuiltFormatRequests {
  requests: sheets_v4.Schema$Request[];
  /** field masks actually used, for receipts and dry-run output */
  fields: string[];
}

const BORDER_SIDES: BorderSide[] = ['top', 'bottom', 'left', 'right', 'innerHorizontal', 'innerVertical'];
const OUTER_SIDES: BorderSide[] = ['top', 'bottom', 'left', 'right'];
const INNER_SIDES: BorderSide[] = ['innerHorizontal', 'innerVertical'];

/**
 * "#RRGGBB" (or bare "RRGGBB") to a Schema$Color. Alpha is left unset: Sheets
 * treats an absent alpha as fully opaque, and sending 1.0 explicitly renders
 * identically while keeping the request smaller.
 */
export const hexToRgbColor = (hex: string): sheets_v4.Schema$Color => {
  if (typeof hex !== 'string') throw new Error(`Invalid color "${hex}": expected "#RRGGBB"`);
  const normalized = hex.trim().replace(/^#/, '');
  if (!/^[0-9a-fA-F]{6}$/.test(normalized)) {
    throw new Error(`Invalid color "${hex}": expected "#RRGGBB" (e.g. "#1a73e8")`);
  }
  return {
    red: parseInt(normalized.slice(0, 2), 16) / 255,
    green: parseInt(normalized.slice(2, 4), 16) / 255,
    blue: parseInt(normalized.slice(4, 6), 16) / 255,
  };
};

/**
 * Convert an A1 range to a 0-based, end-exclusive Schema$GridRange.
 * Unbounded sides stay undefined, which is how the API spells "to the edge".
 */
export const a1ToGridRange = (a1: string, sheetId: number): sheets_v4.Schema$GridRange => {
  const parsed = parseStrictA1Range(a1);
  if (!parsed.isBounded && parsed.startCol === undefined && parsed.startRow === undefined) {
    throw new Error(`Invalid range "${a1}": expected a bounded A1 range like "A1:C9"`);
  }
  const gridRange: sheets_v4.Schema$GridRange = { sheetId };
  if (parsed.startRow !== undefined) gridRange.startRowIndex = parsed.startRow - 1;
  if (parsed.startCol !== undefined) gridRange.startColumnIndex = parsed.startCol - 1;
  // An omitted end index means "to the grid edge" in the API. A bare cell ("B2")
  // or anchor ("B2:") must pin the end to the start or one cell would format the
  // whole row/column/grid. Column-only ("B:D") and row-only ("2:4") ranges keep
  // the other axis unbounded on purpose.
  if (parsed.endRow !== undefined) {
    gridRange.endRowIndex = parsed.endRow;
  } else if (parsed.startRow !== undefined) {
    gridRange.endRowIndex = parsed.startRow;
  }
  if (parsed.endCol !== undefined) {
    gridRange.endColumnIndex = parsed.endCol;
  } else if (parsed.startCol !== undefined) {
    gridRange.endColumnIndex = parsed.startCol;
  }
  return gridRange;
};

/**
 * Build the CellFormat payload and its field mask from a style spec.
 * Only supplied properties appear in the mask - a `--bold` call touches
 * `textFormat.bold` and nothing else.
 */
export const buildCellFormat = (
  style: CellStyleSpec
): { cellFormat: sheets_v4.Schema$CellFormat; fields: string[] } => {
  const cellFormat: sheets_v4.Schema$CellFormat = {};
  const fields: string[] = [];

  const textFormat: sheets_v4.Schema$TextFormat = {};
  if (style.bold !== undefined) textFormat.bold = style.bold;
  if (style.italic !== undefined) textFormat.italic = style.italic;
  if (style.underline !== undefined) textFormat.underline = style.underline;
  if (style.strikethrough !== undefined) textFormat.strikethrough = style.strikethrough;
  if (style.fontSize !== undefined) {
    if (!Number.isFinite(style.fontSize) || style.fontSize <= 0) {
      throw new Error(`Invalid fontSize "${style.fontSize}": expected a positive number`);
    }
    textFormat.fontSize = style.fontSize;
  }
  if (style.fontFamily !== undefined) textFormat.fontFamily = style.fontFamily;
  if (style.textColor !== undefined) {
    textFormat.foregroundColorStyle = { rgbColor: hexToRgbColor(style.textColor) };
  }
  if (Object.keys(textFormat).length > 0) {
    cellFormat.textFormat = textFormat;
    for (const key of Object.keys(textFormat)) {
      fields.push(`userEnteredFormat.textFormat.${key === 'foregroundColorStyle' ? 'foregroundColor' : key}`);
    }
  }

  if (style.backgroundColor !== undefined) {
    cellFormat.backgroundColorStyle = { rgbColor: hexToRgbColor(style.backgroundColor) };
    fields.push('userEnteredFormat.backgroundColor');
  }
  if (style.horizontalAlignment !== undefined) {
    cellFormat.horizontalAlignment = style.horizontalAlignment;
    fields.push('userEnteredFormat.horizontalAlignment');
  }
  if (style.verticalAlignment !== undefined) {
    cellFormat.verticalAlignment = style.verticalAlignment;
    fields.push('userEnteredFormat.verticalAlignment');
  }
  if (style.wrapStrategy !== undefined) {
    cellFormat.wrapStrategy = style.wrapStrategy;
    fields.push('userEnteredFormat.wrapStrategy');
  }
  if (style.numberFormat !== undefined) {
    cellFormat.numberFormat = style.numberFormatType
      ? { type: style.numberFormatType, pattern: style.numberFormat }
      : toGoogleNumberFormat(style.numberFormat);
    fields.push('userEnteredFormat.numberFormat');
  }

  return { cellFormat, fields };
};

const expandSides = (sides: BorderSpecInput['sides']): BorderSide[] => {
  if (sides === 'all') return [...OUTER_SIDES];
  if (sides === 'inner') return [...INNER_SIDES];
  if (!Array.isArray(sides) || sides.length === 0) {
    throw new Error('borders.sides must be a non-empty array, "all" or "inner"');
  }
  const expanded: BorderSide[] = [];
  for (const side of sides) {
    if (side === ('all' as BorderSide)) {
      expanded.push(...OUTER_SIDES);
    } else if (!BORDER_SIDES.includes(side)) {
      throw new Error(`Invalid border side "${side}": expected one of ${[...BORDER_SIDES, 'all'].join(', ')}`);
    } else {
      expanded.push(side);
    }
  }
  // dedupe while keeping order
  return expanded.filter((side, idx) => expanded.indexOf(side) === idx);
};

/**
 * Build the batchUpdate requests for a format spec. `sheetIdFor` resolves the
 * worksheet title embedded in (or defaulted onto) each range to a sheetId; the
 * caller owns that lookup because it needs the spreadsheet metadata.
 */
export const buildFormatRequests = (
  spec: FormatSpec,
  sheetIdFor: (worksheetTitle: string | undefined) => number
): BuiltFormatRequests => {
  if (!spec || !Array.isArray(spec.ranges) || spec.ranges.length === 0) {
    throw new Error('Format spec requires a non-empty "ranges" array');
  }
  if (spec.clear && spec.style) {
    throw new Error('Format spec cannot combine "clear" with style properties');
  }
  if (!spec.clear && !spec.style) {
    throw new Error('Format spec requires either "clear" or style properties');
  }

  const requests: sheets_v4.Schema$Request[] = [];
  const allFields: string[] = [];

  for (const a1 of spec.ranges) {
    const parsed = parseStrictA1Range(a1);
    const sheetId = sheetIdFor(parsed.worksheetTitle ?? spec.worksheetTitle);
    const gridRange = a1ToGridRange(a1, sheetId);

    if (spec.clear) {
      // An empty userEnteredFormat under that exact mask resets every format
      // property in the range; values are untouched because the mask cannot
      // name them.
      requests.push({
        repeatCell: {
          range: gridRange,
          cell: { userEnteredFormat: {} },
          fields: 'userEnteredFormat',
        },
      });
      allFields.push('userEnteredFormat');
      continue;
    }

    const style = spec.style as CellStyleSpec;
    const { cellFormat, fields } = buildCellFormat(style);
    if (fields.length > 0) {
      requests.push({
        repeatCell: {
          range: gridRange,
          cell: { userEnteredFormat: cellFormat },
          fields: fields.join(','),
        },
      });
      allFields.push(...fields);
    }

    if (style.borders) {
      const border: sheets_v4.Schema$Border = {
        style: style.borders.style || 'SOLID',
        colorStyle: { rgbColor: hexToRgbColor(style.borders.color || '#000000') },
      };
      const updateBorders: sheets_v4.Schema$UpdateBordersRequest = { range: gridRange };
      for (const side of expandSides(style.borders.sides)) {
        updateBorders[side] = border;
      }
      requests.push({ updateBorders });
      allFields.push('borders');
    }
  }

  if (requests.length === 0) {
    throw new Error('Format spec produced no requests: supply at least one style property or border');
  }
  return { requests, fields: [...new Set(allFields)] };
};

/**
 * Build a mergeCells or unmergeCells request for one range.
 */
export const buildMergeRequest = (
  a1: string,
  mergeType: MergeType | undefined,
  unmerge: boolean,
  sheetIdFor: (worksheetTitle: string | undefined) => number
): sheets_v4.Schema$Request => {
  const parsed = parseStrictA1Range(a1);
  const gridRange = a1ToGridRange(a1, sheetIdFor(parsed.worksheetTitle));
  if (unmerge) {
    return { unmergeCells: { range: gridRange } };
  }
  return { mergeCells: { range: gridRange, mergeType: mergeType || 'MERGE_ALL' } };
};
