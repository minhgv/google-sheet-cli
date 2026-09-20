/**
 * Coordinate-aware schema discovery and validation-schema handling for `data:schema` and
 * `data:validate`. Everything here is pure: it works on the scoped worksheet metadata the
 * `GoogleSheet.getWorksheetMetadata()` call returns (worksheet properties, bounded grid data
 * with per-cell formula/validation detail, and the spreadsheet's named ranges) and on report
 * schema types. No transport, no writes.
 *
 * Type inference from cell samples is deliberately honest about its limits: a column's kinds
 * are what the sampled cells contain, not an authoritative schema - Google stores dates as
 * numbers, and formatting decides what they look like.
 */

import type { sheets_v4 } from '@googleapis/sheets';
import type { GoogleSheetCli } from './google-sheet';
import { formatA1Cell, colToA1 } from './sheet-batch';
import { FieldSchema, FieldType, TableSchema, ValidationError } from './report/types';
import { GSheetError, GSheetErrorCode, ErrorIssue } from './cli-errors';

const CELL_KINDS = ['string', 'number', 'boolean', 'error'] as const;

export type CellKind = (typeof CELL_KINDS)[number];

export interface DiscoveredColumn {
  /** 1-based absolute column number in the worksheet */
  column: number;
  /** A1 letter of the column (e.g. "C"), absolute in the worksheet */
  letter: string;
  /** the header text as the sheet stores it (trimmed); "" when the header cell is empty */
  header: string;
  /** A1 coordinate of the header cell (e.g. "C5") */
  headerA1: string;
  /** the header cell is empty */
  emptyHeader: boolean;
  /** the same non-empty header text appears on another sampled column */
  duplicate: boolean;
  /** counts of sampled cell kinds, formulas excluded from the kind they compute */
  inferredTypes: Record<CellKind, number>;
  /** more than one kind occurs among the sampled non-formula cells */
  mixed: boolean;
  /** sampled cells with content */
  nonEmptyCells: number;
  /** sampled cells whose user-entered value is a formula */
  formulaCells: number;
  /** distinct data-validation conditions over the sampled cells, with counts */
  validations: { condition: string; count: number }[];
}

export interface NamedRangeInfo {
  name: string;
  sheetId?: number;
  /** A1 rendering of the named range's grid range (start-only when unbounded) */
  range: string;
}

export interface SchemaDiscoveryResult {
  spreadsheetId: string;
  worksheetTitle: string;
  sheetId?: number;
  /** the worksheet's full grid size, when the API reported it */
  gridRows?: number;
  gridColumns?: number;
  /** what was actually sampled - the header row is the first row of the sample */
  sample: {
    range: string;
    /** 1-based first sampled row */
    startRow: number;
    /** 1-based first sampled column */
    startColumn: number;
    rows: number;
    columns: number;
  };
  columns: DiscoveredColumn[];
  /** honest-reporting warnings: empty/duplicate headers, mixed kinds */
  warnings: string[];
  namedRanges: NamedRangeInfo[];
}

const cellKind = (cell: sheets_v4.Schema$CellData | undefined): CellKind | null => {
  const effective = cell?.effectiveValue;
  if (effective?.numberValue !== undefined && effective.numberValue !== null) return 'number';
  if (effective?.boolValue !== undefined && effective.boolValue !== null) return 'boolean';
  if (effective?.stringValue !== undefined && effective.stringValue !== null) return 'string';
  if (effective?.errorValue) return 'error';
  const entered = cell?.userEnteredValue;
  if (entered?.numberValue !== undefined && entered.numberValue !== null) return 'number';
  if (entered?.boolValue !== undefined && entered.boolValue !== null) return 'boolean';
  if (entered?.stringValue !== undefined && entered.stringValue !== null) return 'string';
  // a formula without a computed value says nothing about the result kind
  if (entered?.formulaValue !== undefined) return null;
  const formatted = cell?.formattedValue;
  return formatted ? 'string' : null;
};

const headerText = (cell: sheets_v4.Schema$CellData | undefined): string =>
  (cell?.formattedValue ?? cell?.userEnteredValue?.stringValue ?? '').trim();

const gridRangeToA1 = (range: sheets_v4.Schema$GridRange | undefined): string => {
  if (!range) return '';
  const startCol = (range.startColumnIndex ?? 0) + 1;
  const startRow = (range.startRowIndex ?? 0) + 1;
  const hasEnd = range.endColumnIndex !== undefined && range.endRowIndex !== undefined;
  const start = formatA1Cell(startCol, startRow);
  if (!hasEnd) return start;
  return `${start}:${formatA1Cell(range.endColumnIndex!, range.endRowIndex!)}`;
};

/**
 * Build the schema view of one worksheet from its scoped metadata. The first sampled row is
 * treated as the header row; every reported coordinate is absolute in the worksheet, so a
 * sample taken from a bounded range (e.g. starting at row 5) still reports C5, not C1.
 */
export function discoverSchema(meta: GoogleSheetCli.WorksheetMetadata): SchemaDiscoveryResult {
  const grid = meta.gridData[0];
  const startRow = (grid?.startRow ?? 0) + 1;
  const startColumn = (grid?.startColumn ?? 0) + 1;
  const rows = grid?.rowData ?? [];
  const headerCells = rows[0]?.values ?? [];

  const occurrences = new Map<string, string[]>();
  for (let j = 0; j < headerCells.length; j++) {
    const name = headerText(headerCells[j]);
    if (name) {
      const letter = colToA1(startColumn + j);
      const seen = occurrences.get(name) ?? [];
      seen.push(letter);
      occurrences.set(name, seen);
    }
  }

  const columns: DiscoveredColumn[] = headerCells.map((headerCell: sheets_v4.Schema$CellData, j: number) => {
    const column = startColumn + j;
    const letter = colToA1(column);
    const header = headerText(headerCell);
    const inferredTypes: Record<CellKind, number> = { string: 0, number: 0, boolean: 0, error: 0 };
    let nonEmptyCells = 0;
    let formulaCells = 0;
    const validationCounts = new Map<string, number>();
    for (let i = 1; i < rows.length; i++) {
      const cell = rows[i]?.values?.[j];
      if (!cell) continue;
      const kind = cellKind(cell);
      if (kind) {
        inferredTypes[kind] += 1;
        nonEmptyCells += 1;
      }
      if (typeof cell.userEnteredValue?.formulaValue === 'string') formulaCells += 1;
      const condition = cell.dataValidation?.condition?.type;
      if (condition) validationCounts.set(condition, (validationCounts.get(condition) ?? 0) + 1);
    }
    const kinds = CELL_KINDS.filter((kind) => inferredTypes[kind] > 0);
    const validations = [...validationCounts.entries()]
      .map(([condition, count]: [string, number]) => ({ condition, count }))
      .sort((a, b) => a.condition.localeCompare(b.condition));
    return {
      column,
      letter,
      header,
      headerA1: formatA1Cell(column, startRow),
      emptyHeader: header === '',
      duplicate: header !== '' && (occurrences.get(header)?.length ?? 0) > 1,
      inferredTypes,
      mixed: kinds.length > 1,
      nonEmptyCells,
      formulaCells,
      validations,
    };
  });

  const warnings: string[] = [];
  for (const [name, letters] of occurrences) {
    if (letters.length > 1) warnings.push(`Duplicate header "${name}" at ${letters.map((l) => `${l}${startRow}`).join(', ')}`);
  }
  for (const column of columns) {
    if (column.emptyHeader) warnings.push(`Empty header at ${column.letter}${startRow}`);
    if (column.mixed) {
      const kinds = CELL_KINDS.filter((kind) => column.inferredTypes[kind] > 0)
        .map((kind) => `${kind} (${column.inferredTypes[kind]})`)
        .join(', ');
      warnings.push(`Column ${column.letter} ("${column.header}") has mixed types: ${kinds}`);
    }
  }

  const namedRanges = meta.namedRanges
    .filter((namedRange) => namedRange.range?.sheetId === undefined || namedRange.range?.sheetId === meta.properties.sheetId)
    .map((namedRange) => ({
      name: namedRange.name ?? '',
      // the API spells an absent id `null`; the contract fields are optional numbers
      sheetId: namedRange.range?.sheetId ?? undefined,
      range: gridRangeToA1(namedRange.range),
    }));

  const sampledRows = rows.length;
  const sampledColumns = Math.max(...rows.map((row: sheets_v4.Schema$RowData) => row.values?.length ?? 0), 0);
  const gridProperties = meta.properties.gridProperties ?? {};

  return {
    spreadsheetId: meta.spreadsheetId,
    worksheetTitle: meta.worksheetTitle,
    // the API types these as `number | null | undefined`; the discovery contract carries
    // optional numbers, so an explicit null normalizes to a plain absent field
    sheetId: meta.properties.sheetId ?? undefined,
    gridRows: gridProperties.rowCount ?? undefined,
    gridColumns: gridProperties.columnCount ?? undefined,
    sample: {
      range: `${meta.worksheetTitle ? `${meta.worksheetTitle}!` : ''}${formatA1Cell(startColumn, startRow)}:${formatA1Cell(
        startColumn + Math.max(sampledColumns, 1) - 1,
        startRow + Math.max(sampledRows, 1) - 1
      )}`,
      startRow,
      startColumn,
      rows: sampledRows,
      columns: sampledColumns,
    },
    columns,
    warnings,
    namedRanges,
  };
}

// ============================================================================
// Validation schema handling (data:validate)
// ============================================================================

const FIELD_TYPES: Record<FieldType, true> = { string: true, decimal: true, integer: true, boolean: true, date: true };

const FIELD_KEYS: Record<string, true> = {
  name: true,
  type: true,
  required: true,
  nullable: true,
  emptyAsNull: true,
  dateFormat: true,
  decimalSeparator: true,
  thousandSeparator: true,
  decimalPlaces: true,
  enum: true,
  min: true,
  max: true,
  description: true,
  unique: true,
};

const TABLE_KEYS: Record<string, true> = { fields: true, allowExtraColumns: true };

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

const isPrimitive = (value: unknown): value is string | number | boolean =>
  typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';

const schemaProblem = (field: string, message: string): ErrorIssue => ({
  field,
  code: 'SCHEMA_INVALID',
  message,
});

export interface ParsedValidationSchema {
  schema: TableSchema;
  /** fields declared `unique: true` - not a report FieldSchema property, so tracked separately */
  uniqueFields: ReadonlySet<string>;
}

/**
 * Strictly parse and validate a TableSchema from untrusted input. Unknown keys are rejected
 * instead of silently ignored (a typo'd "requird" must not disable a rule), and every problem
 * is collected into one SCHEMA_INVALID error rather than failing one at a time.
 */
export function parseValidationSchema(input: unknown): ParsedValidationSchema {
  const problems: ErrorIssue[] = [];
  if (!isPlainObject(input)) {
    throw new GSheetError(GSheetErrorCode.SCHEMA_INVALID, 'Validation schema must be a JSON object with a "fields" array');
  }
  for (const key of Object.keys(input)) {
    if (!TABLE_KEYS[key]) problems.push(schemaProblem(key, `Unknown schema property "${key}"`));
  }

  if (!Array.isArray(input.fields) || input.fields.length === 0) {
    throw new GSheetError(GSheetErrorCode.SCHEMA_INVALID, 'Validation schema must have a non-empty "fields" array', {
      details: { issues: problems.length ? problems : undefined },
    });
  }

  const fields: FieldSchema[] = [];
  const uniqueFields = new Set<string>();
  const seenNames = new Map<string, true>();

  input.fields.forEach((rawField: unknown, index: number) => {
    const label = isPlainObject(rawField) && typeof rawField.name === 'string' ? `"${rawField.name}"` : `at index ${index}`;
    if (!isPlainObject(rawField)) {
      problems.push(schemaProblem(label, `Field ${label} must be an object`));
      return;
    }
    for (const key of Object.keys(rawField)) {
      if (!FIELD_KEYS[key]) problems.push(schemaProblem(String(rawField.name ?? key), `Unknown property "${key}" on field ${label}`));
    }
    const name = rawField.name;
    if (typeof name !== 'string' || name.trim() === '') {
      problems.push(schemaProblem(label, `Field ${label} must have a non-empty string "name"`));
      return;
    }
    if (seenNames.has(name)) {
      problems.push(schemaProblem(name, `Duplicate field name "${name}"`));
    }
    seenNames.set(name, true);

    const type = rawField.type;
    if (typeof type !== 'string' || !FIELD_TYPES[type as FieldType]) {
      problems.push(schemaProblem(name, `Field "${name}" must have "type" one of string, decimal, integer, boolean, date`));
    }

    for (const flag of ['required', 'nullable', 'emptyAsNull', 'unique'] as const) {
      if (rawField[flag] !== undefined && typeof rawField[flag] !== 'boolean') {
        problems.push(schemaProblem(name, `Field "${name}" property "${flag}" must be a boolean`));
      }
    }
    if (rawField.unique === true) uniqueFields.add(name);

    if (rawField.enum !== undefined) {
      const options = rawField.enum;
      if (!Array.isArray(options) || options.length === 0 || !options.every(isPrimitive)) {
        problems.push(schemaProblem(name, `Field "${name}" property "enum" must be a non-empty array of strings, numbers or booleans`));
      }
    }
    for (const bound of ['min', 'max'] as const) {
      if (rawField[bound] !== undefined && typeof rawField[bound] !== 'number' && typeof rawField[bound] !== 'string') {
        problems.push(schemaProblem(name, `Field "${name}" property "${bound}" must be a number or string`));
      }
    }
    if (rawField.decimalPlaces !== undefined && (typeof rawField.decimalPlaces !== 'number' || !Number.isInteger(rawField.decimalPlaces) || rawField.decimalPlaces < 0)) {
      problems.push(schemaProblem(name, `Field "${name}" property "decimalPlaces" must be a non-negative integer`));
    }
    for (const separator of ['decimalSeparator', 'thousandSeparator'] as const) {
      if (rawField[separator] !== undefined && (typeof rawField[separator] !== 'string' || rawField[separator].length !== 1)) {
        problems.push(schemaProblem(name, `Field "${name}" property "${separator}" must be a single character`));
      }
    }
    if (rawField.dateFormat !== undefined && (typeof rawField.dateFormat !== 'string' || rawField.dateFormat.trim() === '')) {
      problems.push(schemaProblem(name, `Field "${name}" property "dateFormat" must be a non-empty string`));
    }
    if (rawField.description !== undefined && typeof rawField.description !== 'string') {
      problems.push(schemaProblem(name, `Field "${name}" property "description" must be a string`));
    }

    fields.push({
      name,
      ...(typeof type === 'string' ? { type: type as FieldType } : ({} as { type: FieldType })),
      ...(rawField.required !== undefined ? { required: rawField.required as boolean } : {}),
      ...(rawField.nullable !== undefined ? { nullable: rawField.nullable as boolean } : {}),
      ...(rawField.emptyAsNull !== undefined ? { emptyAsNull: rawField.emptyAsNull as boolean } : {}),
      ...(rawField.dateFormat !== undefined ? { dateFormat: rawField.dateFormat as string } : {}),
      ...(rawField.decimalSeparator !== undefined ? { decimalSeparator: rawField.decimalSeparator as string } : {}),
      ...(rawField.thousandSeparator !== undefined ? { thousandSeparator: rawField.thousandSeparator as string } : {}),
      ...(rawField.decimalPlaces !== undefined ? { decimalPlaces: rawField.decimalPlaces as number } : {}),
      ...(rawField.enum !== undefined ? { enum: rawField.enum as (string | number | boolean)[] } : {}),
      ...(rawField.min !== undefined ? { min: rawField.min as number | string } : {}),
      ...(rawField.max !== undefined ? { max: rawField.max as number | string } : {}),
      ...(rawField.description !== undefined ? { description: rawField.description as string } : {}),
    });
  });

  if (input.allowExtraColumns !== undefined && typeof input.allowExtraColumns !== 'boolean') {
    problems.push(schemaProblem('allowExtraColumns', 'Schema property "allowExtraColumns" must be a boolean'));
  }

  if (problems.length > 0) {
    throw new GSheetError(GSheetErrorCode.SCHEMA_INVALID, `Validation schema is invalid (${problems.length} problem${problems.length === 1 ? '' : 's'})`, {
      details: { issues: problems },
    });
  }

  return {
    schema: {
      fields,
      ...(input.allowExtraColumns !== undefined ? { allowExtraColumns: input.allowExtraColumns as boolean } : {}),
    },
    uniqueFields,
  };
}

/** The scalar a validate run feeds to normalizeTable: RAW cell values, not formatted text. */
export function cellScalar(cell: sheets_v4.Schema$CellData | undefined): string | number | boolean | null {
  if (!cell) return null;
  const effective = cell.effectiveValue;
  if (effective?.numberValue !== undefined && effective.numberValue !== null) return effective.numberValue;
  if (effective?.boolValue !== undefined && effective.boolValue !== null) return effective.boolValue;
  if (effective?.stringValue !== undefined && effective.stringValue !== null) return effective.stringValue;
  const entered = cell.userEnteredValue;
  if (entered?.numberValue !== undefined && entered.numberValue !== null) return entered.numberValue;
  if (entered?.boolValue !== undefined && entered.boolValue !== null) return entered.boolValue;
  if (entered?.stringValue !== undefined && entered.stringValue !== null) return entered.stringValue;
  return null;
}

/** First column per header name; sheet headers must be de-duplicated before validation runs. */
export const columnIndexByHeader = (discovery: SchemaDiscoveryResult): Map<string, DiscoveredColumn> => {
  const byHeader = new Map<string, DiscoveredColumn>();
  for (const column of discovery.columns) {
    if (column.header !== '' && !byHeader.has(column.header)) byHeader.set(column.header, column);
  }
  return byHeader;
};

/**
 * Map collected normalization issues onto exact worksheet coordinates. normalizeTable reports
 * rows 1-based over the matrix it was given (header row included), so sheet row =
 * sample.startRow + row - 1; columns are looked up by the field's header position, which is
 * absolute in the worksheet regardless of the sampled bounds.
 */
export function mapReportIssues(errors: readonly ValidationError[], discovery: SchemaDiscoveryResult): ErrorIssue[] {
  const byHeader = columnIndexByHeader(discovery);
  return errors.map((issue: ValidationError) => {
    const sheetRow = discovery.sample.startRow + issue.row - 1;
    const column = byHeader.get(issue.field)?.column;
    const mapped: ErrorIssue = {
      row: sheetRow,
      field: issue.field,
      code: issue.code,
      message: issue.message,
    };
    if (column !== undefined) {
      mapped.column = column;
      mapped.a1 = formatA1Cell(column, sheetRow);
    }
    if (issue.value === null || typeof issue.value === 'string' || typeof issue.value === 'number' || typeof issue.value === 'boolean') {
      mapped.value = issue.value;
    }
    return mapped;
  });
}

/**
 * Uniqueness violations for the fields declared `unique: true`. Null/empty values never
 * collide (SQL-style), and keys are type-aware: the string "001" and the number 1 are
 * different keys, so leading-zero identifiers keep their meaning.
 */
export function findUniqueViolations(
  normalizedRows: Record<string, unknown>[],
  uniqueFields: ReadonlySet<string>,
  discovery: SchemaDiscoveryResult
): ErrorIssue[] {
  const byHeader = columnIndexByHeader(discovery);
  const issues: ErrorIssue[] = [];
  for (const field of uniqueFields) {
    const firstSeen = new Map<string, number>();
    normalizedRows.forEach((record: Record<string, unknown>, k: number) => {
      const value = record[field];
      if (value === null || value === undefined) return;
      const key = `${typeof value}:${String(value)}`;
      const firstRow = firstSeen.get(key);
      if (firstRow === undefined) {
        firstSeen.set(key, k);
        return;
      }
      const sheetRow = discovery.sample.startRow + k + 1;
      const column = byHeader.get(field)?.column;
      const issue: ErrorIssue = {
        row: sheetRow,
        field,
        code: 'DUPLICATE_VALUE',
        message: `Duplicate value for unique field "${field}" (first seen at sheet row ${discovery.sample.startRow + firstRow + 1})`,
      };
      if (column !== undefined) {
        issue.column = column;
        issue.a1 = formatA1Cell(column, sheetRow);
      }
      if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') issue.value = value;
      issues.push(issue);
    });
  }
  return issues;
}
