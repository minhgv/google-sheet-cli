/**
 * Strict schema normalization and validation with explicit locale/date policies.
 */

import Decimal from 'decimal.js';
import { FieldSchema, TableSchema, ValidationError } from './types';

export class ReportSchemaError extends Error {
  public readonly errors: ValidationError[];

  constructor(errors: ValidationError[], tableName?: string) {
    const header = tableName
      ? `Schema validation failed for table "${tableName}" with ${errors.length} error(s):`
      : `Schema validation failed with ${errors.length} error(s):`;
    const details = errors
      .slice(0, 20)
      .map(
        (e) =>
          `  - Row ${e.row}${e.column ? `, Col ${e.column}` : ''} [${e.field}]: ${e.message} (value: ${JSON.stringify(e.value)})`
      )
      .join('\n');
    const suffix = errors.length > 20 ? `\n  ... and ${errors.length - 20} more errors` : '';
    super(`${header}\n${details}${suffix}`);
    this.name = 'ReportSchemaError';
    this.errors = errors;
  }
}

/**
 * Normalizes and validates tabular data against a TableSchema.
 * Data can be an array of records (Record<string, unknown>[]) or a 2D matrix (unknown[][]).
 */
export function normalizeTable(
  data: unknown,
  schema: TableSchema,
  tableName?: string
): Record<string, unknown>[] {
  if (!schema || !Array.isArray(schema.fields)) {
    throw new Error('Invalid TableSchema: schema.fields must be an array of FieldSchema');
  }

  if (!Array.isArray(data)) {
    throw new ReportSchemaError(
      [
        {
          row: 0,
          field: '_root',
          value: data,
          message: 'Expected an array of records or 2D matrix',
          code: 'INVALID_ROOT_TYPE',
        },
      ],
      tableName
    );
  }

  if (data.length === 0) {
    return [];
  }

  const errors: ValidationError[] = [];
  const normalizedRows: Record<string, unknown>[] = [];
  const fieldMap = new Map<string, FieldSchema>();
  for (const field of schema.fields) {
    fieldMap.set(field.name, field);
  }

  // Determine if input is 2D matrix or array of record objects
  const is2DMatrix = Array.isArray(data[0]);

  let headers: string[] = [];
  let startIndex = 0;

  if (is2DMatrix) {
    // If 2D matrix, row 0 is assumed to be headers if matching schema field names
    const row0 = data[0] as unknown[];
    const row0Strings = row0.map((c) => (c === null || c === undefined ? '' : String(c).trim()));
    const hasMatchingHeaders = row0Strings.some((str) => fieldMap.has(str));

    if (hasMatchingHeaders) {
      headers = row0Strings;
      startIndex = 1;
    } else {
      // Use schema fields in order
      headers = schema.fields.map((f) => f.name);
      startIndex = 0;
    }
  }

  for (let i = startIndex; i < data.length; i++) {
    const rawRow = data[i];
    const rowNumber = i + 1; // 1-based row index in source
    const normalizedRow: Record<string, unknown> = Object.create(null);

    let rowObj: Record<string, unknown>;

    if (is2DMatrix) {
      if (!Array.isArray(rawRow)) {
        errors.push({
          row: rowNumber,
          field: '_row',
          value: rawRow,
          message: 'Expected a matrix row array',
          code: 'INVALID_ROW_SHAPE',
        });
        continue;
      }
      rowObj = Object.create(null);
      const rowArr = rawRow as unknown[];
      for (let c = 0; c < headers.length; c++) {
        const headerName = headers[c];
        if (headerName) {
          rowObj[headerName] = c < rowArr.length ? rowArr[c] : undefined;
        }
      }
    } else {
      if (typeof rawRow !== 'object' || rawRow === null) {
        errors.push({
          row: rowNumber,
          field: '_row',
          value: rawRow,
          message: 'Expected an object record',
          code: 'INVALID_ROW_SHAPE',
        });
        continue;
      }
      rowObj = rawRow as Record<string, unknown>;
    }

    // Check for unexpected extra columns if disallowed
    if (schema.allowExtraColumns === false) {
      for (const key of Object.keys(rowObj)) {
        if (!fieldMap.has(key)) {
          errors.push({
            row: rowNumber,
            field: key,
            value: rowObj[key],
            message: `Unexpected extra column "${key}" not in schema`,
            code: 'EXTRA_COLUMN_DISALLOWED',
          });
        }
      }
    }

    // Validate each schema field
    for (let c = 0; c < schema.fields.length; c++) {
      const field = schema.fields[c];
      const rawVal = rowObj[field.name];
      const colNumber = c + 1;

      try {
        const validatedVal = validateAndNormalizeField(rawVal, field, rowNumber, colNumber);
        normalizedRow[field.name] = validatedVal;
      } catch (err: unknown) {
        if (err && typeof err === 'object' && 'code' in err) {
          errors.push(err as ValidationError);
        } else {
          const msg = err instanceof Error ? err.message : String(err);
          errors.push({
            row: rowNumber,
            column: colNumber,
            field: field.name,
            value: rawVal,
            message: msg,
            code: 'FIELD_VALIDATION_ERROR',
          });
        }
      }
    }

    // Copy extra columns if allowed
    if (schema.allowExtraColumns !== false) {
      for (const [k, v] of Object.entries(rowObj)) {
        if (!fieldMap.has(k)) {
          normalizedRow[k] = v;
        }
      }
    }

    normalizedRows.push(normalizedRow);
  }

  if (errors.length > 0) {
    throw new ReportSchemaError(errors, tableName);
  }

  return normalizedRows;
}

/**
 * Validates and normalizes a single cell value against its FieldSchema.
 */
function validateAndNormalizeField(
  rawVal: unknown,
  field: FieldSchema,
  row: number,
  column: number
): unknown {
  let val = rawVal;

  // Handle undefined / missing
  if (val === undefined) {
    if (field.required) {
      throw {
        row,
        column,
        field: field.name,
        value: val,
        message: `Field "${field.name}" is required but missing`,
        code: 'MISSING_REQUIRED_FIELD',
      } as ValidationError;
    }
    if (field.nullable) {
      return null;
    }
    return null;
  }

  // Handle empty strings
  if (typeof val === 'string' && val.trim() === '') {
    if (field.emptyAsNull || field.type !== 'string') {
      val = null;
    }
  }

  // Handle null
  if (val === null) {
    if (field.required && !field.nullable) {
      throw {
        row,
        column,
        field: field.name,
        value: val,
        message: `Field "${field.name}" is required and cannot be null`,
        code: 'REQUIRED_FIELD_NULL',
      } as ValidationError;
    }
    if (!field.nullable && field.nullable !== undefined) {
      throw {
        row,
        column,
        field: field.name,
        value: val,
        message: `Field "${field.name}" cannot be null`,
        code: 'NULL_NOT_ALLOWED',
      } as ValidationError;
    }
    return null;
  }

  // Type-specific parsing and validation
  switch (field.type) {
    case 'string': {
      const strVal = String(val);

      if (field.min !== undefined && strVal.length < Number(field.min)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `String length (${strVal.length}) is less than minimum length ${field.min}`,
          code: 'MIN_LENGTH_VIOLATION',
        } as ValidationError;
      }

      if (field.max !== undefined && strVal.length > Number(field.max)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `String length (${strVal.length}) exceeds maximum length ${field.max}`,
          code: 'MAX_LENGTH_VIOLATION',
        } as ValidationError;
      }

      if (field.enum && field.enum.length > 0) {
        const allowed = field.enum.map(String);
        if (!allowed.includes(strVal)) {
          throw {
            row,
            column,
            field: field.name,
            value: val,
            message: `Value "${strVal}" is not in allowed enum [${allowed.join(', ')}]`,
            code: 'ENUM_VIOLATION',
          } as ValidationError;
        }
      }

      return strVal;
    }

    case 'decimal': {
      let decStr: string;

      if (typeof val === 'number') {
        if (!Number.isFinite(val)) {
          throw {
            row,
            column,
            field: field.name,
            value: val,
            message: `Decimal value must be a finite number, received ${val}`,
            code: 'NON_FINITE_DECIMAL',
          } as ValidationError;
        }
        decStr = val.toString();
      } else if (typeof val === 'string') {
        let clean = val.trim();

        // Handle explicit thousand & decimal separators
        if (field.thousandSeparator) {
          const escapedThousand = field.thousandSeparator.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
          clean = clean.replace(new RegExp(escapedThousand, 'g'), '');
        }

        if (field.decimalSeparator && field.decimalSeparator !== '.') {
          const escapedDec = field.decimalSeparator.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
          clean = clean.replace(new RegExp(escapedDec, 'g'), '.');
        }

        if (clean === '' || clean === '.') {
          throw {
            row,
            column,
            field: field.name,
            value: val,
            message: `Invalid decimal format: "${val}"`,
            code: 'INVALID_DECIMAL_FORMAT',
          } as ValidationError;
        }

        decStr = clean;
      } else {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Expected number or numeric string for decimal field, received ${typeof val}`,
          code: 'INVALID_DECIMAL_TYPE',
        } as ValidationError;
      }

      try {
        let dec = new Decimal(decStr);
        if (field.decimalPlaces !== undefined) {
          dec = dec.toDecimalPlaces(field.decimalPlaces, Decimal.ROUND_HALF_UP);
        }

        if (field.min !== undefined) {
          const minDec = new Decimal(field.min);
          if (dec.lessThan(minDec)) {
            throw {
              row,
              column,
              field: field.name,
              value: val,
              message: `Decimal value ${dec.toString()} is less than minimum ${field.min}`,
              code: 'MIN_VALUE_VIOLATION',
            } as ValidationError;
          }
        }

        if (field.max !== undefined) {
          const maxDec = new Decimal(field.max);
          if (dec.greaterThan(maxDec)) {
            throw {
              row,
              column,
              field: field.name,
              value: val,
              message: `Decimal value ${dec.toString()} exceeds maximum ${field.max}`,
              code: 'MAX_VALUE_VIOLATION',
            } as ValidationError;
          }
        }

        // Return standard JS number if safely representable, otherwise string to preserve precision
        const numVal = dec.toNumber();
        if (Number.isFinite(numVal) && dec.equals(new Decimal(numVal))) {
          return numVal;
        }
        return dec.toString();
      } catch (err: unknown) {
        if (err && typeof err === 'object' && 'code' in err) {
          throw err;
        }
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Cannot parse "${val}" as valid decimal`,
          code: 'DECIMAL_PARSE_ERROR',
        } as ValidationError;
      }
    }

    case 'integer': {
      let intStr: string;

      if (typeof val === 'number') {
        if (!Number.isInteger(val)) {
          throw {
            row,
            column,
            field: field.name,
            value: val,
            message: `Expected integer but received floating point number ${val}`,
            code: 'NOT_AN_INTEGER',
          } as ValidationError;
        }
        intStr = val.toString();
      } else if (typeof val === 'string') {
        let clean = val.trim();
        if (field.thousandSeparator) {
          const escapedThousand = field.thousandSeparator.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
          clean = clean.replace(new RegExp(escapedThousand, 'g'), '');
        }
        intStr = clean;
      } else {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Expected integer or numeric string, received ${typeof val}`,
          code: 'INVALID_INTEGER_TYPE',
        } as ValidationError;
      }

      if (!/^-?\d+$/.test(intStr)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Invalid integer format: "${val}"`,
          code: 'INVALID_INTEGER_FORMAT',
        } as ValidationError;
      }

      const intNum = Number(intStr);
      if (!Number.isSafeInteger(intNum)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Integer value "${intStr}" is outside safe integer range`,
          code: 'INTEGER_OVERFLOW',
        } as ValidationError;
      }

      if (field.min !== undefined && intNum < Number(field.min)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Integer value ${intNum} is less than minimum ${field.min}`,
          code: 'MIN_VALUE_VIOLATION',
        } as ValidationError;
      }

      if (field.max !== undefined && intNum > Number(field.max)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Integer value ${intNum} exceeds maximum ${field.max}`,
          code: 'MAX_VALUE_VIOLATION',
        } as ValidationError;
      }

      return intNum;
    }

    case 'boolean': {
      if (typeof val === 'boolean') {
        return val;
      }
      if (typeof val === 'string') {
        const lower = val.trim().toLowerCase();
        if (lower === 'true' || lower === '1' || lower === 'yes' || lower === 'y' || lower === 't') {
          return true;
        }
        if (lower === 'false' || lower === '0' || lower === 'no' || lower === 'n' || lower === 'f') {
          return false;
        }
      } else if (typeof val === 'number') {
        if (val === 1) return true;
        if (val === 0) return false;
      }

      throw {
        row,
        column,
        field: field.name,
        value: val,
        message: `Cannot parse "${val}" as boolean (expected true/false/1/0/yes/no)`,
        code: 'INVALID_BOOLEAN_FORMAT',
      } as ValidationError;
    }

    case 'date': {
      let canonicalDate: string | null = null;
      if (val instanceof Date) {
        if (!isNaN(val.getTime())) {
          const y = val.getUTCFullYear();
          const m = val.getUTCMonth() + 1;
          const d = val.getUTCDate();
          canonicalDate = `${y}-${m < 10 ? '0' + m : m}-${d < 10 ? '0' + d : d}`;
        }
      } else if (typeof val === 'number') {
        // Excel / Google Sheets 1900 date system serial number (days since 1899-12-30)
        // Valid calendar window: serial >= 61 (1900-03-01) through 2958465 (9999-12-31) where Excel and Google Sheets agree 100%.
        // Serial numbers < 61 (pre-1900-03-01) and serial 60 (fictitious 1900-02-29) diverge between Excel and Google Sheets epochs and are rejected.
        if (Number.isFinite(val) && val >= 1 && val <= 2958465) {
          if (val < 61) {
            throw {
              row,
              column,
              field: field.name,
              value: val,
              message: `Serial date ${val} is in the pre-1900-03-01 epoch ambiguity band (< 61) where Excel and Google Sheets serial systems diverge. Supply historical dates as explicit ISO YYYY-MM-DD strings.`,
              code: 'AMBIGUOUS_SERIAL_DATE',
            } as ValidationError;
          }

          const ms = Math.round((val - 25569) * 86400 * 1000);
          const dObj = new Date(ms);
          if (!isNaN(dObj.getTime())) {
            const y = dObj.getUTCFullYear();
            const m = dObj.getUTCMonth() + 1;
            const d = dObj.getUTCDate();
            canonicalDate = `${y}-${m < 10 ? '0' + m : m}-${d < 10 ? '0' + d : d}`;
          }
        }
      } else {
        const dateStr = typeof val === 'string' ? val.trim() : String(val).trim();
        canonicalDate = parseAndValidateDate(dateStr, field.dateFormat);
      }

      if (!canonicalDate) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Invalid calendar date "${String(val)}"${field.dateFormat ? ` (expected format: ${field.dateFormat})` : ''}`,
          code: 'INVALID_DATE_FORMAT',
        } as ValidationError;
      }

      if (field.min !== undefined && canonicalDate < String(field.min)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Date ${canonicalDate} is earlier than minimum allowed date ${field.min}`,
          code: 'MIN_DATE_VIOLATION',
        } as ValidationError;
      }

      if (field.max !== undefined && canonicalDate > String(field.max)) {
        throw {
          row,
          column,
          field: field.name,
          value: val,
          message: `Date ${canonicalDate} is later than maximum allowed date ${field.max}`,
          code: 'MAX_DATE_VIOLATION',
        } as ValidationError;
      }

      return canonicalDate;
    }

    default:
      throw {
        row,
        column,
        field: field.name,
        value: val,
        message: `Unknown field type "${(field as FieldSchema).type}" in schema`,
        code: 'UNKNOWN_FIELD_TYPE',
      } as ValidationError;
  }
}

/**
 * Validates a date string against a calendar and returns canonical YYYY-MM-DD.
 */
function parseAndValidateDate(dateStr: string, explicitFormat?: string): string | null {
  if (!dateStr) return null;

  let year: number;
  let month: number;
  let day: number;

  if (explicitFormat === 'DD/MM/YYYY' || explicitFormat === 'DD-MM-YYYY') {
    const match = dateStr.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (!match) return null;
    day = parseInt(match[1], 10);
    month = parseInt(match[2], 10);
    year = parseInt(match[3], 10);
  } else if (explicitFormat === 'MM/DD/YYYY' || explicitFormat === 'MM-DD-YYYY') {
    const match = dateStr.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (!match) return null;
    month = parseInt(match[1], 10);
    day = parseInt(match[2], 10);
    year = parseInt(match[3], 10);
  } else if (explicitFormat === 'YYYY/MM/DD' || explicitFormat === 'YYYY-MM-DD') {
    const match = dateStr.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
    if (!match) return null;
    year = parseInt(match[1], 10);
    month = parseInt(match[2], 10);
    day = parseInt(match[3], 10);
  } else {
    // Default ISO / standard matching: YYYY-MM-DD or YYYY-MM-DDTHH:mm:ss...
    const isoMatch = dateStr.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/);
    if (isoMatch) {
      year = parseInt(isoMatch[1], 10);
      month = parseInt(isoMatch[2], 10);
      day = parseInt(isoMatch[3], 10);
    } else {
      // Try DD/MM/YYYY fallback
      const slashMatch = dateStr.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
      if (slashMatch) {
        day = parseInt(slashMatch[1], 10);
        month = parseInt(slashMatch[2], 10);
        year = parseInt(slashMatch[3], 10);
      } else {
        return null;
      }
    }
  }

  // Calendar validity checks
  if (year < 1000 || year > 9999) return null;
  if (month < 1 || month > 12) return null;

  const daysInMonth = [0, 31, isLeapYear(year) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (day < 1 || day > daysInMonth[month]) {
    return null;
  }

  const mm = month < 10 ? `0${month}` : `${month}`;
  const dd = day < 10 ? `0${day}` : `${day}`;
  return `${year}-${mm}-${dd}`;
}

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}
