import * as fs from 'fs';
import { GoogleSheetCli } from './google-sheet';
import { parseInput, parseJsonSafely, readInput } from './report/input';
import { ValidationError } from './validation-error';

export interface CommandInputOptions {
  dataArg?: string;
  inputFile?: string;
  inputFormat?: 'json' | 'csv';
}

/**
 * Resolves and validates a 2D matrix (RawData) from either a positional DATA argument
 * or an --input flag (file or '-' for stdin).
 */
export async function resolveDataMatrix(
  dataArg?: string,
  inputFile?: string,
  inputFormat?: string
): Promise<GoogleSheetCli.RawData> {
  if (dataArg !== undefined && dataArg !== '' && inputFile !== undefined && inputFile !== '') {
    throw new ValidationError('Mutually exclusive input sources: specify either positional data argument or --input flag, not both');
  }

  if ((dataArg === undefined || dataArg === '') && (inputFile === undefined || inputFile === '')) {
    throw new ValidationError('No data provided. Specify either positional data argument or --input flag (pipe data in via --input=-)');
  }

  const format = inputFormat ? (inputFormat.toLowerCase() as 'json' | 'csv') : undefined;
  if (format && format !== 'json' && format !== 'csv') {
    throw new ValidationError(`Unsupported input format "${inputFormat}"; expected "json" or "csv"`);
  }

  let rawParsed: unknown;

  if (dataArg !== undefined && dataArg !== '') {
    let parsed: unknown = dataArg;
    // Support nested stringified JSON as in legacy 2.x/3.x
    if (typeof parsed === 'string') {
      const trimmed = parsed.trim();
      if (format === 'csv') {
        parsed = parseInput(dataArg, 'csv', { matrix: true });
      } else {
        try {
          parsed = parseJsonSafely(trimmed);
          while (typeof parsed === 'string' && (parsed.trim().startsWith('[') || parsed.trim().startsWith('{'))) {
            parsed = parseJsonSafely(parsed);
          }
        } catch (err: unknown) {
          const msg = err instanceof Error ? err.message : String(err);
          throw new ValidationError(`"data" input has to be valid JSON (${msg})`);
        }
      }
    }
    rawParsed = parsed;
  } else if (inputFile !== undefined && inputFile !== '') {
    if (inputFile === '-') {
      // Read stdin
      const content = await readStdin();
      const detectedFormat = format || (content.trim().startsWith('[') || content.trim().startsWith('{') ? 'json' : 'csv');
      rawParsed = parseInput(content, detectedFormat, { matrix: true });
    } else {
      const stat = await fs.promises.stat(inputFile).catch((err) => {
        throw new ValidationError(`Cannot read input file "${inputFile}": ${err.message}`);
      });
      if (stat.size > 20 * 1024 * 1024) {
        throw new ValidationError(`Input file size (${stat.size} bytes) exceeds maximum limit of 20MB`);
      }
      const content = await fs.promises.readFile(inputFile, 'utf8');
      let detectedFormat = format;
      if (!detectedFormat) {
        if (inputFile.toLowerCase().endsWith('.json')) {
          detectedFormat = 'json';
        } else if (inputFile.toLowerCase().endsWith('.csv')) {
          detectedFormat = 'csv';
        } else {
          detectedFormat = content.trim().startsWith('[') || content.trim().startsWith('{') ? 'json' : 'csv';
        }
      }
      rawParsed = parseInput(content, detectedFormat, { matrix: true });
    }
  }

  return validate2DMatrix(rawParsed);
}

/**
 * Validates that an unknown value is a valid 2D array of scalar values.
 */
export function validate2DMatrix(val: unknown): GoogleSheetCli.RawData {
  if (!Array.isArray(val)) {
    throw new ValidationError('Data input must be a 2D array: (string | number | boolean | null)[][]');
  }

  const result: GoogleSheetCli.RawData = [];
  for (let r = 0; r < val.length; r++) {
    const row = val[r];
    if (!Array.isArray(row)) {
      throw new ValidationError(`Row at index ${r} must be an array, but received ${typeof row}`);
    }
    const validatedRow: (string | number | boolean | null)[] = [];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell === null || cell === undefined) {
        validatedRow.push(null);
      } else if (typeof cell === 'string' || typeof cell === 'number' || typeof cell === 'boolean') {
        validatedRow.push(cell);
      } else {
        throw new ValidationError(
          `Cell at row ${r}, column ${c} must be a string, number, boolean, or null; received ${typeof cell}`
        );
      }
    }
    result.push(validatedRow);
  }

  return result;
}

/**
 * Resolves and validates batch updates input from DATA argument or --input flag.
 * Input format must be an array of objects: { range: string, values: (string | number | boolean | null)[][] }
 */
export async function resolveBatchUpdates(
  dataArg?: string,
  inputFile?: string,
  inputFormat?: string
): Promise<{ range: string; values: GoogleSheetCli.RawData }[]> {
  if (dataArg !== undefined && dataArg !== '' && inputFile !== undefined && inputFile !== '') {
    throw new ValidationError('Mutually exclusive input sources: specify either positional data argument or --input flag, not both');
  }

  if ((dataArg === undefined || dataArg === '') && (inputFile === undefined || inputFile === '')) {
    throw new ValidationError('No data provided. Specify either positional data argument or --input flag (pipe data in via --input=-)');
  }

  let rawParsed: unknown;

  if (dataArg !== undefined && dataArg !== '') {
    let parsed: unknown = dataArg;
    if (typeof parsed === 'string') {
      try {
        parsed = parseJsonSafely(parsed);
        while (typeof parsed === 'string' && (parsed.trim().startsWith('[') || parsed.trim().startsWith('{'))) {
          parsed = parseJsonSafely(parsed);
        }
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw new ValidationError(`"data" input has to be valid JSON (${msg})`);
      }
    }
    rawParsed = parsed;
  } else if (inputFile !== undefined && inputFile !== '') {
    if (inputFile === '-') {
      const content = await readStdin();
      rawParsed = parseInput(content, 'json');
    } else {
      const content = await fs.promises.readFile(inputFile, 'utf8');
      rawParsed = parseInput(content, 'json');
    }
  }

  if (!Array.isArray(rawParsed)) {
    throw new ValidationError('Batch update data must be an array of { range: string, values: any[][] } objects');
  }

  if (rawParsed.length === 0) {
    throw new ValidationError('Batch update data array cannot be empty');
  }

  const updates: { range: string; values: GoogleSheetCli.RawData }[] = [];

  for (let i = 0; i < rawParsed.length; i++) {
    const item = rawParsed[i];
    if (!item || typeof item !== 'object') {
      throw new ValidationError(`Batch update item at index ${i} must be an object with "range" and "values" properties`);
    }

    const { range, values } = item as { range?: unknown; values?: unknown };
    if (!range || typeof range !== 'string' || !range.trim()) {
      throw new ValidationError(`Batch update item at index ${i} is missing a valid "range" string`);
    }

    if (!Array.isArray(values)) {
      throw new ValidationError(`Batch update item at index ${i} ("${range}") must contain a 2D "values" array`);
    }

    const validatedMatrix = validate2DMatrix(values);
    updates.push({
      range: range.trim(),
      values: validatedMatrix,
    });
  }

  return updates;
}

/**
 * Parses and validates the --ranges flag as a JSON array of strings.
 */
export function parseRangesFlag(rangesFlag?: string): string[] {
  if (!rangesFlag || typeof rangesFlag !== 'string' || !rangesFlag.trim()) {
    throw new ValidationError('The --ranges flag is required and must be a non-empty JSON array of A1 range strings');
  }

  const trimmed = rangesFlag.trim();
  let parsed: unknown;
  try {
    parsed = parseJsonSafely(trimmed);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new ValidationError(`Failed to parse --ranges flag: input must be a valid JSON array of range strings (${msg})`);
  }

  if (!Array.isArray(parsed)) {
    throw new ValidationError('The --ranges flag must be a JSON array of A1 range strings (e.g. \'["Sheet1!A1:B10", "Sheet2!C1:D5"]\')');
  }

  if (parsed.length === 0) {
    throw new ValidationError('The --ranges flag array cannot be empty');
  }

  const ranges: string[] = [];
  for (let i = 0; i < parsed.length; i++) {
    const r = parsed[i];
    if (typeof r !== 'string' || !r.trim()) {
      throw new ValidationError(`Range at index ${i} must be a non-empty string`);
    }
    ranges.push(r.trim());
  }

  return ranges;
}

/**
 * Reads all content from stdin up to a maximum limit of 20MB.
 */
async function readStdin(maxBytes = 20 * 1024 * 1024): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const chunks: Buffer[] = [];
    let totalBytes = 0;

    const onData = (chunk: Buffer) => {
      totalBytes += chunk.length;
      if (totalBytes > maxBytes) {
        process.stdin.removeListener('data', onData);
        process.stdin.removeListener('end', onEnd);
        process.stdin.removeListener('error', onError);
        reject(new Error(`Stdin input exceeded maximum limit of ${maxBytes} bytes`));
        return;
      }
      chunks.push(chunk);
    };

    const onEnd = () => {
      resolve(Buffer.concat(chunks).toString('utf8'));
    };

    const onError = (err: Error) => {
      reject(new Error(`Failed to read from stdin: ${err.message}`));
    };

    if (process.stdin.isTTY) {
      resolve('');
      return;
    }

    process.stdin.on('data', onData);
    process.stdin.on('end', onEnd);
    process.stdin.on('error', onError);
    process.stdin.resume();
  });
}
