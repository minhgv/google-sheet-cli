/**
 * Secure, bounded input reading and parsing for JSON and CSV datasets.
 */

import * as fs from 'fs';
import { parse as parseCsvSync } from 'csv-parse/sync';
import { ParseInputOptions, ReadInputOptions } from './types';

const DEFAULT_MAX_BYTES = 20 * 1024 * 1024; // 20 MB

/**
 * Reads and parses input from inline data, a file, or stdin.
 */
export async function readInput(options: ReadInputOptions): Promise<unknown> {
  if (!options) {
    throw new Error('readInput options must be provided');
  }

  const { data, file, format, maxBytes = DEFAULT_MAX_BYTES } = options;

  if (data !== undefined && file !== undefined) {
    throw new Error('Mutually exclusive options: specify either "data" or "file", not both');
  }

  if (data === undefined && file === undefined) {
    throw new Error('Must provide either "data" or "file" to readInput');
  }

  if (format !== undefined && format !== 'json' && format !== 'csv') {
    throw new Error(`Unsupported input format "${format}"; must be "json" or "csv"`);
  }

  let rawContent: string;
  let detectedFormat: 'json' | 'csv' | undefined = format;

  if (data !== undefined) {
    if (typeof data !== 'string') {
      throw new Error('Input "data" must be a string');
    }
    const byteLength = Buffer.byteLength(data, 'utf8');
    if (byteLength > maxBytes) {
      throw new Error(`Input data size (${byteLength} bytes) exceeds limit of ${maxBytes} bytes`);
    }
    rawContent = data;
  } else if (file === '-') {
    rawContent = await readStdin(maxBytes);
  } else if (typeof file === 'string') {
    const stat = await fs.promises.stat(file);
    if (stat.size > maxBytes) {
      throw new Error(`Input file size (${stat.size} bytes) exceeds limit of ${maxBytes} bytes`);
    }
    rawContent = await fs.promises.readFile(file, 'utf8');
    if (!detectedFormat) {
      if (file.toLowerCase().endsWith('.json')) {
        detectedFormat = 'json';
      } else if (file.toLowerCase().endsWith('.csv')) {
        detectedFormat = 'csv';
      }
    }
  } else {
    throw new Error('Invalid file option; expected file path string or "-" for stdin');
  }

  return parseInput(rawContent, detectedFormat);
}

/**
 * Reads all chunks from stdin with a strict maximum byte limit.
 */
async function readStdin(maxBytes: number): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const chunks: Buffer[] = [];
    let totalBytes = 0;

    const onData = (chunk: Buffer) => {
      totalBytes += chunk.length;
      if (totalBytes > maxBytes) {
        process.stdin.removeListener('data', onData);
        process.stdin.removeListener('end', onEnd);
        process.stdin.removeListener('error', onError);
        reject(new Error(`Stdin input exceeded limit of ${maxBytes} bytes`));
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

/**
 * Parses raw JSON or CSV string content safely with prototype pollution defense
 * and strict leading-zero preservation for CSV strings.
 */
export function parseInput(
  content: string,
  format?: 'json' | 'csv',
  options?: ParseInputOptions
): unknown {
  if (typeof content !== 'string') {
    throw new Error('parseInput expects a string content');
  }

  // Strip BOM if present
  const cleanContent = content.charCodeAt(0) === 0xfeff ? content.slice(1) : content;
  const trimmed = cleanContent.trim();

  let resolvedFormat = format;
  if (!resolvedFormat) {
    if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
      resolvedFormat = 'json';
    } else {
      resolvedFormat = 'csv';
    }
  }

  if (resolvedFormat === 'json') {
    return parseJsonSafely(trimmed);
  } else if (resolvedFormat === 'csv') {
    return parseCsvSafely(cleanContent, options);
  } else {
    throw new Error(`Unknown format "${resolvedFormat}"; expected "json" or "csv"`);
  }
}

/**
 * Parses JSON safely preventing prototype pollution.
 */
export function parseJsonSafely(jsonStr: string): unknown {
  if (!jsonStr) {
    throw new Error('Cannot parse empty JSON input');
  }

  const parsed = JSON.parse(jsonStr, (key, value) => {
    if (key === '__proto__' || key === 'constructor' || key === 'prototype') {
      // Omit polluted keys
      return undefined;
    }
    return value;
  });

  return sanitizeObject(parsed);
}

/**
 * Recursively deep-sanitizes parsed objects to ensure clean prototypes.
 */
function sanitizeObject(val: unknown): unknown {
  if (val === null || typeof val !== 'object') {
    return val;
  }

  if (Array.isArray(val)) {
    return val.map((item) => sanitizeObject(item));
  }

  const cleanObj: Record<string, unknown> = Object.create(null);
  for (const [k, v] of Object.entries(val as Record<string, unknown>)) {
    if (k !== '__proto__' && k !== 'constructor' && k !== 'prototype') {
      cleanObj[k] = sanitizeObject(v);
    }
  }
  return cleanObj;
}

/**
 * Parses CSV safely with standard csv-parse and leading-zero preservation.
 */
function parseCsvSafely(
  csvStr: string,
  options?: ParseInputOptions
): Record<string, unknown>[] | (string | null)[][] {
  if (!csvStr || csvStr.trim() === '') {
    return options?.matrix ? [] : [];
  }

  const isMatrix = options?.matrix === true;
  const delimiter = options?.delimiter || ',';

  try {
    const rawRecords: unknown = parseCsvSync(csvStr, {
      columns: isMatrix ? false : true,
      skip_empty_lines: true,
      trim: false,
      cast: false, // DO NOT cast to numbers automatically - preserve leading zeros and formatting
      delimiter,
      bom: true,
      relax_column_count: !options?.strictColumnCount,
    });

    if (isMatrix) {
      if (!Array.isArray(rawRecords)) return [];
      return (rawRecords as (string | null)[][]).map((row) =>
        row.map((cell) => (cell === '' ? null : String(cell)))
      );
    }

    if (!Array.isArray(rawRecords)) {
      return [];
    }

    // Convert to sanitized record objects with clean prototype
    const result: Record<string, unknown>[] = [];
    for (const record of rawRecords as Record<string, string>[]) {
      const cleanRecord: Record<string, unknown> = Object.create(null);
      for (const [key, val] of Object.entries(record)) {
        if (key !== '__proto__' && key !== 'constructor' && key !== 'prototype') {
          // Keep string as-is (leading zeros preserved), empty string kept as "" or handled by schema
          cleanRecord[key] = val;
        }
      }
      result.push(cleanRecord);
    }

    return result;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(`CSV parse error: ${msg}`);
  }
}
