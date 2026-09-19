/**
 * Report Core Engine Entry Point.
 * Synchronous, pure-core report generator with deterministic provenance.
 */

import * as crypto from 'crypto';
import { buildFinanceReport } from './finance';
import { buildManpowerReport } from './manpower';
import { normalizeTable } from './schema';
import { executeTransforms } from './transform';
import {
  OutputColumn,
  ReportCell,
  ReportDocument,
  ReportNumberFormat,
  ReportScalar,
  ReportSheet,
  ReportTemplateV1,
  TableSchema,
} from './types';

export * from './types';
export * from './input';
export * from './schema';
export * from './transform';
// finance.ts and manpower.ts each export a same-named toSafeNumericCell helper,
// so star re-exporting both collides; expose only the public preset builders.
// Finance/manpower value types remain public via the './types' re-export above.
export { buildFinanceReport } from './finance';
export { buildManpowerReport } from './manpower';

/**
 * Builds a ReportDocument synchronously from a template and input data.
 * Validates template shape, executes schema normalization, declarative transforms,
 * and produces output worksheets with deterministic provenance.
 */
export function buildReport(template: unknown, input: unknown): ReportDocument {
  const validatedTemplate = validateTemplateShape(template);
  const sourceHash = computeSourceHash(input);

  switch (validatedTemplate.kind) {
    case 'finance':
      return buildFinanceReport(validatedTemplate, input, sourceHash);

    case 'manpower':
      return buildManpowerReport(validatedTemplate, input, sourceHash);

    case 'table':
      return buildTableReport(validatedTemplate, input, sourceHash);

    default: {
      const unknownKind = (validatedTemplate as { kind?: string }).kind;
      throw new Error(`Unsupported template kind: "${unknownKind || 'unknown'}"; expected "table", "finance", or "manpower"`);
    }
  }
}

/**
 * Computes a deterministic SHA-256 hash of normalized input data.
 */
export function computeSourceHash(input: unknown): string {
  const canonicalString = canonicalJsonStringify(input);
  return crypto.createHash('sha256').update(canonicalString, 'utf8').digest('hex');
}

/**
 * Deterministic JSON serializer that sorts object keys recursively.
 */
function canonicalJsonStringify(val: unknown): string {
  if (val === null || val === undefined) {
    return 'null';
  }
  if (typeof val !== 'object') {
    return JSON.stringify(val);
  }
  if (Array.isArray(val)) {
    return `[${val.map((item) => canonicalJsonStringify(item)).join(',')}]`;
  }

  const obj = val as Record<string, unknown>;
  const sortedKeys = Object.keys(obj).sort();
  const pairs = sortedKeys.map((k) => `${JSON.stringify(k)}:${canonicalJsonStringify(obj[k])}`);
  return `{${pairs.join(',')}}`;
}

/**
 * Validates the runtime shape of a ReportTemplateV1.
 */
function validateTemplateShape(template: unknown): ReportTemplateV1 {
  if (!template || typeof template !== 'object' || Array.isArray(template)) {
    throw new Error('Template must be a valid JSON object');
  }

  const t = template as Record<string, unknown>;

  if (!t.id || typeof t.id !== 'string' || t.id.trim() === '') {
    throw new Error('Template "id" is required and must be a non-empty string');
  }

  if (t.version === undefined || t.version === null) {
    throw new Error('Template "version" is required (expected 1)');
  }

  const versionNum = Number(t.version);
  if (versionNum !== 1) {
    throw new Error(`Unsupported template version "${t.version}"; strictly version 1 is supported`);
  }

  if (!t.kind || typeof t.kind !== 'string') {
    throw new Error('Template "kind" is required ("table", "finance", or "manpower")');
  }

  const kind = t.kind.toLowerCase();
  if (kind !== 'table' && kind !== 'finance' && kind !== 'manpower') {
    throw new Error(`Unknown template kind "${t.kind}"; expected "table", "finance", or "manpower"`);
  }

  return {
    id: t.id.trim(),
    version: 1,
    kind: kind as 'table' | 'finance' | 'manpower',
    title: t.title ? String(t.title) : undefined,
    description: t.description ? String(t.description) : undefined,
    schema: t.schema as TableSchema | Record<string, TableSchema> | undefined,
    transforms: t.transforms as ReportTemplateV1['transforms'],
    outputs: t.outputs as ReportTemplateV1['outputs'],
    config: t.config as Record<string, unknown> | undefined,
    policy: t.policy as ReportTemplateV1['policy'],
  };
}

/**
 * Builds a generic Table ReportDocument from schema, transforms, and output mappings.
 */
export function buildTableReport(
  template: ReportTemplateV1,
  input: unknown,
  sourceHash: string
): ReportDocument {
  let datasets: Record<string, Record<string, unknown>[]> = Object.create(null);
  let mainData: Record<string, unknown>[] = [];

  // 1. Schema normalization
  if (template.schema) {
    if (isTableSchema(template.schema)) {
      // Single main table schema
      mainData = normalizeTable(input, template.schema);
      datasets['main'] = mainData;
    } else {
      // Multi-table schema dictionary
      const inputObj = (input && typeof input === 'object' && !Array.isArray(input))
        ? (input as Record<string, unknown>)
        : { main: input };

      for (const [tableName, schema] of Object.entries(template.schema)) {
        const tableRaw = inputObj[tableName];
        if (tableRaw !== undefined) {
          const normalized = normalizeTable(tableRaw, schema, tableName);
          datasets[tableName] = normalized;
          if (tableName === 'main' || mainData.length === 0) {
            mainData = normalized;
          }
        }
      }
    }
  } else {
    // No schema specified - accept raw array of records
    if (Array.isArray(input)) {
      mainData = input as Record<string, unknown>[];
      datasets['main'] = mainData;
    } else if (input && typeof input === 'object') {
      const inputObj = input as Record<string, unknown>;
      for (const [k, v] of Object.entries(inputObj)) {
        if (Array.isArray(v)) {
          datasets[k] = v as Record<string, unknown>[];
        }
      }
      mainData = datasets['main'] || (Object.values(datasets)[0] || []);
    } else {
      throw new Error('Input data must be an array of records or an object containing tables');
    }
  }

  // 2. Declarative transforms
  const context: Record<string, unknown> = {
    ...datasets,
    config: template.config || {},
  };

  const processedData = executeTransforms(mainData, template.transforms, context);

  // 3. Output sheets generation
  const sheets: ReportSheet[] = [];

  if (template.outputs && template.outputs.length > 0) {
    for (const outDef of template.outputs) {
      const sourceKey = outDef.sourceDataset || 'main';
      const sheetData = sourceKey === 'main' ? processedData : (datasets[sourceKey] || processedData);
      const sheet = buildOutputWorksheet(outDef, sheetData);
      sheets.push(sheet);
    }
  } else {
    // Default sheet from processedData
    sheets.push(buildDefaultWorksheet(template.title || 'Report', processedData));
  }

  return {
    sheets,
    provenance: {
      templateId: template.id,
      templateVersion: template.version,
      sourceHash,
      generatedAt: new Date().toISOString(),
      summary: {
        totalRows: processedData.length,
        sheetCount: sheets.length,
      },
    },
  };
}

function isTableSchema(schema: unknown): schema is TableSchema {
  return typeof schema === 'object' && schema !== null && Array.isArray((schema as TableSchema).fields);
}

/**
 * Validates that a precomputed formula result is a valid ReportScalar.
 * Rejects nested formula objects, functions, arrays, or non-scalar objects.
 */
export function validateReportScalarResult(val: unknown, context: string): ReportScalar {
  if (val === null || val === undefined) {
    return null;
  }
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return val;
  }
  if (val instanceof Date) {
    return val.toISOString();
  }
  throw new Error(
    `Formula result for ${context} must be a scalar (string, number, boolean, null, or Date), received non-scalar: ${
      typeof val === 'object' ? JSON.stringify(val) : String(val)
    }`
  );
}

function validateReportCell(val: unknown, context: string): ReportCell {
  if (val === null || val === undefined) return null;
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return val;
  if (val instanceof Date) return val.toISOString();
  if (typeof val === 'object' && val !== null && 'formula' in val && typeof (val as { formula: unknown }).formula === 'string') {
    const fObj = val as { formula: string; result?: unknown };
    const cleanFormula = fObj.formula.startsWith('=') ? fObj.formula.slice(1) : fObj.formula;
    return {
      formula: cleanFormula,
      result: fObj.result !== undefined ? validateReportScalarResult(fObj.result, `${context} cell formula "${cleanFormula}"`) : undefined,
    };
  }
  return String(val);
}

/**
 * Builds a single ReportSheet according to an output definition.
 */
function buildOutputWorksheet(
  outDef: NonNullable<ReportTemplateV1['outputs']>[0],
  data: Record<string, unknown>[]
): ReportSheet {
  const headers: ReportCell[] = outDef.columns.map((c) => c.header);
  const rows: ReportCell[][] = [headers];

  for (let i = 0; i < data.length; i++) {
    const rowObj = data[i];
    const excelRow = i + 2; // 1-based index in Excel, row 1 is header
    const rowCells: ReportCell[] = [];

    for (const col of outDef.columns) {
      if (col.formulaTemplate) {
        let formulaStr = col.formulaTemplate.replace(/\{row\}/g, String(excelRow));
        if (formulaStr.startsWith('=')) {
          formulaStr = formulaStr.slice(1);
        }

        let resultVal: unknown = undefined;
        if (col.resultField && rowObj[col.resultField] !== undefined) {
          resultVal = rowObj[col.resultField];
        } else if (col.field && rowObj[col.field] !== undefined) {
          resultVal = rowObj[col.field];
        }

        let evaluatedResult: ReportScalar | undefined = undefined;
        if (resultVal !== undefined) {
          evaluatedResult = validateReportScalarResult(
            resultVal,
            `column "${col.header}" row ${excelRow} formula "${formulaStr}"`
          );
        }

        rowCells.push({
          formula: formulaStr,
          result: evaluatedResult,
        });
      } else if (col.field) {
        const val = rowObj[col.field];
        rowCells.push(val === undefined ? null : validateReportCell(val, `column "${col.header}" row ${excelRow}`));
      } else {
        rowCells.push(null);
      }
    }

    rows.push(rowCells);
  }

  // Summary rows
  if (outDef.summaryRows && outDef.summaryRows.length > 0) {
    for (const sRow of outDef.summaryRows) {
      const summaryCells: ReportCell[] = [];
      if (sRow.label !== undefined) {
        summaryCells.push(sRow.label);
      }

      for (const c of sRow.cells) {
        if (c.formula) {
          let formulaStr = c.formula;
          if (formulaStr.startsWith('=')) {
            formulaStr = formulaStr.slice(1);
          }
          let evaluatedResult: ReportScalar | undefined = undefined;
          const rawResult = c.result !== undefined ? c.result : (c.value !== undefined ? c.value : undefined);
          if (rawResult !== undefined) {
            evaluatedResult = validateReportScalarResult(
              rawResult,
              `summary row formula "${formulaStr}"`
            );
          }
          summaryCells.push({
            formula: formulaStr,
            result: evaluatedResult,
          });
        } else if (c.value !== undefined) {
          summaryCells.push(validateReportScalarResult(c.value, 'summary row value'));
        } else {
          summaryCells.push(null);
        }
      }
      rows.push(summaryCells);
    }
  }

  // Formats and widths
  const columnWidths = outDef.columns.map((c) => c.width || 18);
  const numberFormats: ReportNumberFormat[] = [];
  for (let c = 0; c < outDef.columns.length; c++) {
    const col = outDef.columns[c];
    if (col.format) {
      numberFormats.push({ column: c + 1, format: col.format });
    }
  }

  return {
    name: outDef.name,
    rows,
    freezeRows: outDef.freezeRows !== undefined ? outDef.freezeRows : 1,
    columnWidths,
    numberFormats,
    startCell: outDef.startCell,
    targetNamedRange: outDef.targetNamedRange,
    clearManagedRange: outDef.clearManagedRange,
  };
}

/**
 * Builds a fallback default worksheet when no explicit output columns are defined.
 */
function buildDefaultWorksheet(
  name: string,
  data: Record<string, unknown>[]
): ReportSheet {
  if (data.length === 0) {
    return {
      name,
      rows: [['No Data']],
      freezeRows: 1,
    };
  }

  const columns = Object.keys(data[0]);
  const headers: ReportCell[] = [...columns];
  const rows: ReportCell[][] = [headers];

  for (const row of data) {
    const rowCells: ReportCell[] = columns.map((col) => {
      const v = row[col];
      return v === undefined ? null : (v as ReportCell);
    });
    rows.push(rowCells);
  }

  return {
    name,
    rows,
    freezeRows: 1,
  };
}
