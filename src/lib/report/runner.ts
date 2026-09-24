import * as fs from 'fs';
import { GoogleSheetCli, default as GoogleSheet } from '../google-sheet';
import { XlsxSaveResult, XlsxWorkbook } from '../xlsx';
import { parseJsonSafely, readInput } from './input';
import { buildReport } from './index';
import {
  ReportDocument,
  ReportProvenance,
  ReportTemplateV1,
} from './types';

export interface ReportInputSourceOptions {
  /** File path (or '-' for stdin), or inline data string */
  input?: {
    file?: string;
    data?: string;
    format?: 'json' | 'csv';
  };
  /** Local XLSX workbook source */
  sourceWorkbook?: {
    filePath: string;
    ranges: string[];
    allowCachedFormulaValues?: boolean;
  };
  /** Cloud Google Spreadsheet source */
  sourceSpreadsheet?: {
    spreadsheetId: string;
    ranges: string[];
    gsheet: GoogleSheet;
  };
}

export interface ReportTargetOptions {
  /** Local XLSX output file */
  outputFile?: {
    filePath: string;
    workbookTemplatePath?: string;
    overwrite?: boolean;
    inPlace?: boolean;
  };
  /** Cloud Google Spreadsheet output */
  targetSpreadsheet?: {
    spreadsheetId: string;
    gsheet: GoogleSheet;
    overwrite?: boolean;
  };
}

export interface ReportRunOptions {
  /** Template definition object or file path to JSON template */
  template: ReportTemplateV1 | string;
  /** Data source configuration */
  source: ReportInputSourceOptions;
  /** Destination target configuration */
  target: ReportTargetOptions;
  /** Dry run preview mode (no mutations) */
  dryRun?: boolean;
  /** Permit overwriting existing formula cells in target */
  overwriteFormulas?: boolean;
  /** Permit saving a workbook whose unsupported features would be dropped */
  discardUnsupported?: boolean;
}

export interface ReportRunReceipt {
  templateId: string;
  templateVersion: string | number;
  sourceHash: string;
  sourceKind: 'input' | 'sourceWorkbook' | 'sourceSpreadsheet';
  targetKind: 'outputFile' | 'targetSpreadsheet';
  dryRun: boolean;
  sheetsGenerated: number;
  sheets: {
    name: string;
    rowCount: number;
    colCount: number;
  }[];
  xlsxSaved?: XlsxSaveResult;
  googleReceipt?: GoogleSheetCli.ApplyReportReceipt;
  summary?: Record<string, unknown>;
  provenance: ReportProvenance;
}

/**
 * Executes an end-to-end report automation pipeline:
 * 1. Validates options and template
 * 2. Ingests raw data from JSON/CSV/stdin/XLSX/Google
 * 3. Builds a ReportDocument synchronously via ReportCore
 * 4. Renders and applies document to XLSX file or Google Spreadsheet
 */
export async function runReport(options: ReportRunOptions): Promise<ReportRunReceipt> {
  if (!options) {
    throw new Error('ReportRunOptions must be provided');
  }

  // 1. Validate template
  const template = await resolveTemplate(options.template);

  // 2. Validate input source exclusivity
  const sourceCount =
    (options.source?.input ? 1 : 0) +
    (options.source?.sourceWorkbook ? 1 : 0) +
    (options.source?.sourceSpreadsheet ? 1 : 0);

  if (sourceCount === 0) {
    throw new Error(
      'No input source specified. Must specify exactly one of: source.input, source.sourceWorkbook, or source.sourceSpreadsheet'
    );
  }
  if (sourceCount > 1) {
    throw new Error(
      'Mutually exclusive input sources: specify only one of source.input, source.sourceWorkbook, or source.sourceSpreadsheet'
    );
  }

  // 3. Validate output target exclusivity
  const targetCount =
    (options.target?.outputFile ? 1 : 0) +
    (options.target?.targetSpreadsheet ? 1 : 0);

  if (targetCount === 0) {
    throw new Error(
      'No output target specified. Must specify exactly one of: target.outputFile or target.targetSpreadsheet'
    );
  }
  if (targetCount > 1) {
    throw new Error(
      'Mutually exclusive output targets: specify only one of target.outputFile or target.targetSpreadsheet'
    );
  }

  // Validate file collision before expensive reads/writes
  if (options.target.outputFile) {
    const outPath = options.target.outputFile.filePath;
    const inPlace = Boolean(options.target.outputFile.inPlace);

    if (options.source.sourceWorkbook && isSamePathOrFile(outPath, options.source.sourceWorkbook.filePath) && !inPlace) {
      throw new Error(
        `Output destination "${outPath}" collides with source workbook "${options.source.sourceWorkbook.filePath}". To modify in place, specify inPlace=true.`
      );
    }
    if (options.target.outputFile.workbookTemplatePath && isSamePathOrFile(outPath, options.target.outputFile.workbookTemplatePath) && !inPlace) {
      throw new Error(
        `Output destination "${outPath}" collides with workbook template "${options.target.outputFile.workbookTemplatePath}". To modify in place, specify inPlace=true.`
      );
    }
    if (options.source.input?.file && options.source.input.file !== '-' && isSamePathOrFile(outPath, options.source.input.file) && !inPlace) {
      throw new Error(
        `Output destination "${outPath}" collides with input data file "${options.source.input.file}". Specify a different output path.`
      );
    }
  }
  let rawInput: unknown;
  let sourceKind: 'input' | 'sourceWorkbook' | 'sourceSpreadsheet';

  if (options.source.input) {
    sourceKind = 'input';
    const { file, data, format } = options.source.input;
    rawInput = await readInput({ file, data, format });
  } else if (options.source.sourceWorkbook) {
    sourceKind = 'sourceWorkbook';
    const { filePath, ranges, allowCachedFormulaValues } = options.source.sourceWorkbook;
    if (!ranges || ranges.length === 0) {
      throw new Error('sourceWorkbook requires at least one range specified in "ranges"');
    }

    const wb = await XlsxWorkbook.load(filePath);
    const rangeResults = ranges.map((r) => wb.read(r, { mode: 'unformatted' }));

    // Formula freshness guard
    for (const r of rangeResults) {
      if (r.formulaCount > 0 && allowCachedFormulaValues !== true) {
        throw new Error(
          `Source workbook "${filePath}" range "${r.range}" contains ${r.formulaCount} formula cell(s) whose cache freshness cannot be guaranteed. Pass --allowCachedFormulaValues to explicitly permit using cached formula values from the workbook.`
        );
      }
    }

    if (ranges.length === 1) {
      rawInput = rangeResults[0].values;
    } else {
      const map: Record<string, unknown> = Object.create(null);
      const sheetNameCounts: Record<string, number> = {};
      for (let i = 0; i < ranges.length; i++) {
        const sheetName = rangeResults[i].sheetName;
        sheetNameCounts[sheetName] = (sheetNameCounts[sheetName] || 0) + 1;
      }
      for (let i = 0; i < ranges.length; i++) {
        map[ranges[i]] = rangeResults[i].values;
        const sheetName = rangeResults[i].sheetName;
        if (sheetNameCounts[sheetName] === 1) {
          map[sheetName] = rangeResults[i].values;
        } else {
          map[`${sheetName}_${i + 1}`] = rangeResults[i].values;
        }
      }
      map._ranges = rangeResults.map((r) => ({
        range: r.range,
        sheet: r.sheetName,
        values: r.values,
      }));
      rawInput = map;
    }
  } else if (options.source.sourceSpreadsheet) {
    sourceKind = 'sourceSpreadsheet';
    const { spreadsheetId, ranges, gsheet } = options.source.sourceSpreadsheet;
    if (!ranges || ranges.length === 0) {
      throw new Error('sourceSpreadsheet requires at least one range specified in "ranges"');
    }

    const batchResults = await gsheet.getDataBatch(
      ranges,
      { valueRenderOption: GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE },
      spreadsheetId
    );

    if (ranges.length === 1) {
      rawInput = batchResults[0].values;
    } else {
      const map: Record<string, unknown> = Object.create(null);
      for (const item of batchResults) {
        map[item.range] = item.values;
      }
      map._ranges = batchResults;
      rawInput = map;
    }
  } else {
    throw new Error('Unreachable: input source not resolved');
  }

  // 5. Build report document synchronously
  const document = buildReport(template, rawInput);

  // 6. Apply to output destination
  let targetKind: 'outputFile' | 'targetSpreadsheet';
  let xlsxSaved: XlsxSaveResult | undefined;
  let googleReceipt: GoogleSheetCli.ApplyReportReceipt | undefined;

  if (options.target.outputFile) {
    targetKind = 'outputFile';
    const { filePath, workbookTemplatePath, overwrite, inPlace } = options.target.outputFile;

    const wb = workbookTemplatePath
      ? await XlsxWorkbook.load(workbookTemplatePath)
      : XlsxWorkbook.create();

    wb.apply(document, {
      dryRun: options.dryRun,
      overwriteFormulas: options.overwriteFormulas,
      clearManagedRange: true,
    });

    if (!options.dryRun) {
      xlsxSaved = await wb.save(filePath, {
        overwrite,
        inPlace,
        allowUnsupportedFeatures: options.discardUnsupported,
      });
    }
  } else if (options.target.targetSpreadsheet) {
    targetKind = 'targetSpreadsheet';
    const { spreadsheetId, gsheet } = options.target.targetSpreadsheet;

    googleReceipt = await gsheet.applyReport(
      document,
      {
        dryRun: options.dryRun,
        overwriteFormulas: options.overwriteFormulas,
        overwrite: options.target.targetSpreadsheet.overwrite,
      },
      spreadsheetId
    );
  } else {
    throw new Error('Unreachable: output target not resolved');
  }

  return {
    templateId: document.provenance.templateId,
    templateVersion: document.provenance.templateVersion,
    sourceHash: document.provenance.sourceHash,
    sourceKind,
    targetKind,
    dryRun: Boolean(options.dryRun),
    sheetsGenerated: document.sheets.length,
    sheets: document.sheets.map((s) => ({
      name: s.name,
      rowCount: s.rows.length,
      colCount: s.rows.length > 0 ? Math.max(...s.rows.map((r) => r.length)) : 0,
    })),
    xlsxSaved,
    googleReceipt,
    summary: document.provenance.summary,
    provenance: document.provenance,
  };
}

/**
 * Resolves a template from either an in-memory object or a JSON file path.
 */
async function resolveTemplate(templateInput: ReportTemplateV1 | string): Promise<ReportTemplateV1> {
  if (!templateInput) {
    throw new Error('Report template must be provided');
  }

  if (typeof templateInput === 'string') {
    const filePath = templateInput;
    const stat = await fs.promises.stat(filePath).catch((err) => {
      throw new Error(`Cannot open template file at "${filePath}": ${err.message}`);
    });

    if (stat.size > 10 * 1024 * 1024) {
      throw new Error(`Template file size (${stat.size} bytes) exceeds limit of 10MB`);
    }

    const content = await fs.promises.readFile(filePath, 'utf8');
    const parsed = parseJsonSafely(content);
    if (!parsed || typeof parsed !== 'object') {
      throw new Error(`Template at "${filePath}" must be a valid JSON object`);
    }
    return parsed as ReportTemplateV1;
  }

  if (typeof templateInput === 'object') {
    return templateInput;
  }

  throw new Error('Invalid template input: expected template JSON file path or ReportTemplateV1 object');
}

function isSamePathOrFile(pathA?: string, pathB?: string): boolean {
  if (!pathA || !pathB) return false;
  if (pathA === pathB) return true;
  try {
    const realA = fs.realpathSync(pathA);
    const realB = fs.realpathSync(pathB);
    if (realA === realB) return true;
    const statA = fs.statSync(pathA);
    const statB = fs.statSync(pathB);
    if (statA.ino && statB.ino && statA.dev === statB.dev && statA.ino === statB.ino) {
      return true;
    }
  } catch {
    // Files may not exist yet
  }
  return false;
}
