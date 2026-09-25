import * as fs from 'fs';
import * as path from 'path';
import ExcelJS from 'exceljs';
import {
  ReportCell,
  ReportDocument,
  ReportFormula,
  ReportNumberFormat,
  ReportScalar,
  ReportSheet,
  XlsxApplyOptions,
  XlsxApplyResult,
  XlsxCellFreshness,
  XlsxCellReadInfo,
  XlsxCellSnapshot,
  XlsxCreateOptions,
  XlsxClearOptions,
  XlsxClearResult,
  XlsxDefinedName,
  XlsxDiffAction,
  XlsxDiffItem,
  XlsxDiffResult,
  XlsxFreshnessSummary,
  XlsxFindMatch,
  XlsxFindOptions,
  XlsxFindResult,
  XlsxFreezeOptions,
  XlsxFreezeResult,
  XlsxFormatCellsOptions,
  XlsxFormatCellsResult,
  XlsxAddNameOptions,
  XlsxAddNameResult,
  XlsxAddSheetResult,
  XlsxHideOptions,
  XlsxHideResult,
  XlsxInspection,
  XlsxLoadOptions,
  XlsxInspectOptions,
  XlsxManagedExtent,
  XlsxMergeOptions,
  XlsxMergeResult,
  XlsxPreflightResult,
  XlsxReadMode,
  XlsxReadOptions,
  XlsxReadResult,
  XlsxRemoveNameOptions,
  XlsxRemoveNameResult,
  XlsxRemoveSheetResult,
  XlsxRenameSheetResult,
  XlsxResizeOptions,
  XlsxResizeResult,
  XlsxResizeTarget,
  XlsxSaveOptions,
  XlsxSaveResult,
  XlsxSheetMetadata,
  XlsxSheetOpOptions,
  XlsxSetCellsOptions,
  XlsxSetCellsResult,
  XlsxSpliceOptions,
  XlsxSpliceResult,
} from './xlsx-types';
import {
  calculateTargetRange,
  cellAddress,
  formatA1Range,
  parseA1Range,
  parseCellAddress,
  colLetterToIndex,
  indexToColLetter,
  MAX_EXCEL_COLS,
  MAX_EXCEL_ROWS,
  ParsedA1Range,
} from './xlsx-range';
import {
  diffDefinedNameRanges,
  rewriteFormulaRefs,
  XlsxRefSpliceOp,
} from './xlsx-refs';
import { inspectZipBuffer } from './xlsx-zip';
import { computeFileSha256, saveBufferAtomic } from './xlsx-file';

export * from './xlsx-types';
export * from './xlsx-range';
export * from './xlsx-zip';
export * from './xlsx-file';

const MANAGED_RANGE_PREFIX = '_GS_MANAGED_';

const XLSX_DEFAULT_ROW_HEIGHT = 15;
const XLSX_DEFAULT_COL_WIDTH = 9;
const XLSX_AUTO_WIDTH_PADDING = 2;
const XLSX_MAX_AUTO_WIDTH = 60;
const XLSX_MAX_COL_WIDTH = 255;
const XLSX_MAX_ROW_HEIGHT = 409.5;

/** style keys formatCells accepts; --clear refuses combinations with any of them */
const XLSX_FORMAT_KEYS = [
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'fontSize',
  'fontFamily',
  'textColor',
  'backgroundColor',
  'horizontalAlignment',
  'verticalAlignment',
  'wrapText',
  'numberFormat',
  'borders',
] as const;

/** Sheets-side border style names mapped onto ExcelJS line styles; NONE removes a side */
const XLSX_BORDER_STYLE_MAP: Record<string, ExcelJS.BorderStyle | undefined> = {
  DOTTED: 'dotted',
  DASHED: 'dashed',
  SOLID: 'thin',
  SOLID_MEDIUM: 'medium',
  SOLID_THICK: 'thick',
  DOUBLE: 'double',
  NONE: undefined,
};

/** refersTo must be a sheet-quoted or bare-title qualified A1 range; ExcelJS's own add() accepts anything */
const XLSX_REFERS_TO_RE = /^(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

type XlsxBorderSide = 'top' | 'bottom' | 'left' | 'right' | 'innerHorizontal' | 'innerVertical';

/** Style spec resolved once from flags, then applied to every cell of the range. */
interface PreparedCellStyle {
  clear: boolean;
  /** style keys that were applied, used for receipts and dry-run previews */
  applied: string[];
  font?: Partial<ExcelJS.Font>;
  fill?: ExcelJS.Fill;
  alignment?: Partial<ExcelJS.Alignment>;
  numFmt?: string;
  borderSides?: Set<XlsxBorderSide>;
  /** undefined means the addressed sides are removed (borderStyle NONE) */
  borderStyle?: ExcelJS.BorderStyle;
  borderColor?: { argb: string };
}

interface CellExtraction {
  scalarValue: ReportScalar;
  formattedText?: string;
  formulaString?: string;
  cachedResult?: ReportScalar;
  isFormula: boolean;
  freshness: XlsxCellFreshness;
  hasStyle: boolean;
  numberFormat?: string;
}

/**
 * XlsxWorkbook adapter providing local ExcelJS report generation, inspection,
 * diffing, range reading with formula freshness tracking, and secure atomic persistence.
 */
export class XlsxWorkbook {
  public readonly workbook: ExcelJS.Workbook;
  public readonly filePath?: string;
  public readonly fileSize?: number;
  public readonly originalHash?: string;
  public readonly preflight?: XlsxPreflightResult;
  private readonly _managedExtents: Map<string, string> = new Map();

  /**
   * Loads an XLSX file from disk with preflight checks and hash validation.
   */
  public static async load(
    filePath: string,
    options?: XlsxLoadOptions
  ): Promise<XlsxWorkbook> {
    if (!filePath || typeof filePath !== 'string' || !filePath.trim()) {
      throw new Error('XLSX file path must be a non-empty string.');
    }

    const resolvedPath = path.resolve(filePath);
    const stats = await fs.promises.stat(resolvedPath).catch((err) => {
      throw new Error(`Cannot open XLSX file at "${resolvedPath}": ${err.message}`);
    });

    if (!stats.isFile()) {
      throw new Error(`Path "${resolvedPath}" is not a regular file.`);
    }

    const maxFileSize = options?.maxFileSize ?? 50 * 1024 * 1024;
    if (stats.size > maxFileSize) {
      throw new Error(
        `XLSX file size (${stats.size} bytes) exceeds limit of ${maxFileSize} bytes.`
      );
    }

    const buffer = await fs.promises.readFile(resolvedPath);
    const originalHash = computeFileSha256(resolvedPath).catch(() => undefined);
    const preflight = await inspectZipBuffer(buffer, options);

    const wb = new ExcelJS.Workbook();
    try {
      // Type boundary: ExcelJS declares load(buffer: Buffer) with its own ambient
      // Buffer typing that does not unify with Node's fs Buffer, although the
      // runtime accepts Node Buffer. Assert through unknown to load's declared
      // parameter type instead of copying the buffer.
      await wb.xlsx.load(buffer as unknown as Parameters<typeof wb.xlsx.load>[0]);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(`ExcelJS failed to load workbook from "${resolvedPath}": ${msg}`);
    }

    const instance = new XlsxWorkbook(wb, {
      filePath: resolvedPath,
      fileSize: stats.size,
      originalHash: await originalHash,
      preflight,
    });

    instance._loadManagedExtentsFromDefinedNames();
    return instance;
  }

  /**
   * Creates a new empty XlsxWorkbook.
   */
  public static create(options?: XlsxCreateOptions): XlsxWorkbook {
    const wb = new ExcelJS.Workbook();
    if (options?.creator) wb.creator = options.creator;
    if (options?.created) wb.created = options.created;
    if (options?.modified) wb.modified = options.modified;

    return new XlsxWorkbook(wb, {});
  }

  public constructor(
    workbook?: ExcelJS.Workbook,
    context?: {
      filePath?: string;
      fileSize?: number;
      originalHash?: string;
      preflight?: XlsxPreflightResult;
    }
  ) {
    this.workbook = workbook || new ExcelJS.Workbook();
    this.filePath = context?.filePath;
    this.fileSize = context?.fileSize;
    this.originalHash = context?.originalHash;
    this.preflight = context?.preflight;
  }

  /**
   * Inspects metadata, sheets, defined names, preflight results, and capabilities.
   */
  public inspect(options?: XlsxInspectOptions): XlsxInspection {
    const sheets: XlsxSheetMetadata[] = this.workbook.worksheets.map((ws) => {
      let formulaCellCount = 0;
      const formulaCells: string[] = [];
      ws.eachRow((row) => {
        row.eachCell((cell) => {
          if (
            cell.type === ExcelJS.ValueType.Formula ||
            (cell.value && typeof cell.value === 'object' && 'formula' in cell.value)
          ) {
            formulaCellCount++;
            if (options?.includeFormulaCells) formulaCells.push(cell.address);
          }
        });
      });

      const mergedRanges = ((ws.model.merges as string[] | undefined) ?? []).slice().sort();

      const frozenView = (ws.views ?? []).find(
        (v) => v.state === 'frozen' && (((v as { xSplit?: number }).xSplit ?? 0) > 0 || ((v as { ySplit?: number }).ySplit ?? 0) > 0)
      ) as { xSplit?: number; ySplit?: number } | undefined;
      const frozen = frozenView
        ? { rows: frozenView.ySplit ?? 0, columns: frozenView.xSplit ?? 0 }
        : undefined;

      // dataValidations and conditionalFormattings exist at runtime but are not
      // on the public Worksheet typing - same boundary-cast pattern as definedNames.
      const dataValidations = Object.entries(
        (((ws as unknown as { dataValidations?: { model?: Record<string, { type?: string }> } }).dataValidations)?.model) ?? {}
      ).map(([range, rule]) => ({ range, type: rule?.type }));

      const conditionalFormatting = (((ws as unknown as { conditionalFormattings?: { ref: string; rules?: unknown[] }[] }).conditionalFormattings) ?? []).map((cf) => ({
        range: cf.ref,
        ruleCount: Array.isArray(cf.rules) ? cf.rules.length : 0,
      }));

      const managedRange = this._managedExtents.get(ws.name);

      return {
        id: Number(ws.id),
        name: ws.name,
        state: (ws.state as 'visible' | 'hidden' | 'veryHidden') || 'visible',
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
        hasFormulas: formulaCellCount > 0,
        formulaCellCount,
        ...(options?.includeFormulaCells ? { formulaCells } : {}),
        mergedRanges,
        ...(frozen ? { frozen } : {}),
        ...(dataValidations.length > 0 ? { dataValidations } : {}),
        ...(conditionalFormatting.length > 0 ? { conditionalFormatting } : {}),
        managedRange,
      };
    });

    const definedNames: XlsxDefinedName[] = this._readDefinedNames();

    const hasUnsupported = this.preflight?.hasUnsupportedFeatures ?? false;
    const unsupportedFeatures = this.preflight?.unsupportedFeatures ?? [];

    const warnings: string[] = [...(this.preflight?.warnings ?? [])];
    if (hasUnsupported) {
      warnings.push(
        `Workbook contains ${unsupportedFeatures.length} unsupported feature(s). Modifying and saving without explicit override will discard these features.`
      );
    }

    return {
      filePath: this.filePath,
      fileSize: this.fileSize,
      sha256Hash: this.originalHash,
      sheetCount: sheets.length,
      sheets,
      definedNames,
      preflight: this.preflight || {
        hasUnsupportedFeatures: false,
        unsupportedFeatures: [],
        sheetNames: sheets.map((s) => s.name),
        entryCount: 0,
        totalUncompressedSize: 0,
        warnings: [],
      },
      hasUnsupportedFeatures: hasUnsupported,
      unsupportedFeatures,
      capabilities: {
        canRead: true,
        canWriteSafely: !hasUnsupported,
        canRecalculateFormulas: false, // ExcelJS cannot evaluate arbitrary formulas
        preservesVba: false,
        preservesCharts: false,
        preservesPivotTables: false,
      },
      warnings,
    };
  }

  /**
   * Reads a bounded A1 range with explicit formula-cache freshness metadata.
   */
  public read(rangeStr: string, options?: XlsxReadOptions): XlsxReadResult {
    const parsed = parseA1Range(rangeStr, {
      maxRows: options?.maxRows,
      maxCols: options?.maxCols,
      defaultSheet: options?.defaultSheet,
    });

    // Locate target worksheet
    let ws: ExcelJS.Worksheet | undefined;
    if (parsed.sheetName) {
      ws = this.workbook.getWorksheet(parsed.sheetName);
    } else if (options?.defaultSheet) {
      ws = this.workbook.getWorksheet(options.defaultSheet);
    } else if (this.workbook.worksheets.length > 0) {
      ws = this.workbook.worksheets[0];
    }

    if (!ws) {
      const targetName = parsed.sheetName || options?.defaultSheet || 'Default';
      const available = this.workbook.worksheets.map((w) => `"${w.name}"`).join(', ');
      throw new Error(
        `Worksheet "${targetName}" not found in workbook. Available sheets: [${available}].`
      );
    }

    const sheetName = ws.name;
    const mode = options?.mode || 'unformatted';

    // Bound full row / full column ranges to actual sheet dimensions
    let endRow = parsed.endRow;
    let endCol = parsed.endCol;

    if (parsed.isFullCol) {
      const actualRows = Math.max(ws.actualRowCount || 0, ws.rowCount || 0, 1);
      endRow = Math.min(actualRows, options?.maxRows ?? 100000);
    }
    if (parsed.isFullRow) {
      const actualCols = Math.max(ws.actualColumnCount || 0, ws.columnCount || 0, 1);
      endCol = Math.min(actualCols, options?.maxCols ?? 16384);
    }

    const values: ReportCell[][] = [];
    const cellDetails: XlsxCellReadInfo[][] = [];

    const freshnessSummary: XlsxFreshnessSummary = {
      cachedCount: 0,
      unknownCount: 0,
      emptyCount: 0,
      notFormulaCount: 0,
    };
    let formulaCount = 0;

    for (let r = parsed.startRow; r <= endRow; r++) {
      const rowValues: ReportCell[] = [];
      const rowDetails: XlsxCellReadInfo[] = [];

      for (let c = parsed.startCol; c <= endCol; c++) {
        const cell = ws.getCell(r, c);
        const extracted = this._extractCellData(cell);

        if (extracted.isFormula) {
          formulaCount++;
        }

        switch (extracted.freshness) {
          case 'cached':
            freshnessSummary.cachedCount++;
            break;
          case 'unknown':
            freshnessSummary.unknownCount++;
            break;
          case 'empty':
            freshnessSummary.emptyCount++;
            break;
          case 'not-formula':
            freshnessSummary.notFormulaCount++;
            break;
        }

        const address = cellAddress(r, c);
        const readInfo: XlsxCellReadInfo = {
          address,
          row: r,
          col: c,
          value: extracted.scalarValue,
          formattedText: extracted.formattedText,
          formula: extracted.formulaString,
          cachedResult: extracted.cachedResult,
          freshness: extracted.freshness,
          numberFormat: extracted.numberFormat,
          hasStyle: extracted.hasStyle,
        };
        rowDetails.push(readInfo);

        // Map output cell according to requested mode
        if (mode === 'formula' && extracted.isFormula && extracted.formulaString) {
          rowValues.push({
            formula: extracted.formulaString,
            result: extracted.cachedResult,
          });
        } else if (mode === 'formatted' && extracted.formattedText !== undefined) {
          rowValues.push(extracted.formattedText);
        } else {
          rowValues.push(extracted.scalarValue);
        }
      }

      values.push(rowValues);
      cellDetails.push(rowDetails);
    }

    const resolvedRange = formatA1Range({
      sheetName,
      startCol: parsed.startCol,
      startRow: parsed.startRow,
      endCol,
      endRow,
    });

    return {
      range: resolvedRange,
      sheetName,
      mode,
      values,
      cellDetails,
      formulaCount,
      freshnessSummary,
      rowCount: values.length,
      colCount: values.length > 0 ? values[0].length : 0,
    };
  }

  /**
   * Scans a bounded range for cells matching a condition and returns their A1
   * coordinates - the local counterpart of GoogleSheet.findData with the same
   * result shape, so callers can target follow-up writes without downloading
   * the whole sheet.
   */
  public find(options: XlsxFindOptions): XlsxFindResult {
    const modes = [options.equals !== undefined, options.contains !== undefined, options.regex !== undefined].filter(Boolean).length;
    if (modes !== 1) {
      throw new Error('Exactly one match mode is required: equals, contains or regex');
    }

    const ws = this._resolveWorksheet(options.worksheetTitle);

    // Resolve the scan range: explicit --range wins, otherwise the used extent.
    let rangeStr = options.range;
    if (!rangeStr) {
      const endRow = Math.max(ws.actualRowCount || 0, ws.rowCount || 0, 1);
      const endCol = Math.max(ws.actualColumnCount || 0, ws.columnCount || 0, 1);
      rangeStr = formatA1Range({ sheetName: ws.name, startCol: 1, startRow: 1, endCol, endRow });
    }
    const read = this.read(rangeStr, { mode: 'formatted', defaultSheet: ws.name });
    const parsed = parseA1Range(rangeStr, { defaultSheet: ws.name });

    const ignoreCase = options.ignoreCase !== false;
    const needle = options.equals ?? options.contains;
    const foldedNeedle = ignoreCase && needle !== undefined ? needle.toLowerCase() : needle;
    const regex = options.regex !== undefined ? new RegExp(options.regex, ignoreCase ? 'i' : '') : undefined;

    // Restrict to one column when --column or --header is given. --header resolves
    // against the first scanned row, matching the Sheets-side semantics.
    let restrictCol: number | undefined;
    if (options.column) {
      restrictCol = colLetterToIndex(options.column);
    } else if (options.header !== undefined) {
      const headerNeedle = ignoreCase ? options.header.toLowerCase() : options.header;
      const headerRow = read.cellDetails[0] ?? [];
      for (const cell of headerRow) {
        const text = cell.formattedText ?? (cell.value !== null ? String(cell.value) : '');
        const folded = ignoreCase ? text.toLowerCase() : text;
        if (folded === headerNeedle) {
          restrictCol = cell.col;
          break;
        }
      }
      if (restrictCol === undefined) {
        throw new Error(`Header "${options.header}" not found in the first row of ${read.range}.`);
      }
    }

    const all: XlsxFindMatch[] = [];
    const seenRows = new Set<number>();
    for (const rowDetails of read.cellDetails) {
      for (const cell of rowDetails) {
        if (restrictCol !== undefined && cell.col !== restrictCol) continue;
        if (options.byRow && seenRows.has(cell.row)) continue;

        const text = cell.formattedText ?? (cell.value !== null ? String(cell.value) : '');
        const folded = ignoreCase ? text.toLowerCase() : text;
        const matched =
          regex !== undefined
            ? regex.test(text)
            : options.equals !== undefined
              ? folded === foldedNeedle
              : folded.includes(foldedNeedle ?? '');
        if (!matched) continue;

        if (options.byRow) seenRows.add(cell.row);
        all.push({
          a1: cell.address,
          row: cell.row,
          column: cell.col,
          columnLetter: indexToColLetter(cell.col),
          value: text,
          ...(options.byRow
            ? {
                // raw cell values, matching the Sheets-side rowValues contract
                rowValues: (read.cellDetails[cell.row - parsed.startRow] ?? []).map((c) => c.value),
              }
            : {}),
        });
      }
    }

    const limit = options.limit ?? 100;
    return {
      range: read.range,
      matchCount: all.length,
      truncated: all.length > limit,
      matches: all.slice(0, limit),
    };
  }

  /**
   * Inserts or deletes rows/columns on a worksheet. Merges are managed explicitly:
   * ExcelJS's own splice remerge is unreliable at the boundary, so every merge
   * touching the affected region is unmerged first, the splice runs, then merges
   * are re-applied at their adjusted coordinates. A merge intersecting the splice
   * boundary refuses the operation unless force is passed - with force, an insert
   * extends the merge over the new rows/columns and a delete shrinks it.
   *
   * Formula references are left stale by default: the receipt reports how many
   * formula cells reference the shifted region so the caller can verify them.
   * With updateRefs, same-sheet references in every formula of the worksheet
   * are rewritten against the splice (refs inside a deleted span become #REF!)
   * and defined-name shifts are counted - cross-sheet references are
   * deliberately untouched (phase 1 of F7).
   */
  public splice(operation: 'insert' | 'delete', options: XlsxSpliceOptions): XlsxSpliceResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const dimension = options.dimension;
    const start = options.start;
    const count = options.count ?? 1;
    const warnings: string[] = [];
    const willRewriteRefs = options.updateRefs === true && options.dryRun !== true;
    // ExcelJS's own spliceRows/spliceColumns shift defined names on the spliced
    // sheet as a side effect; snapshot the ranges beforehand so the receipt can
    // report the change without shifting them twice.
    const definedNamesBefore = options.updateRefs === true ? this._definedNameSignature() : undefined;

    if (!Number.isInteger(start) || start < 1) {
      throw new Error(`--start must be a positive 1-based index, got ${start}.`);
    }
    if (!Number.isInteger(count) || count < 1) {
      throw new Error(`--count must be a positive integer, got ${count}.`);
    }

    const isRows = dimension === 'ROWS';
    const spanEnd = start + count - 1;

    // Classify merges against the splice boundary.
    const merges = this._mergeBounds(ws);
    const mergeConflicts: string[] = [];
    const mergesAdjusted: { before: string; after?: string }[] = [];
    const unmergeList: { top: number; left: number; bottom: number; right: number }[] = [];

    for (const m of merges) {
      const lo = isRows ? m.top : m.left;
      const hi = isRows ? m.bottom : m.right;
      const rangeStr = formatA1Range({ startCol: m.left, startRow: m.top, endCol: m.right, endRow: m.bottom });

      if (hi < start) continue; // entirely before the splice point: untouched

      if (operation === 'insert') {
        if (lo >= start) {
          // entirely at/after: shifts down/right by count
          const adjusted = isRows
            ? { ...m, top: m.top + count, bottom: m.bottom + count }
            : { ...m, left: m.left + count, right: m.right + count };
          unmergeList.push(m);
          mergesAdjusted.push({ before: rangeStr, after: this._boundsToA1(adjusted) });
        } else {
          // straddles the insertion point
          if (!options.force) {
            mergeConflicts.push(rangeStr);
            continue;
          }
          const adjusted = isRows ? { ...m, bottom: m.bottom + count } : { ...m, right: m.right + count };
          unmergeList.push(m);
          mergesAdjusted.push({ before: rangeStr, after: this._boundsToA1(adjusted) });
        }
      } else {
        // delete: [start, spanEnd] is removed
        if (lo > spanEnd) {
          // entirely after the deleted span: shifts up/left
          const adjusted = isRows
            ? { ...m, top: m.top - count, bottom: m.bottom - count }
            : { ...m, left: m.left - count, right: m.right - count };
          unmergeList.push(m);
          mergesAdjusted.push({ before: rangeStr, after: this._boundsToA1(adjusted) });
        } else {
          // intersects the deleted span
          if (!options.force) {
            mergeConflicts.push(rangeStr);
            continue;
          }
          const overlap = Math.min(hi, spanEnd) - Math.max(lo, start) + 1;
          // shrink by the overlap; a merge starting inside the span re-anchors at start
          const finalLo = lo >= start ? start : lo;
          const finalHi = hi - overlap;
          const adjusted = isRows
            ? { ...m, top: finalLo, bottom: finalHi }
            : { ...m, left: finalLo, right: finalHi };
          unmergeList.push(m);
          // a merge survives shrinking unless it collapses to a single cell or less
          const collapses = finalHi <= finalLo && m.left === m.right;
          if (finalHi >= finalLo && !collapses) {
            mergesAdjusted.push({ before: rangeStr, after: this._boundsToA1(adjusted) });
          } else {
            mergesAdjusted.push({ before: rangeStr }); // merge collapsed to a single cell: dropped
          }
        }
      }
    }

    if (mergeConflicts.length > 0) {
      throw new Error(
        `Splice blocked: ${mergeConflicts.length} merged range(s) intersect the ${dimension.toLowerCase()} boundary at ${start}: [${mergeConflicts.join(', ')}]. Pass --force to adjust them.`
      );
    }

    // Formula cells referencing the shifted region are never rewritten unless
    // updateRefs was requested; count them so the receipt can warn honestly.
    const formulasAtRisk = this._countFormulasReferencing(ws, start, isRows);
    if (formulasAtRisk > 0 && !willRewriteRefs) {
      warnings.push(
        `${formulasAtRisk} formula cell(s) on "${ws.name}" may reference the shifted region; formula references are not rewritten - verify them after the splice.`
      );
    }
    warnings.push('Data validations and conditional formatting ranges are not adjusted by splice.');

    // Capture values about to be removed (delete dry-run parity with the cloud command).
    let removedValues: ReportCell[][] | undefined;
    if (operation === 'delete') {
      removedValues = [];
      if (isRows) {
        for (let r = start; r <= spanEnd; r++) {
          const rowVals: ReportCell[] = [];
          const row = ws.getRow(r);
          row.eachCell({ includeEmpty: true }, (cell) => {
            rowVals.push(this._extractCellData(cell).scalarValue);
          });
          removedValues.push(rowVals);
        }
      } else {
        const maxRow = Math.max(ws.actualRowCount || 0, 1);
        for (let r = 1; r <= maxRow; r++) {
          const rowVals: ReportCell[] = [];
          for (let c = start; c <= spanEnd; c++) {
            rowVals.push(this._extractCellData(ws.getCell(r, c)).scalarValue);
          }
          if (rowVals.some((v) => v !== null)) removedValues.push(rowVals);
        }
      }
    }

    if (options.dryRun) {
      return {
        operation,
        sheet: ws.name,
        dimension,
        start,
        count,
        mergesAdjusted,
        mergeConflicts,
        formulasAtRisk,
        refsRewritten: 0,
        refsBroken: 0,
        removedValues,
        dryRun: true,
        warnings,
      };
    }

    // Snapshot every formula cell's effective text BEFORE the physical splice.
    // Shared-formula slave cells translate the master's text live through their
    // offset, so reading them after the master has been rewritten would
    // double-shift them - and ExcelJS's splice copies values into fresh cell
    // objects anyway, so the snapshot is keyed by pre-splice address.
    const formulaSnapshot = willRewriteRefs ? this._snapshotFormulaTexts(ws) : undefined;

    // Unmerge everything affected BEFORE splicing so ExcelJS's own remerge logic
    // (which mishandles boundary-straddling merges) has nothing left to touch.
    for (const m of unmergeList) {
      ws.unMergeCells(m.top, m.left, m.bottom, m.right);
    }

    if (isRows) {
      if (operation === 'insert') {
        const blanks = new Array(count).fill([]);
        ws.spliceRows(start, 0, ...blanks);
        if (options.inheritFromBefore && start > 1) {
          this._copyRowStyle(ws, start - 1, start, count);
        }
      } else {
        ws.spliceRows(start, count);
      }
    } else {
      if (operation === 'insert') {
        const blanks = new Array(count).fill([]);
        ws.spliceColumns(start, 0, ...blanks);
        if (options.inheritFromBefore && start > 1) {
          this._copyColStyle(ws, start - 1, start, count);
        }
      } else {
        ws.spliceColumns(start, count);
      }
    }

    // Re-apply adjusted merges.
    for (const adj of mergesAdjusted) {
      if (!adj.after) continue;
      const p = parseA1Range(adj.after);
      ws.mergeCells(p.startRow, p.startCol, p.endRow, p.endCol);
    }

    // The snapshot holds pre-splice coordinates (the physical splice moves
    // values but never touches formula strings), so rewriting against the
    // original operation parameters and re-resolving each entry's spliced
    // address is correct here.
    let refsRewritten = 0;
    let refsBroken = 0;
    if (willRewriteRefs) {
      const op: XlsxRefSpliceOp = { dimension, start, count, mode: operation };
      const cellRefs = this._rewriteFormulaRefs(ws, formulaSnapshot!, op, isRows);
      refsRewritten = cellRefs.rewritten;
      refsBroken = cellRefs.refErrors;
      if (definedNamesBefore) {
        const nameDiff = diffDefinedNameRanges(definedNamesBefore, this._definedNameSignature());
        refsRewritten += nameDiff.rewritten;
        refsBroken += nameDiff.broken;
      }
    }

    return {
      operation,
      sheet: ws.name,
      dimension,
      start,
      count,
      mergesAdjusted,
      mergeConflicts,
      formulasAtRisk,
      refsRewritten,
      refsBroken,
      removedValues,
      dryRun: false,
      warnings,
    };
  }

  /**
   * Merges or unmerges cells over a bounded range. MERGE_COLUMNS and MERGE_ROWS
   * decompose into one merge per column/row, matching the Sheets-side semantics;
   * the top-left value is kept and other values in the range are hidden.
   */
  public mergeCellsRange(options: XlsxMergeOptions): XlsxMergeResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const parsed = parseA1Range(options.range, { defaultSheet: ws.name });
    if (parsed.isFullCol || parsed.isFullRow) {
      throw new Error(`Merge range "${options.range}" must be bounded on both axes.`);
    }

    const mergeType = options.mergeType ?? 'MERGE_ALL';
    const targets: { top: number; left: number; bottom: number; right: number }[] = [];
    if (mergeType === 'MERGE_ALL') {
      targets.push({ top: parsed.startRow, left: parsed.startCol, bottom: parsed.endRow, right: parsed.endCol });
    } else if (mergeType === 'MERGE_COLUMNS') {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        targets.push({ top: parsed.startRow, left: c, bottom: parsed.endRow, right: c });
      }
    } else {
      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        targets.push({ top: r, left: parsed.startCol, bottom: r, right: parsed.endCol });
      }
    }

    const affected = targets.map((t) => this._boundsToA1(t));

    if (!options.dryRun) {
      if (options.unmerge) {
        for (const t of targets) {
          ws.unMergeCells(t.top, t.left, t.bottom, t.right);
        }
      } else {
        for (const t of targets) {
          ws.mergeCells(t.top, t.left, t.bottom, t.right);
        }
      }
    }

    return {
      sheet: ws.name,
      range: parsed.raw,
      affected,
      unmerge: Boolean(options.unmerge),
      dryRun: Boolean(options.dryRun),
    };
  }

  /**
   * Writes individual cells by A1 address, leaving every other cell untouched.
   * Formula cells refuse the write unless overwriteFormulas is passed; an
   * identical normalized formula re-applies idempotently, same as apply().
   */
  public setCells(
    entries: { a1: string; value: ReportCell }[],
    options?: XlsxSetCellsOptions
  ): XlsxSetCellsResult {
    if (!Array.isArray(entries) || entries.length === 0) {
      throw new Error('setCells requires a non-empty list of {a1, value} entries.');
    }

    const changes: XlsxDiffItem[] = [];
    const conflicts: XlsxDiffItem[] = [];

    for (const entry of entries) {
      const parsed = parseA1Range(entry.a1, { defaultSheet: options?.worksheetTitle });
      if (parsed.startRow !== parsed.endRow || parsed.startCol !== parsed.endCol || parsed.isFullCol || parsed.isFullRow) {
        throw new Error(`--cells entry "${entry.a1}" must be a single cell address (e.g. "B5" or "Sheet1!B5").`);
      }
      const ws = this._resolveWorksheet(parsed.sheetName ?? options?.worksheetTitle);
      const cell = ws.getCell(parsed.startRow, parsed.startCol);
      const extracted = this._extractCellData(cell);
      const address = cellAddress(parsed.startRow, parsed.startCol);

      let after: XlsxCellSnapshot;
      if (typeof entry.value === 'object' && entry.value !== null && 'formula' in entry.value) {
        const raw = (entry.value as ReportFormula).formula;
        after = { value: (entry.value as ReportFormula).result ?? null, formula: raw.startsWith('=') ? raw.slice(1) : raw };
      } else {
        after = { value: entry.value as ReportScalar };
      }

      let identical = false;
      if (extracted.isFormula && after.formula && extracted.formulaString) {
        identical =
          after.formula.trim().replace(/^=/, '').toUpperCase() ===
          extracted.formulaString.trim().replace(/^=/, '').toUpperCase();
      }

      const item: XlsxDiffItem = {
        sheet: ws.name,
        address,
        row: parsed.startRow,
        col: parsed.startCol,
        action: extracted.isFormula && !options?.overwriteFormulas && !identical ? 'formula-conflict' : extracted.scalarValue !== null || extracted.isFormula ? 'update' : 'insert',
        before: { value: extracted.scalarValue, formula: extracted.formulaString, formattedText: extracted.formattedText },
        after,
        ...(extracted.isFormula && !options?.overwriteFormulas && !identical
          ? { conflictReason: `Cell "${address}" already contains formula "${extracted.formulaString}" and overwriteFormulas is not enabled.` }
          : {}),
      };
      changes.push(item);
      if (item.action === 'formula-conflict') conflicts.push(item);
    }

    if (conflicts.length > 0 && options?.overwriteFormulas !== true) {
      const list = conflicts.map((c) => c.address).join(', ');
      throw new Error(
        `Cannot write cells: ${conflicts.length} formula overwrite conflict(s) in [${list}]. Pass --overwriteFormulas to allow overwriting existing formulas.`
      );
    }

    if (!options?.dryRun) {
      for (const change of changes) {
        const ws = this._resolveWorksheet(change.sheet);
        const cell = ws.getCell(change.row, change.col);
        if (change.after?.formula) {
          cell.value = { formula: change.after.formula, result: change.after.value ?? undefined };
        } else {
          cell.value = change.after?.value ?? null;
        }
      }
    }

    return {
      applied: options?.dryRun ? 0 : changes.length,
      changes,
      conflicts,
      dryRun: Boolean(options?.dryRun),
    };
  }

  /**
   * Clears cell values in an explicitly bounded range, keeping styles and every
   * other cell property. Formula cells refuse the clear unless overwriteFormulas
   * is passed - the same contract as the Sheets-side data:clear.
   */
  public clearRange(options: XlsxClearOptions): XlsxClearResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const parsed = parseA1Range(options.range, { defaultSheet: ws.name });
    if (parsed.isFullCol || parsed.isFullRow) {
      throw new Error(`Clear range "${options.range}" must be bounded on both axes (e.g. "Sheet1!A2:D20").`);
    }

    const formulasInRange: string[] = [];
    let cellsCleared = 0;
    for (let r = parsed.startRow; r <= parsed.endRow; r++) {
      for (let c = parsed.startCol; c <= parsed.endCol; c++) {
        const cell = ws.getCell(r, c);
        const extracted = this._extractCellData(cell);
        if (extracted.isFormula) formulasInRange.push(cellAddress(r, c));
        if (extracted.scalarValue !== null || extracted.isFormula) cellsCleared++;
      }
    }

    if (formulasInRange.length > 0 && options?.overwriteFormulas !== true) {
      throw new Error(
        `Cannot clear range "${parsed.raw}": ${formulasInRange.length} formula cell(s) present [${formulasInRange.join(', ')}]. Pass --overwriteFormulas to clear them.`
      );
    }

    if (!options?.dryRun) {
      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        for (let c = parsed.startCol; c <= parsed.endCol; c++) {
          ws.getCell(r, c).value = null;
        }
      }
    }

    return {
      sheet: ws.name,
      range: parsed.raw,
      cellsCleared,
      formulasOverwritten: formulasInRange,
      dryRun: Boolean(options?.dryRun),
    };
  }

  /**
   * Applies a formatting subset to every cell of a bounded A1 range without
   * touching values or formulas. Only the style groups ExcelJS models natively
   * are supported (font, fill, alignment, wrap, number format, borders); cloud
   * flags without an ExcelJS equivalent are refused at the command layer.
   * clear resets the whole style of every cell and cannot be combined with
   * style keys.
   */
  public formatCells(options: XlsxFormatCellsOptions): XlsxFormatCellsResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const parsed = parseA1Range(options.range, { defaultSheet: ws.name });
    if (parsed.isFullCol || parsed.isFullRow) {
      throw new Error(`Format range "${options.range}" must be bounded on both axes (e.g. "Sheet1!A1:J50").`);
    }

    const prepared = this._prepareCellStyle(options);

    if (!options.dryRun) {
      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        for (let c = parsed.startCol; c <= parsed.endCol; c++) {
          this._applyCellStyle(ws.getCell(r, c), prepared, r, c, parsed);
        }
      }
    }

    return {
      sheet: ws.name,
      range: parsed.raw,
      cellsFormatted: (parsed.endRow - parsed.startRow + 1) * (parsed.endCol - parsed.startCol + 1),
      applied: prepared.applied,
      clear: Boolean(options.clear),
      dryRun: Boolean(options.dryRun),
    };
  }

  /**
   * Freezes or unfreezes rows and columns on a worksheet by rewriting the
   * first sheet view (ExcelJS models the frozen pane as views[0]). At least
   * one of rows/columns must be given; an unspecified axis keeps whatever the
   * prior view froze on it, 0 on an axis unfreezes it, and 0 on both unfreezes
   * the pane entirely. The new view is merged over the old one so unrelated
   * view attributes (zoomScale, showGridLines, ...) survive.
   */
  public freezePanes(options: XlsxFreezeOptions): XlsxFreezeResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const rows = options.rows;
    const columns = options.columns;
    if (rows === undefined && columns === undefined) {
      throw new Error('freezePanes requires rows or columns to be given (0 unfreezes that axis).');
    }
    for (const [label, value] of [['rows', rows], ['columns', columns]] as const) {
      if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
        throw new Error(`Freeze ${label} must be a non-negative integer, got ${value}.`);
      }
    }

    // views is Partial<WorksheetView>[] and its split fields only exist on the
    // frozen/split variants; reading them off the prior view needs the
    // structural cast (serialized views render xSplit/ySplit as numbers).
    const priorView = (ws.views ?? [])[0] as { xSplit?: number; ySplit?: number } | undefined;
    const frozenRows = rows ?? priorView?.ySplit ?? 0;
    const frozenCols = columns ?? priorView?.xSplit ?? 0;
    const frozen = frozenRows > 0 || frozenCols > 0;
    const model = frozen
      ? `views[0] = {state:'frozen', xSplit:${frozenCols}, ySplit:${frozenRows}}`
      : `views[0] = {state:'normal'}`;

    if (!options.dryRun) {
      if (frozen) {
        ws.views = [
          {
            ...priorView,
            state: 'frozen',
            xSplit: frozenCols,
            ySplit: frozenRows,
            topLeftCell: `${indexToColLetter(frozenCols + 1)}${frozenRows + 1}`,
          },
        ];
      } else {
        // Unfreezing must clear the pane's leftover splits and topLeftCell, or
        // the normal view would serialize stale frozen-pane attributes. The
        // split fields only exist on the frozen/split view variants, so the
        // zeroed literal needs the structural cast.
        const normal = {
          ...priorView,
          state: 'normal',
          xSplit: 0,
          ySplit: 0,
          topLeftCell: undefined,
        } as Partial<ExcelJS.WorksheetView>;
        ws.views = [normal];
      }
    }

    return { sheet: ws.name, rows: frozenRows, columns: frozenCols, frozen, model, dryRun: Boolean(options.dryRun) };
  }

  /**
   * Resizes rows or columns. --pixels is converted from screen pixels into
   * Excel units (columns: width characters via the 7px/char + 5px padding
   * approximation of the default font; rows: points at 0.75pt/px). --auto
   * sizes columns to their longest cell text (capped) and resets rows to the
   * default height; ExcelJS has no measured auto-fit.
   */
  public resizeGrid(options: XlsxResizeOptions): XlsxResizeResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const dimension = options.dimension;
    if (dimension !== 'ROWS' && dimension !== 'COLUMNS') {
      throw new Error(`resizeGrid requires dimension ROWS or COLUMNS, got "${dimension}".`);
    }
    const start = options.start;
    const count = options.count ?? 1;
    if (!Number.isInteger(start) || start < 1) {
      throw new Error(`--start must be a positive 1-based index, got ${start}.`);
    }
    if (!Number.isInteger(count) || count < 1) {
      throw new Error(`--count must be a positive integer, got ${count}.`);
    }
    if (options.auto && options.pixels !== undefined) {
      throw new Error('resizeGrid accepts either pixels or auto, not both.');
    }
    if (!options.auto && options.pixels === undefined) {
      throw new Error('resizeGrid requires either a pixel size or auto sizing.');
    }
    if (options.pixels !== undefined && (!Number.isInteger(options.pixels) || options.pixels < 1)) {
      throw new Error(`--pixels must be a positive integer, got ${options.pixels}.`);
    }

    const isCols = dimension === 'COLUMNS';
    const spanEnd = start + count - 1;
    const limit = isCols ? MAX_EXCEL_COLS : MAX_EXCEL_ROWS;
    if (spanEnd > limit) {
      throw new Error(`Resize span ${start}..${spanEnd} exceeds the Excel limit of ${limit} ${isCols ? 'columns' : 'rows'}.`);
    }

    const resized: XlsxResizeTarget[] = [];
    for (let index = start; index <= spanEnd; index++) {
      if (isCols) {
        const col = ws.getColumn(index);
        const before = col.width;
        const after = options.auto ? this._autoColumnWidth(ws, index) : this._pixelsToWidth(options.pixels as number);
        if (!options.dryRun) col.width = after;
        resized.push({ index, before, after });
      } else {
        const row = ws.getRow(index);
        const before = row.height;
        const after = options.auto ? XLSX_DEFAULT_ROW_HEIGHT : this._pixelsToHeight(options.pixels as number);
        if (!options.dryRun) row.height = after;
        resized.push({ index, before, after });
      }
    }

    return {
      sheet: ws.name,
      dimension,
      start,
      count,
      unit: isCols ? 'width-chars' : 'height-points',
      resized,
      dryRun: Boolean(options.dryRun),
    };
  }

  /**
   * Hides or unhides rows or columns (col.hidden / row.hidden). Purely a
   * visibility flag: values, styles and formulas are untouched.
   */
  public setGridHidden(options: XlsxHideOptions): XlsxHideResult {
    const ws = this._resolveWorksheet(options.worksheetTitle);
    const dimension = options.dimension;
    if (dimension !== 'ROWS' && dimension !== 'COLUMNS') {
      throw new Error(`setGridHidden requires dimension ROWS or COLUMNS, got "${dimension}".`);
    }
    const start = options.start;
    const count = options.count ?? 1;
    if (!Number.isInteger(start) || start < 1) {
      throw new Error(`--start must be a positive 1-based index, got ${start}.`);
    }
    if (!Number.isInteger(count) || count < 1) {
      throw new Error(`--count must be a positive integer, got ${count}.`);
    }

    const isCols = dimension === 'COLUMNS';
    const spanEnd = start + count - 1;
    const limit = isCols ? MAX_EXCEL_COLS : MAX_EXCEL_ROWS;
    if (spanEnd > limit) {
      throw new Error(`Hide span ${start}..${spanEnd} exceeds the Excel limit of ${limit} ${isCols ? 'columns' : 'rows'}.`);
    }

    const hidden = options.unhide !== true;
    if (!options.dryRun) {
      for (let index = start; index <= spanEnd; index++) {
        if (isCols) ws.getColumn(index).hidden = hidden;
        else ws.getRow(index).hidden = hidden;
      }
    }

    const noun = isCols ? 'column' : 'row';
    const span = count > 1 ? `${start}-${spanEnd}` : `${start}`;
    return {
      sheet: ws.name,
      dimension,
      start,
      count,
      hidden,
      model: `${noun} ${span}: hidden = ${hidden}`,
      dryRun: Boolean(options.dryRun),
    };
  }

  /**
   * Adds an empty worksheet. Duplicate titles are refused (ExcelJS would throw
   * a bare internal error; this one names the existing sheet up front).
   */
  public addSheet(title: string, options?: XlsxSheetOpOptions): XlsxAddSheetResult {
    const trimmed = title?.trim();
    if (!trimmed) {
      throw new Error('worksheetTitle is required to add a sheet on the local backend.');
    }
    if (this.workbook.getWorksheet(trimmed)) {
      throw new Error(`A worksheet named "${trimmed}" already exists. Pick another title or remove the existing sheet first.`);
    }
    if (!options?.dryRun) {
      this.workbook.addWorksheet(trimmed);
    }
    return { operation: 'add', sheet: trimmed, dryRun: Boolean(options?.dryRun) };
  }

  /**
   * Removes a worksheet by title. The workbook must keep at least one visible
   * sheet afterwards - removing the last visible one is refused.
   */
  public removeSheet(title: string, options?: XlsxSheetOpOptions): XlsxRemoveSheetResult {
    const trimmed = title?.trim();
    if (!trimmed) {
      throw new Error('worksheetTitle is required to remove a sheet on the local backend.');
    }
    const ws = this.workbook.getWorksheet(trimmed);
    if (!ws) {
      const available = this.workbook.worksheets.map((w) => `"${w.name}"`).join(', ');
      throw new Error(`No worksheet named "${trimmed}" in this workbook. Available sheets: ${available || '(none)'}.`);
    }
    const remainingVisible = this.workbook.worksheets.filter(
      (w) => w !== ws && w.state !== 'hidden' && w.state !== 'veryHidden'
    );
    if (remainingVisible.length === 0) {
      throw new Error(`Cannot remove "${trimmed}": it is the last visible worksheet in the workbook.`);
    }
    if (!options?.dryRun) {
      // ExcelJS's removeWorksheet never touches definedNames.matrixMap, so
      // names scoped to the removed sheet would serialize dangling.
      this._dropDefinedNamesForSheet(trimmed);
      this.workbook.removeWorksheet(trimmed);
    }
    return { operation: 'remove', sheet: trimmed, dryRun: Boolean(options?.dryRun) };
  }

  /**
   * Deletes defined names whose ranges all reference the given sheet (matched
   * by "Title!" / "'Title'!" range prefixes). Names spanning several sheets are
   * left untouched - partial rewrites are phase 2, the same scope line as
   * renameSheet.
   */
  private _dropDefinedNamesForSheet(title: string): void {
    const matrixMap = (this.workbook.definedNames as unknown as {
      matrixMap?: Record<string, unknown>;
    }).matrixMap;
    if (!matrixMap) return;
    const prefixes = [`${title}!`, `'${title}'!`];
    for (const { name, ranges } of this._readDefinedNames()) {
      if (ranges.length > 0 && ranges.every((range) => prefixes.some((prefix) => range.startsWith(prefix)))) {
        delete matrixMap[name];
      }
    }
  }

  /**
   * Renames a worksheet. The new title must not collide with an existing sheet.
   * Defined names are not rewritten (phase 2 of --updateRefs; refs pointing at
   * the old title go stale exactly like ExcelJS's own rename).
   */
  public renameSheet(from: string, to: string, options?: XlsxSheetOpOptions): XlsxRenameSheetResult {
    const fromTrimmed = from?.trim();
    const toTrimmed = to?.trim();
    if (!fromTrimmed || !toTrimmed) {
      throw new Error('renameSheet requires the current worksheetTitle and the new title.');
    }
    const ws = this.workbook.getWorksheet(fromTrimmed);
    if (!ws) {
      const available = this.workbook.worksheets.map((w) => `"${w.name}"`).join(', ');
      throw new Error(`No worksheet named "${fromTrimmed}" in this workbook. Available sheets: ${available || '(none)'}.`);
    }
    if (toTrimmed !== fromTrimmed && this.workbook.getWorksheet(toTrimmed)) {
      throw new Error(`A worksheet named "${toTrimmed}" already exists. Pick another title.`);
    }
    if (!options?.dryRun) {
      ws.name = toTrimmed;
    }
    return { operation: 'rename', from: fromTrimmed, to: toTrimmed, dryRun: Boolean(options?.dryRun) };
  }

  /**
   * Lists the workbook's defined names (named ranges). Read-only counterpart
   * of inspect().definedNames.
   */
  public listDefinedNames(): XlsxDefinedName[] {
    return this._readDefinedNames();
  }

  /**
   * Adds a defined name. refersTo must be a sheet-qualified A1 range - ExcelJS's
   * own DefinedNames.add silently accepts garbage location strings, so the
   * format is validated here first. Adding over an existing name is refused;
   * remove it first.
   */
  public addDefinedName(options: XlsxAddNameOptions): XlsxAddNameResult {
    const name = options.name?.trim();
    if (!name) {
      throw new Error('A defined name is required (a non-empty --name).');
    }
    const refersTo = options.refersTo?.trim();
    if (!refersTo || !XLSX_REFERS_TO_RE.test(refersTo)) {
      throw new Error(
        `refersTo "${options.refersTo}" is not a sheet-qualified A1 range. Use the form Sheet!$A$1:$A$9 (quoted titles allowed: 'My Sheet'!$A$1).`
      );
    }
    if (this._readDefinedNames().some((n) => n.name === name)) {
      throw new Error(`A defined name "${name}" already exists. Remove it first or pick another name.`);
    }
    if (!options.dryRun) {
      // ExcelJS argument order is (location, name)
      this.workbook.definedNames.add(refersTo, name);
    }
    return { name, refersTo, dryRun: Boolean(options.dryRun) };
  }

  /**
   * Removes a defined name together with all of its ranges. ExcelJS's public
   * DefinedNames.remove(locStr, name) only deletes a single decoded cell and
   * leaves range-backed names intact, so the name's matrix is dropped from the
   * internal matrixMap instead - the same store the serializer reads.
   */
  public removeDefinedName(options: XlsxRemoveNameOptions): XlsxRemoveNameResult {
    const name = options.name?.trim();
    if (!name) {
      throw new Error('A defined name is required (a non-empty --name).');
    }
    const existing = this._readDefinedNames().find((n) => n.name === name);
    if (!existing) {
      const available = this._readDefinedNames().map((n) => `"${n.name}"`).join(', ');
      throw new Error(`No defined name "${name}" in this workbook. Available names: ${available || '(none)'}.`);
    }
    if (!options.dryRun) {
      const matrixMap = (this.workbook.definedNames as unknown as { matrixMap: Record<string, unknown> }).matrixMap;
      delete matrixMap[name];
    }
    return { name, removedRanges: existing.ranges, dryRun: Boolean(options.dryRun) };
  }

  /**
   * Resolves the style spec once per formatCells call so per-cell application
   * stays a plain merge. Validates every flag here - including the ones the
   * command layer already typed - so engine callers get the same refusals.
   */
  private _prepareCellStyle(options: XlsxFormatCellsOptions): PreparedCellStyle {
    const clear = options.clear === true;
    if (clear) {
      const clash = XLSX_FORMAT_KEYS.filter((key) => options[key] !== undefined);
      if (clash.length > 0) {
        throw new Error(`clear cannot be combined with style flags: ${clash.join(', ')}.`);
      }
      return { clear: true, applied: ['clear'] };
    }

    const applied: string[] = [];
    const prepared: PreparedCellStyle = { clear: false, applied };

    const font: Partial<ExcelJS.Font> = {};
    if (options.bold !== undefined) {
      font.bold = options.bold;
      applied.push('bold');
    }
    if (options.italic !== undefined) {
      font.italic = options.italic;
      applied.push('italic');
    }
    if (options.underline !== undefined) {
      font.underline = options.underline;
      applied.push('underline');
    }
    if (options.strikethrough !== undefined) {
      font.strike = options.strikethrough;
      applied.push('strikethrough');
    }
    if (options.fontSize !== undefined) {
      if (!Number.isFinite(options.fontSize) || options.fontSize < 1 || options.fontSize > 409) {
        throw new Error(`fontSize must be between 1 and 409 points, got ${options.fontSize}.`);
      }
      font.size = options.fontSize;
      applied.push('fontSize');
    }
    if (options.fontFamily !== undefined) {
      if (!options.fontFamily.trim()) {
        throw new Error('fontFamily must be a non-empty font name.');
      }
      font.name = options.fontFamily;
      applied.push('fontFamily');
    }
    if (options.textColor !== undefined) {
      font.color = { argb: this._argbFromHex(options.textColor, 'textColor') };
      applied.push('textColor');
    }
    if (Object.keys(font).length > 0) prepared.font = font;

    if (options.backgroundColor !== undefined) {
      prepared.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: this._argbFromHex(options.backgroundColor, 'backgroundColor') },
      };
      applied.push('backgroundColor');
    }

    const alignment: Partial<ExcelJS.Alignment> = {};
    if (options.horizontalAlignment !== undefined) {
      const horizontal = ({ LEFT: 'left', CENTER: 'center', RIGHT: 'right' } as const)[options.horizontalAlignment];
      if (!horizontal) {
        throw new Error(`horizontalAlignment must be LEFT, CENTER or RIGHT, got "${options.horizontalAlignment}".`);
      }
      alignment.horizontal = horizontal;
      applied.push('horizontalAlignment');
    }
    if (options.verticalAlignment !== undefined) {
      const vertical = ({ TOP: 'top', MIDDLE: 'middle', BOTTOM: 'bottom' } as const)[options.verticalAlignment];
      if (!vertical) {
        throw new Error(`verticalAlignment must be TOP, MIDDLE or BOTTOM, got "${options.verticalAlignment}".`);
      }
      alignment.vertical = vertical;
      applied.push('verticalAlignment');
    }
    if (options.wrapText !== undefined) {
      alignment.wrapText = options.wrapText;
      applied.push('wrapText');
    }
    if (Object.keys(alignment).length > 0) prepared.alignment = alignment;

    if (options.numberFormat !== undefined) {
      if (!options.numberFormat.trim()) {
        throw new Error('numberFormat must be a non-empty pattern.');
      }
      prepared.numFmt = options.numberFormat;
      applied.push('numberFormat');
    }

    if (options.borders !== undefined) {
      prepared.borderSides = this._parseBorderSides(options.borders);
      if (options.borderStyle !== undefined && !(options.borderStyle in XLSX_BORDER_STYLE_MAP)) {
        throw new Error(
          `borderStyle "${options.borderStyle}" is not supported. Supported styles: ${Object.keys(XLSX_BORDER_STYLE_MAP).join(', ')}.`
        );
      }
      prepared.borderStyle = options.borderStyle === undefined ? 'thin' : XLSX_BORDER_STYLE_MAP[options.borderStyle];
      prepared.borderColor = {
        argb: this._argbFromHex(options.borderColor ?? '#000000', 'borderColor'),
      };
      applied.push(`borders(${[...prepared.borderSides].join(',')})`);
    }

    if (applied.length === 0) {
      throw new Error('No formatting flags were given. Pass style flags (e.g. bold, backgroundColor) or clear.');
    }
    return prepared;
  }

  /** Applies one prepared style spec to a single cell, merging with its existing style. */
  private _applyCellStyle(
    cell: ExcelJS.Cell,
    prepared: PreparedCellStyle,
    row: number,
    col: number,
    parsed: ParsedA1Range
  ): void {
    if (prepared.clear) {
      cell.style = {};
      return;
    }
    if (prepared.font) cell.font = { ...cell.font, ...prepared.font };
    if (prepared.fill) cell.fill = prepared.fill;
    if (prepared.alignment) cell.alignment = { ...cell.alignment, ...prepared.alignment };
    if (prepared.numFmt !== undefined) cell.numFmt = prepared.numFmt;
    if (prepared.borderSides) {
      const style = prepared.borderStyle;
      const next: Partial<ExcelJS.Borders> = { ...cell.border };
      for (const side of this._concreteBorderSides(prepared.borderSides, row, col, parsed)) {
        if (style === undefined) next[side] = undefined;
        else next[side] = { style, color: prepared.borderColor };
      }
      cell.border = next;
    }
  }

  /**
   * Expands the requested border sides for one cell position: "inner" sides
   * exist only between two cells of the range, so a cell on the range edge
   * gets no border there.
   */
  private _concreteBorderSides(
    sides: Set<XlsxBorderSide>,
    row: number,
    col: number,
    parsed: ParsedA1Range
  ): ('top' | 'bottom' | 'left' | 'right')[] {
    const out = new Set<'top' | 'bottom' | 'left' | 'right'>();
    for (const side of sides) {
      if (side === 'top' || side === 'bottom' || side === 'left' || side === 'right') {
        out.add(side);
      } else if (side === 'innerHorizontal') {
        if (row > parsed.startRow) out.add('top');
        if (row < parsed.endRow) out.add('bottom');
      } else {
        if (col > parsed.startCol) out.add('left');
        if (col < parsed.endCol) out.add('right');
      }
    }
    return [...out];
  }

  private _parseBorderSides(raw: string): Set<XlsxBorderSide> {
    const valid: XlsxBorderSide[] = ['top', 'bottom', 'left', 'right', 'innerHorizontal', 'innerVertical'];
    const sides = raw
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    if (sides.length === 1 && sides[0] === 'all') {
      return new Set<XlsxBorderSide>(['top', 'bottom', 'left', 'right']);
    }
    if (sides.length === 1 && sides[0] === 'inner') {
      return new Set<XlsxBorderSide>(['innerHorizontal', 'innerVertical']);
    }
    const out = new Set<XlsxBorderSide>();
    for (const side of sides) {
      if (!(valid as string[]).includes(side)) {
        throw new Error(`Invalid border side "${side}". Use a comma list of ${valid.join(', ')} or "all"/"inner".`);
      }
      out.add(side as XlsxBorderSide);
    }
    if (out.size === 0) {
      throw new Error(`borders needs at least one side. Use a comma list of ${valid.join(', ')} or "all"/"inner".`);
    }
    return out;
  }

  private _argbFromHex(raw: string, label: string): string {
    const match = /^#?([0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/.exec(raw.trim());
    if (!match) {
      throw new Error(`${label} color "${raw}" must be #RRGGBB (or #AARRGGBB).`);
    }
    const hex = match[1].toUpperCase();
    return hex.length === 6 ? `FF${hex}` : hex;
  }

  /**
   * Auto-fit approximation for one column: the longest cell text (ExcelJS's
   * derived text, cached formula results included) plus padding, capped. An
   * empty column keeps a default width.
   */
  private _autoColumnWidth(ws: ExcelJS.Worksheet, col: number): number {
    let maxLen = 0;
    // Iterate allocated rows: actualRowCount counts rows WITH data, not the
    // last row index, so a sparse sheet (data in rows 1 and 1000) measured only
    // the leading run under a 1..actualRowCount loop.
    ws.eachRow((row) => {
      const text = row.getCell(col).text;
      if (text) maxLen = Math.max(maxLen, String(text).length);
    });
    if (maxLen === 0) return XLSX_DEFAULT_COL_WIDTH;
    return Math.min(maxLen + XLSX_AUTO_WIDTH_PADDING, XLSX_MAX_AUTO_WIDTH);
  }

  /** Converts a pixel size into Excel width characters (7px/char + 5px padding, default font). */
  private _pixelsToWidth(pixels: number): number {
    const width = Math.round(((pixels - 5) / 7) * 100) / 100;
    if (width <= 0) {
      throw new Error(`A pixel size of ${pixels} is too small for a column width; use at least 6.`);
    }
    return Math.min(width, XLSX_MAX_COL_WIDTH);
  }

  /** Converts a pixel size into row height points (0.75pt per px at 96dpi). */
  private _pixelsToHeight(pixels: number): number {
    return Math.min(Math.round(pixels * 0.75 * 100) / 100, XLSX_MAX_ROW_HEIGHT);
  }

  /**
   * Reads the workbook's defined names. Type boundary: ExcelJS's DefinedNames
   * typing lacks `model`; at runtime the getter returns a fresh array (or a
   * record keyed by name) on every access - read-only here, mutation of the
   * returned value would be a no-op.
   */
  private _readDefinedNames(): XlsxDefinedName[] {
    const definedNames = this.workbook.definedNames as unknown as {
      model?: Record<string, { name: string; ranges: string[] }> | { name: string; ranges: string[] }[];
    };
    const model = definedNames.model;
    const out: XlsxDefinedName[] = [];
    if (Array.isArray(model)) {
      for (const entry of model) {
        if (entry && typeof entry === 'object' && entry.name && Array.isArray(entry.ranges)) {
          out.push({ name: entry.name, ranges: entry.ranges });
        }
      }
    } else if (model && typeof model === 'object') {
      for (const [key, val] of Object.entries(model)) {
        if (val && typeof val === 'object' && Array.isArray(val.ranges)) {
          out.push({ name: val.name || key, ranges: val.ranges });
        }
      }
    }
    return out;
  }

  /**
   * Resolves a worksheet by title, falling back to the first sheet. Throws with
   * the available sheet names when the title does not exist.
   */
  private _resolveWorksheet(title?: string): ExcelJS.Worksheet {
    const ws = title ? this.workbook.getWorksheet(title) : this.workbook.worksheets[0];
    if (!ws) {
      const available = this.workbook.worksheets.map((w) => `"${w.name}"`).join(', ');
      throw new Error(
        `Worksheet "${title ?? 'Default'}" not found in workbook. Available sheets: [${available}].`
      );
    }
    return ws;
  }

  /** Merge bounds as {top,left,bottom,right} parsed from ws.model.merges. */
  private _mergeBounds(ws: ExcelJS.Worksheet): { top: number; left: number; bottom: number; right: number }[] {
    const out: { top: number; left: number; bottom: number; right: number }[] = [];
    for (const rangeStr of (ws.model.merges as string[] | undefined) ?? []) {
      try {
        const p = parseA1Range(rangeStr);
        out.push({ top: p.startRow, left: p.startCol, bottom: p.endRow, right: p.endCol });
      } catch {
        // unparseable merge entry: leave it alone
      }
    }
    return out;
  }

  private _boundsToA1(b: { top: number; left: number; bottom: number; right: number }): string {
    return formatA1Range({ startCol: b.left, startRow: b.top, endCol: b.right, endRow: b.bottom });
  }

  /**
   * Counts formula cells whose formula text references a row/column at or after
   * the splice point on the same sheet. Heuristic only - it drives the warning
   * count, never a rewrite.
   */
  private _countFormulasReferencing(ws: ExcelJS.Worksheet, start: number, isRows: boolean): number {
    const refRe = isRows
      ? new RegExp(`(?<![A-Z0-9_!])\\$?[A-Z]{1,3}\\$?(${start}|[1-9][0-9]*)(?![0-9(])`, 'g')
      : /(?<![A-Z0-9_!])\$?([A-Z]{1,3})\$?[0-9]+(?![0-9(])/g;
    const startColLetter = isRows ? undefined : indexToColLetter(start);
    let count = 0;
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        const extracted = this._extractCellData(cell);
        if (!extracted.isFormula || !extracted.formulaString) return;
        const f = extracted.formulaString;
        if (isRows) {
          refRe.lastIndex = 0;
          let m: RegExpExecArray | null;
          while ((m = refRe.exec(f)) !== null) {
            if (parseInt(m[1], 10) >= start) {
              count++;
              return;
            }
          }
        } else {
          refRe.lastIndex = 0;
          let m: RegExpExecArray | null;
          while ((m = refRe.exec(f)) !== null) {
            try {
              if (colLetterToIndex(m[1]) >= colLetterToIndex(startColLetter!)) {
                count++;
                return;
              }
            } catch {
              // not a column ref
            }
          }
        }
      });
    });
    return count;
  }

  /**
   * Snapshots every formula cell's effective text BEFORE a physical splice.
   * Master cells contribute their own text; shared-formula slave cells
   * contribute the text ExcelJS translates for them through their offset from
   * the master - read now, while the master still holds pre-splice
   * coordinates. Rewriting a slave from its snapshot promotes it to a
   * standalone formula (semantics preserved, and the master is rewritten too).
   * Entries carry the pre-splice address because spliceRows/spliceColumns copy
   * values into fresh cell objects: cell references do not survive the splice.
   */
  private _snapshotFormulaTexts(ws: ExcelJS.Worksheet): {
    row: number;
    col: number;
    current: string;
    result: ExcelJS.CellFormulaValue['result'];
  }[] {
    const snapshot: {
      row: number;
      col: number;
      current: string;
      result: ExcelJS.CellFormulaValue['result'];
    }[] = [];
    ws.eachRow((row) => {
      row.eachCell((cell, colNumber) => {
        const value = cell.value;
        if (!value || typeof value !== 'object') return;
        if (!('formula' in value) && !('sharedFormula' in value)) return;
        // the `in` guard narrows value to ExcelJS's CellFormulaValue/CellSharedFormulaValue
        const current = typeof value.formula === 'string' ? value.formula : cell.formula;
        if (typeof current !== 'string' || current === '') return;
        snapshot.push({ row: row.number, col: colNumber, current, result: value.result });
      });
    });
    return snapshot;
  }

  /**
   * Rewrites the pre-splice formula snapshot against a splice operation and
   * writes changed formulas back to their post-splice addresses: inserts shift
   * cells at/after the splice point by count, deletes shift cells beyond the
   * span up/left and drop the span itself.
   */
  private _rewriteFormulaRefs(
    ws: ExcelJS.Worksheet,
    snapshot: { row: number; col: number; current: string; result: ExcelJS.CellFormulaValue['result'] }[],
    op: XlsxRefSpliceOp,
    isRows: boolean
  ): { rewritten: number; refErrors: number } {
    let rewritten = 0;
    let refErrors = 0;
    const spanEnd = op.start + op.count - 1;
    for (const entry of snapshot) {
      const v = isRows ? entry.row : entry.col;
      let row = entry.row;
      let col = entry.col;
      if (op.mode === 'insert') {
        if (v >= op.start) {
          if (isRows) row += op.count;
          else col += op.count;
        }
      } else if (v > spanEnd) {
        if (isRows) row -= op.count;
        else col -= op.count;
      } else if (v >= op.start) {
        continue; // the cell itself fell inside the deleted span
      }
      const result = rewriteFormulaRefs(entry.current, op);
      if (!result.changed) continue;
      ws.getCell(row, col).value = { formula: result.formula, result: entry.result };
      rewritten++;
      refErrors += result.refErrors;
    }
    return { rewritten, refErrors };
  }

  /**
   * Snapshots defined-name ranges as name -> range strings. ExcelJS keeps names
   * in a private cell matrix; the model getter hands out an array of
   * {name, ranges} with fully anchored "Sheet!$A$1:$B$2" strings. Names on other
   * sheets are included and simply never change.
   */
  private _definedNameSignature(): Map<string, string[]> {
    const signature = new Map<string, string[]>();
    const definedNamesModel = this.workbook.definedNames as unknown as { model?: { name: string; ranges: string[] }[] };
    const model = definedNamesModel.model;
    if (Array.isArray(model)) {
      for (const entry of model) {
        if (entry && typeof entry.name === 'string' && Array.isArray(entry.ranges)) {
          signature.set(entry.name, [...entry.ranges]);
        }
      }
    }
    return signature;
  }

  /** Clones per-cell styles from one row onto count rows starting at targetStart. */
  private _copyRowStyle(ws: ExcelJS.Worksheet, sourceRow: number, targetStart: number, count: number): void {
    const src = ws.findRow(sourceRow);
    if (!src) return;
    for (let i = 0; i < count; i++) {
      const dst = ws.getRow(targetStart + i);
      dst.height = src.height;
      src.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        dst.getCell(colNumber).style = { ...cell.style };
      });
    }
  }

  /** Clones per-cell styles from one column onto count columns starting at targetStart. */
  private _copyColStyle(ws: ExcelJS.Worksheet, sourceCol: number, targetStart: number, count: number): void {
    const maxRow = Math.max(ws.actualRowCount || 0, 1);
    for (let i = 0; i < count; i++) {
      const dstCol = targetStart + i;
      ws.getColumn(dstCol).width = ws.getColumn(sourceCol).width;
      for (let r = 1; r <= maxRow; r++) {
        const srcCell = ws.findCell(r, sourceCol);
        if (srcCell) ws.getCell(r, dstCol).style = { ...srcCell.style };
      }
    }
  }

  /**
   * Generates a preview diff of applying a ReportDocument to the workbook without mutating it.
   */
  public preview(
    document: ReportDocument,
    options?: XlsxApplyOptions
  ): XlsxDiffResult {
    this._validateReportDocument(document);

    const changes: XlsxDiffItem[] = [];
    const conflicts: XlsxDiffItem[] = [];
    const affectedSheets = new Set<string>();
    const managedExtents: XlsxManagedExtent[] = [];

    for (const sheet of document.sheets) {
      affectedSheets.add(sheet.name);

      const targetBounds = calculateTargetRange(sheet, options?.defaultStartCell);
      const ws = this.workbook.getWorksheet(sheet.name);

      // Check for managed range clearing
      const shouldClear = sheet.clearManagedRange ?? options?.clearManagedRange ?? false;
      const previousManaged = this._managedExtents.get(sheet.name);
      let clearedRangeStr: string | undefined;

      if (shouldClear && previousManaged && ws) {
        try {
          const parsedPrev = parseA1Range(previousManaged);
          // If previous extent was larger than new bounds, identify cells to clear
          for (let r = parsedPrev.startRow; r <= parsedPrev.endRow; r++) {
            for (let c = parsedPrev.startCol; c <= parsedPrev.endCol; c++) {
              const insideNew =
                r >= targetBounds.startRow &&
                r <= targetBounds.endRow &&
                c >= targetBounds.startCol &&
                c <= targetBounds.endCol;

              if (!insideNew) {
                const existingCell = ws.getCell(r, c);
                const extracted = this._extractCellData(existingCell);
                if (extracted.scalarValue !== null || extracted.isFormula) {
                  changes.push({
                    sheet: sheet.name,
                    address: cellAddress(r, c),
                    row: r,
                    col: c,
                    action: 'clear',
                    before: {
                      value: extracted.scalarValue,
                      formula: extracted.formulaString,
                      formattedText: extracted.formattedText,
                    },
                    after: { value: null },
                  });
                }
              }
            }
          }
          clearedRangeStr = previousManaged;
        } catch {
          // If previous range string was invalid, ignore
        }
      }

      // Check cells to be written
      for (let r = 0; r < sheet.rows.length; r++) {
        const row = sheet.rows[r];
        const targetRow = targetBounds.startRow + r;

        for (let c = 0; c < row.length; c++) {
          const targetCol = targetBounds.startCol + c;
          const incomingItem = row[c];
          const address = cellAddress(targetRow, targetCol);

          let beforeSnapshot: XlsxCellSnapshot | undefined;
          let isExistingFormula = false;

          if (ws) {
            const existingCell = ws.getCell(targetRow, targetCol);
            const extracted = this._extractCellData(existingCell);
            isExistingFormula = extracted.isFormula;
            beforeSnapshot = {
              value: extracted.scalarValue,
              formula: extracted.formulaString,
              formattedText: extracted.formattedText,
            };
          }

          // Determine after snapshot
          let afterSnapshot: XlsxCellSnapshot;
          if (
            typeof incomingItem === 'object' &&
            incomingItem !== null &&
            'formula' in incomingItem
          ) {
            const rawFormula = (incomingItem as ReportFormula).formula;
            const cleanFormula = rawFormula.startsWith('=') ? rawFormula.slice(1) : rawFormula;
            afterSnapshot = {
              value: (incomingItem as ReportFormula).result ?? null,
              formula: cleanFormula,
            };
          } else {
            afterSnapshot = {
              value: incomingItem as ReportScalar,
            };
          }

          // Check for formula overwrite safety (identical normalized formulas are allowed idempotently)
          const overwriteOpt = options?.overwriteFormulas === true;
          let isIdenticalFormula = false;
          if (isExistingFormula && afterSnapshot.formula && beforeSnapshot?.formula) {
            const normIncoming = afterSnapshot.formula.trim().replace(/^=/, '').toUpperCase();
            const normExisting = beforeSnapshot.formula.trim().replace(/^=/, '').toUpperCase();
            if (normIncoming === normExisting) {
              isIdenticalFormula = true;
            }
          }

          if (isExistingFormula && !overwriteOpt && !isIdenticalFormula) {
            const conflictItem: XlsxDiffItem = {
              sheet: sheet.name,
              address,
              row: targetRow,
              col: targetCol,
              action: 'formula-conflict',
              before: beforeSnapshot,
              after: afterSnapshot,
              conflictReason: `Target cell "${address}" already contains formula "${beforeSnapshot?.formula}" and overwriteFormulas is not enabled.`,
            };
            conflicts.push(conflictItem);
            changes.push(conflictItem);
          } else {
            const action: XlsxDiffAction = beforeSnapshot?.value !== null || beforeSnapshot?.formula ? 'update' : 'insert';
            changes.push({
              sheet: sheet.name,
              address,
              row: targetRow,
              col: targetCol,
              action,
              before: beforeSnapshot,
              after: afterSnapshot,
            });
          }
        }
      }

      managedExtents.push({
        sheet: sheet.name,
        range: targetBounds.rangeString,
        clearedRange: clearedRangeStr,
      });
    }

    return {
      documentProvenance: document.provenance,
      changes,
      conflicts,
      totalChanges: changes.length,
      totalConflicts: conflicts.length,
      affectedSheets: Array.from(affectedSheets),
      managedExtents,
    };
  }

  /**
   * Applies a ReportDocument to the workbook in-memory with safety checks,
   * style preservation, presentation settings, and managed extent management.
   */
  public apply(
    document: ReportDocument,
    options?: XlsxApplyOptions
  ): XlsxApplyResult {
    const diff = this.preview(document, options);

    if (diff.totalConflicts > 0 && options?.overwriteFormulas !== true) {
      const conflictList = diff.conflicts.map((c) => c.address).join(', ');
      throw new Error(
        `Cannot apply report: ${diff.totalConflicts} formula overwrite conflict(s) in [${conflictList}]. Pass overwriteFormulas: true to allow overwriting existing formulas.`
      );
    }

    if (options?.dryRun) {
      return {
        success: true,
        appliedChanges: 0,
        diff,
        warnings: ['Dry run enabled. Workbook was not modified.'],
      };
    }

    const warnings: string[] = [];

    for (const sheet of document.sheets) {
      let ws = this.workbook.getWorksheet(sheet.name);
      if (!ws) {
        ws = this.workbook.addWorksheet(sheet.name);
      }

      const targetBounds = calculateTargetRange(sheet, options?.defaultStartCell);

      // 1. Clear previous managed range if requested
      const shouldClear = sheet.clearManagedRange ?? options?.clearManagedRange ?? false;
      const previousManaged = this._managedExtents.get(sheet.name);

      if (shouldClear && previousManaged) {
        try {
          const parsedPrev = parseA1Range(previousManaged);
          for (let r = parsedPrev.startRow; r <= parsedPrev.endRow; r++) {
            for (let c = parsedPrev.startCol; c <= parsedPrev.endCol; c++) {
              const insideNew =
                r >= targetBounds.startRow &&
                r <= targetBounds.endRow &&
                c >= targetBounds.startCol &&
                c <= targetBounds.endCol;

              if (!insideNew) {
                const cell = ws.getCell(r, c);
                cell.value = null;
              }
            }
          }
        } catch {
          // Ignore range parse error on previous extent
        }
      }

      // 2. Build number format map for quick lookup by 1-based column
      const numFmtMap = new Map<number, string>();
      if (sheet.numberFormats && Array.isArray(sheet.numberFormats)) {
        for (const nf of sheet.numberFormats) {
          if (nf && typeof nf.column === 'number' && typeof nf.format === 'string') {
            numFmtMap.set(nf.column, nf.format);
          }
        }
      }

      // 3. Write rows
      for (let r = 0; r < sheet.rows.length; r++) {
        const row = sheet.rows[r];
        const targetRow = targetBounds.startRow + r;

        for (let c = 0; c < row.length; c++) {
          const targetCol = targetBounds.startCol + c;
          const cell = ws.getCell(targetRow, targetCol);
          const incoming = row[c];

          // Check if cell has existing format
          const hasExistingNumFmt = Boolean(cell.numFmt && cell.numFmt !== 'General');

          // Write cell value
          if (
            typeof incoming === 'object' &&
            incoming !== null &&
            'formula' in incoming
          ) {
            const rawFormula = (incoming as ReportFormula).formula;
            const cleanFormula = rawFormula.startsWith('=') ? rawFormula.slice(1) : rawFormula;
            cell.value = {
              formula: cleanFormula,
              result: (incoming as ReportFormula).result ?? undefined,
            };
          } else if (typeof incoming === 'string') {
            // Ordinary strings stay literal string, even if starting with '='
            cell.value = incoming;
          } else {
            cell.value = incoming as ReportScalar;
          }

          // Apply number formatting: template style wins
          const preserveStyles = options?.preserveTemplateStyles !== false;
          if (!preserveStyles || !hasExistingNumFmt) {
            // Relative column index in report sheet (1-based)
            const colIndexInReport = c + 1;
            const formatForCol = numFmtMap.get(colIndexInReport) || numFmtMap.get(targetCol);
            if (formatForCol) {
              cell.numFmt = formatForCol;
            }
          }
        }
      }

      // 4. Apply presentation: freeze rows
      if (sheet.freezeRows && sheet.freezeRows > 0) {
        ws.views = [
          {
            state: 'frozen',
            xSplit: 0,
            ySplit: sheet.freezeRows,
            activeCell: cellAddress(sheet.freezeRows + 1, 1),
          },
        ];
      }

      // 5. Apply presentation: column widths
      if (sheet.columnWidths && Array.isArray(sheet.columnWidths)) {
        for (let i = 0; i < sheet.columnWidths.length; i++) {
          const w = sheet.columnWidths[i];
          if (typeof w === 'number' && w > 0) {
            const colIdx = targetBounds.startCol + i;
            ws.getColumn(colIdx).width = w;
          }
        }
      }

      // 6. Record managed extent
      this._managedExtents.set(sheet.name, targetBounds.rangeString);
      if (options?.recordProvenance !== false) {
        this._recordManagedRangeDefinedName(sheet.name, targetBounds.rangeString);
      }
    }

    return {
      success: true,
      appliedChanges: diff.totalChanges,
      diff,
      warnings,
    };
  }

  /**
   * Safely saves workbook to disk atomically with conflict checking and backup.
   */
  public async save(
    targetPath?: string,
    options?: XlsxSaveOptions
  ): Promise<XlsxSaveResult> {
    const destination = targetPath || (options?.inPlace ? this.filePath : undefined);
    if (!destination) {
      throw new Error(
        'No target path specified for save. Pass a destination path or use inPlace: true.'
      );
    }

    // Check unsupported features preflight guard
    const hasUnsupported = this.preflight?.hasUnsupportedFeatures ?? false;
    const allowUnsupported = options?.allowUnsupportedFeatures === true;
    const skipPreflight = options?.skipPreflight === true;

    if (hasUnsupported && !allowUnsupported && !skipPreflight) {
      const descriptions = (this.preflight?.unsupportedFeatures ?? [])
        .map((f) => f.description)
        .join(', ');
      throw new Error(
        `Cannot safely save workbook containing unsupported features: [${descriptions}]. ExcelJS would silently drop these features. Pass allowUnsupportedFeatures: true to force save.`
      );
    }

    const arrayBuffer = await this.workbook.xlsx.writeBuffer();
    const buffer = Buffer.from(arrayBuffer as ArrayBuffer);
    return saveBufferAtomic(buffer, destination, options, {
      originalPath: this.filePath,
      originalHash: this.originalHash,
    });
  }

  /**
   * Helper to extract normalized cell data and formula freshness.
   */
  private _extractCellData(cell: ExcelJS.Cell): CellExtraction {
    const rawValue = cell.value;
    const numFmt = cell.numFmt;
    const hasStyle = Boolean(
      (cell.font && Object.keys(cell.font).length > 0) ||
      (cell.fill && Object.keys(cell.fill).length > 0) ||
      (cell.border && Object.keys(cell.border).length > 0) ||
      (numFmt && numFmt !== 'General')
    );

    if (rawValue === null || rawValue === undefined) {
      return {
        scalarValue: null,
        formattedText: '',
        isFormula: false,
        freshness: 'empty',
        hasStyle,
        numberFormat: numFmt,
      };
    }

    // Case 1: ExcelJS Formula Value
    if (
      cell.type === ExcelJS.ValueType.Formula ||
      (typeof rawValue === 'object' && rawValue !== null && ('formula' in rawValue || 'sharedFormula' in rawValue))
    ) {
      const formulaObj = rawValue as { formula?: string; result?: unknown; sharedFormula?: string };
      // ExcelJS translates shared-formula slave cells into relative formulas via cell.formula getter
      let formulaStr = cell.formula || formulaObj.formula || '';
      if (!formulaStr && formulaObj.sharedFormula && cell.worksheet) {
        try {
          const masterCell = cell.worksheet.getCell(formulaObj.sharedFormula);
          if (masterCell && masterCell.formula) {
            formulaStr = cell.formula || '';
          }
        } catch {
          // Master cell unresolvable
        }
      }

      const cached = (cell.result !== undefined ? cell.result : formulaObj.result) as ReportScalar | undefined;
      const scalarVal = this._toReportScalar(cached !== undefined ? cached : null);
      const isFresh = cached !== undefined && cached !== null;

      return {
        scalarValue: scalarVal,
        formattedText: cell.text || (scalarVal !== null ? String(scalarVal) : ''),
        formulaString: formulaStr || undefined,
        cachedResult: scalarVal,
        isFormula: true,
        freshness: isFresh ? 'cached' : 'unknown',
        hasStyle,
        numberFormat: numFmt,
      };
    }

    // Case 2: Rich Text
    if (typeof rawValue === 'object' && rawValue !== null && 'richText' in rawValue) {
      const richObj = rawValue as { richText: Array<{ text: string }> };
      const combined = richObj.richText.map((t) => t.text).join('');
      return {
        scalarValue: combined,
        formattedText: combined,
        isFormula: false,
        freshness: 'not-formula',
        hasStyle,
        numberFormat: numFmt,
      };
    }

    // Case 3: Hyperlink
    if (typeof rawValue === 'object' && rawValue !== null && 'hyperlink' in rawValue) {
      const linkObj = rawValue as { text?: string; hyperlink: string };
      const text = linkObj.text || linkObj.hyperlink;
      return {
        scalarValue: text,
        formattedText: text,
        isFormula: false,
        freshness: 'not-formula',
        hasStyle,
        numberFormat: numFmt,
      };
    }

    // Case 4: Cell Error
    if (typeof rawValue === 'object' && rawValue !== null && 'error' in rawValue) {
      const errObj = rawValue as { error: string };
      return {
        scalarValue: errObj.error,
        formattedText: errObj.error,
        isFormula: false,
        freshness: 'not-formula',
        hasStyle,
        numberFormat: numFmt,
      };
    }

    // Case 5: Date object
    if (rawValue instanceof Date) {
      const iso = rawValue.toISOString();
      return {
        scalarValue: iso,
        formattedText: cell.text || iso,
        isFormula: false,
        freshness: 'not-formula',
        hasStyle,
        numberFormat: numFmt,
      };
    }

    // Case 6: Primitives (string, number, boolean)
    const scalar = this._toReportScalar(rawValue);
    return {
      scalarValue: scalar,
      formattedText: cell.text || (scalar !== null ? String(scalar) : ''),
      isFormula: false,
      freshness: 'not-formula',
      hasStyle,
      numberFormat: numFmt,
    };
  }

  private _toReportScalar(val: unknown): ReportScalar {
    if (val === null || val === undefined) return null;
    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
      return val;
    }
    if (val instanceof Date) {
      return val.toISOString();
    }
    return String(val);
  }

  private _validateReportDocument(doc: unknown): asserts doc is ReportDocument {
    if (!doc || typeof doc !== 'object') {
      throw new Error('Invalid ReportDocument: must be a non-null object.');
    }
    const d = doc as Partial<ReportDocument>;
    if (!Array.isArray(d.sheets)) {
      throw new Error('Invalid ReportDocument: "sheets" must be an array.');
    }
    if (!d.provenance || typeof d.provenance !== 'object') {
      throw new Error('Invalid ReportDocument: "provenance" must be an object.');
    }
    if (!d.provenance.templateId || !d.provenance.sourceHash) {
      throw new Error('Invalid ReportDocument: provenance must contain templateId and sourceHash.');
    }
    for (let i = 0; i < d.sheets.length; i++) {
      const s = d.sheets[i];
      if (!s || typeof s !== 'object' || typeof s.name !== 'string' || !Array.isArray(s.rows)) {
        throw new Error(`Invalid ReportSheet at index ${i}: must have "name" (string) and "rows" (array).`);
      }
    }
  }

  private _loadManagedExtentsFromDefinedNames(): void {
    const model = (this.workbook.definedNames as unknown as { model?: Record<string, { name: string; ranges: string[] }> })?.model;
    if (model && typeof model === 'object') {
      for (const [key, val] of Object.entries(model)) {
        const name = val?.name || key;
        if (name.startsWith(MANAGED_RANGE_PREFIX) && Array.isArray(val.ranges) && val.ranges.length > 0) {
          const sheetName = name.slice(MANAGED_RANGE_PREFIX.length);
          this._managedExtents.set(sheetName, val.ranges[0]);
        }
      }
    }
  }

  private _recordManagedRangeDefinedName(sheetName: string, rangeString: string): void {
    try {
      const definedNameKey = `${MANAGED_RANGE_PREFIX}${sheetName}`;
      // ExcelJS definedNames management
      const ranges = [rangeString];
      const definedNames = this.workbook.definedNames as unknown as {
        model?: Record<string, { name: string; ranges: string[] }>;
      };
      if (!definedNames.model) {
        definedNames.model = {};
      }
      definedNames.model[definedNameKey] = {
        name: definedNameKey,
        ranges,
      };
    } catch {
      // Non-critical metadata recording
    }
  }
}
