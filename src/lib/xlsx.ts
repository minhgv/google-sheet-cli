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
  XlsxDefinedName,
  XlsxDiffAction,
  XlsxDiffItem,
  XlsxDiffResult,
  XlsxFreshnessSummary,
  XlsxInspection,
  XlsxLoadOptions,
  XlsxManagedExtent,
  XlsxPreflightResult,
  XlsxReadMode,
  XlsxReadOptions,
  XlsxReadResult,
  XlsxSaveOptions,
  XlsxSaveResult,
  XlsxSheetMetadata,
} from './xlsx-types';
import {
  calculateTargetRange,
  cellAddress,
  formatA1Range,
  parseA1Range,
  parseCellAddress,
} from './xlsx-range';
import { inspectZipBuffer } from './xlsx-zip';
import { computeFileSha256, saveBufferAtomic } from './xlsx-file';

export * from './xlsx-types';
export * from './xlsx-range';
export * from './xlsx-zip';
export * from './xlsx-file';

const MANAGED_RANGE_PREFIX = '_GS_MANAGED_';

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
  public inspect(): XlsxInspection {
    const sheets: XlsxSheetMetadata[] = this.workbook.worksheets.map((ws) => {
      let hasFormulas = false;
      ws.eachRow((row) => {
        row.eachCell((cell) => {
          if (
            cell.type === ExcelJS.ValueType.Formula ||
            (cell.value && typeof cell.value === 'object' && 'formula' in cell.value)
          ) {
            hasFormulas = true;
          }
        });
      });

      const managedRange = this._managedExtents.get(ws.name);

      return {
        id: Number(ws.id),
        name: ws.name,
        state: (ws.state as 'visible' | 'hidden' | 'veryHidden') || 'visible',
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
        hasFormulas,
        managedRange,
      };
    });

    const definedNames: XlsxDefinedName[] = [];
    const model = (this.workbook.definedNames as unknown as { model?: Record<string, { name: string; ranges: string[] }> })?.model;
    if (model && typeof model === 'object') {
      for (const [key, val] of Object.entries(model)) {
        if (val && typeof val === 'object' && Array.isArray(val.ranges)) {
          definedNames.push({ name: val.name || key, ranges: val.ranges });
        }
      }
    }

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
