/**
 * Shared Report and XLSX Adapter Types
 *
 * Fixed ReportDocument contract matching ReportCore:
 * - ReportScalar: string | number | boolean | null
 * - ReportFormula: { formula: string; result?: ReportScalar }
 * - ReportCell: ReportScalar | ReportFormula
 * - ReportSheet: { name, rows, freezeRows?, columnWidths?, numberFormats?, startCell?, targetNamedRange?, clearManagedRange?, metadata? }
 * - ReportDocument: { sheets, provenance: { templateId, templateVersion, sourceHash, generatedAt?, summary? } }
 */

export type ReportScalar = string | number | boolean | null;

export interface ReportFormula {
  formula: string;
  result?: ReportScalar;
}

export type ReportCell = ReportScalar | ReportFormula;

export interface ReportNumberFormat {
  column: number; // 1-based index
  format: string;
}

export interface ReportSheet {
  name: string;
  rows: ReportCell[][];
  freezeRows?: number;
  columnWidths?: number[];
  numberFormats?: ReportNumberFormat[];
  startCell?: string;
  targetNamedRange?: string;
  clearManagedRange?: boolean;
  metadata?: Record<string, unknown>;
}

export interface ReportProvenance {
  templateId: string;
  templateVersion: string | number;
  sourceHash: string;
  generatedAt?: string;
  summary?: Record<string, unknown>;
}

export interface ReportDocument {
  sheets: ReportSheet[];
  provenance: ReportProvenance;
}

// ---------------------------------------------------------------------------
// XLSX Read Types
// ---------------------------------------------------------------------------

export type XlsxReadMode = 'formatted' | 'unformatted' | 'formula';

export type XlsxCellFreshness = 'cached' | 'empty' | 'unknown' | 'not-formula';

export interface XlsxCellReadInfo {
  address: string; // e.g. "A1"
  row: number; // 1-based index
  col: number; // 1-based index
  value: ReportScalar;
  formattedText?: string;
  formula?: string;
  cachedResult?: ReportScalar;
  freshness: XlsxCellFreshness;
  numberFormat?: string;
  hasStyle: boolean;
}

export interface XlsxReadOptions {
  mode?: XlsxReadMode;
  maxRows?: number;
  maxCols?: number;
  includeEmpty?: boolean;
  defaultSheet?: string;
}

export interface XlsxFreshnessSummary {
  cachedCount: number;
  unknownCount: number;
  emptyCount: number;
  notFormulaCount: number;
}

export interface XlsxReadResult {
  range: string;
  sheetName: string;
  mode: XlsxReadMode;
  values: ReportCell[][];
  cellDetails: XlsxCellReadInfo[][];
  formulaCount: number;
  freshnessSummary: XlsxFreshnessSummary;
  rowCount: number;
  colCount: number;
}

// ---------------------------------------------------------------------------
// XLSX Preflight & Inspection Types
// ---------------------------------------------------------------------------

export type XlsxUnsupportedFeatureType =
  | 'macro'
  | 'chart'
  | 'pivotTable'
  | 'externalLink'
  | 'embeddedObject'
  | 'dataConnection'
  | 'customXml'
  | 'slicer'
  | 'modelExtension'
  | 'other';

export interface XlsxUnsupportedFeature {
  type: XlsxUnsupportedFeatureType;
  description: string;
  path: string;
}

export interface XlsxPreflightResult {
  hasUnsupportedFeatures: boolean;
  unsupportedFeatures: XlsxUnsupportedFeature[];
  sheetNames: string[];
  entryCount: number;
  totalUncompressedSize: number;
  warnings: string[];
}

export interface XlsxSheetMetadata {
  id: number;
  name: string;
  state?: 'visible' | 'hidden' | 'veryHidden';
  rowCount: number;
  columnCount: number;
  hasFormulas: boolean;
  /** number of cells holding a formula (present whenever hasFormulas is true) */
  formulaCellCount?: number;
  /** formula cell addresses, only when inspect() is called with includeFormulaCells */
  formulaCells?: string[];
  /** merged ranges in A1 notation, e.g. ["B5:B10"] */
  mergedRanges?: string[];
  /** frozen pane split, when the sheet has one */
  frozen?: { rows: number; columns: number };
  /** data-validation ranges and their rule types */
  dataValidations?: { range: string; type?: string }[];
  /** conditional-formatting ranges with the number of rules on each */
  conditionalFormatting?: { range: string; ruleCount: number }[];
  managedRange?: string;
}

export interface XlsxInspectOptions {
  /** include the full formula-cell address list per sheet (bounded output otherwise) */
  includeFormulaCells?: boolean;
}

export interface XlsxDefinedName {
  name: string;
  ranges: string[];
}

export interface XlsxCapabilities {
  canRead: boolean;
  canWriteSafely: boolean;
  canRecalculateFormulas: boolean; // always false in ExcelJS
  preservesVba: boolean; // false in ExcelJS
  preservesCharts: boolean; // false in ExcelJS
  preservesPivotTables: boolean; // false in ExcelJS
}

export interface XlsxInspection {
  filePath?: string;
  fileSize?: number;
  sha256Hash?: string;
  sheetCount: number;
  sheets: XlsxSheetMetadata[];
  definedNames: XlsxDefinedName[];
  preflight: XlsxPreflightResult;
  hasUnsupportedFeatures: boolean;
  unsupportedFeatures: XlsxUnsupportedFeature[];
  capabilities: XlsxCapabilities;
  warnings: string[];
}

// ---------------------------------------------------------------------------
// XLSX Diff, Apply & Preview Types
// ---------------------------------------------------------------------------

export type XlsxDiffAction = 'update' | 'insert' | 'clear' | 'formula-conflict' | 'unchanged';

export interface XlsxCellSnapshot {
  value: ReportScalar;
  formula?: string;
  formattedText?: string;
}

export interface XlsxDiffItem {
  sheet: string;
  address: string;
  row: number;
  col: number;
  action: XlsxDiffAction;
  before?: XlsxCellSnapshot;
  after?: XlsxCellSnapshot;
  conflictReason?: string;
}

export interface XlsxManagedExtent {
  sheet: string;
  range: string;
  clearedRange?: string;
}

export interface XlsxDiffResult {
  documentProvenance: ReportProvenance;
  changes: XlsxDiffItem[];
  conflicts: XlsxDiffItem[];
  totalChanges: number;
  totalConflicts: number;
  affectedSheets: string[];
  managedExtents: XlsxManagedExtent[];
}

export interface XlsxApplyOptions {
  dryRun?: boolean;
  overwriteFormulas?: boolean;
  clearManagedRange?: boolean;
  defaultStartCell?: string;
  preserveTemplateStyles?: boolean;
  recordProvenance?: boolean;
}

export interface XlsxApplyResult {
  success: boolean;
  appliedChanges: number;
  diff: XlsxDiffResult;
  warnings: string[];
}

// ---------------------------------------------------------------------------
// XLSX Save & Load Types
// ---------------------------------------------------------------------------

export interface XlsxSaveOptions {
  overwrite?: boolean;
  inPlace?: boolean;
  backup?: boolean;
  backupPath?: string;
  expectedHash?: string;
  allowUnsupportedFeatures?: boolean;
  skipPreflight?: boolean;
}

export interface XlsxSaveResult {
  savedPath: string;
  bytesWritten: number;
  sha256Hash: string;
  backupCreated?: string;
  isOverwritten: boolean;
}

export interface XlsxLoadOptions {
  maxFileSize?: number; // default 50 MB
  maxUncompressedSize?: number; // default 200 MB
  maxEntries?: number; // default 10,000
  allowUnsupportedFeatures?: boolean;
  computeHash?: boolean; // default true
}

// ---------------------------------------------------------------------------
// XLSX Find Types (local counterpart of GoogleSheetCli.FindOptions/FindResult)
// ---------------------------------------------------------------------------

export interface XlsxFindOptions {
  worksheetTitle?: string;
  /** A1 range bounding the scan; defaults to the used range of the sheet */
  range?: string;
  equals?: string;
  contains?: string;
  regex?: string;
  /** restrict the scan to one column letter (e.g. "B") */
  column?: string;
  /** restrict the scan to the column whose header cell matches this value */
  header?: string;
  ignoreCase?: boolean;
  /** max matches returned; matchCount still reports the true total */
  limit?: number;
  /** collapse matches to unique rows and include full row values */
  byRow?: boolean;
}

export interface XlsxFindMatch {
  a1: string;
  row: number;
  column: number;
  columnLetter: string;
  value: string;
  /** present only when byRow is set: the full scanned row the match sits in */
  rowValues?: (string | number | boolean | null)[];
}

export interface XlsxFindResult {
  /** resolved A1 range that was scanned */
  range: string;
  matchCount: number;
  truncated: boolean;
  matches: XlsxFindMatch[];
}

// ---------------------------------------------------------------------------
// XLSX Structural Mutation Types (splice, merge, sparse write, clear)
// ---------------------------------------------------------------------------

export type XlsxDimension = 'ROWS' | 'COLUMNS';

export interface XlsxSpliceOptions {
  worksheetTitle?: string;
  dimension: XlsxDimension;
  /** 1-based first row/column index to affect */
  start: number;
  /** how many rows/columns the operation covers */
  count?: number;
  /** insert: clone the style of the row/column before the insertion point */
  inheritFromBefore?: boolean;
  /** allow merges intersecting the splice boundary to be adjusted instead of rejected */
  force?: boolean;
  dryRun?: boolean;
}

export interface XlsxSpliceResult {
  operation: 'insert' | 'delete';
  sheet: string;
  dimension: XlsxDimension;
  start: number;
  count: number;
  /** merges that were shifted, extended, shrunk or dropped by the splice */
  mergesAdjusted: { before: string; after?: string }[];
  /** merges intersecting the boundary that blocked the splice (empty when applied) */
  mergeConflicts: string[];
  /** formula cells whose references may now be stale (never rewritten - see warnings) */
  formulasAtRisk: number;
  /** values removed by a delete (empty for insert) */
  removedValues?: ReportCell[][];
  dryRun: boolean;
  warnings: string[];
}

export interface XlsxMergeOptions {
  worksheetTitle?: string;
  range: string;
  mergeType?: 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS';
  unmerge?: boolean;
  dryRun?: boolean;
}

export interface XlsxMergeResult {
  sheet: string;
  range: string;
  /** concrete ranges merged/unmerged (MERGE_COLUMNS/ROWS decompose into one per line) */
  affected: string[];
  unmerge: boolean;
  dryRun: boolean;
}

export interface XlsxSetCellEntry {
  /** A1 address, optionally sheet-qualified ("Sheet1!B2") */
  a1: string;
  value: ReportCell;
}

export interface XlsxSetCellsOptions {
  worksheetTitle?: string;
  overwriteFormulas?: boolean;
  dryRun?: boolean;
}

export interface XlsxSetCellsResult {
  applied: number;
  changes: XlsxDiffItem[];
  conflicts: XlsxDiffItem[];
  dryRun: boolean;
}

export interface XlsxClearOptions {
  worksheetTitle?: string;
  range: string;
  overwriteFormulas?: boolean;
  dryRun?: boolean;
}

export interface XlsxClearResult {
  sheet: string;
  range: string;
  cellsCleared: number;
  /** formula cells that would be / were cleared */
  formulasOverwritten: string[];
  dryRun: boolean;
}

export interface XlsxCreateOptions {
  creator?: string;
  created?: Date;
  modified?: Date;
}
