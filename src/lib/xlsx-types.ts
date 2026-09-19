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
  managedRange?: string;
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

export interface XlsxCreateOptions {
  creator?: string;
  created?: Date;
  modified?: Date;
}
