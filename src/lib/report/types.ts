/**
 * Shared Report Document and Template Types.
 * Fixed contract for Google Sheets and XLSX report automation.
 */

// ============================================================================
// Core Document Types (Fixed Contract)
// ============================================================================

export type ReportScalar = string | number | boolean | null;

export interface ReportFormula {
  /** Formula expression without leading '=' (e.g. "SUM(A1:A10)", "D2*E2") */
  formula: string;
  /** Optional pre-computed deterministic result */
  result?: ReportScalar;
}

export type ReportCell = ReportScalar | ReportFormula;

export interface ReportNumberFormat {
  /** 1-based column index */
  column: number;
  /** Format string (e.g. "#,##0.00", "YYYY-MM-DD", "$#,##0.00", "0.0%") */
  format: string;
}

export interface ReportSheet {
  /** Worksheet / tab name */
  name: string;
  /** 2D grid of cells (including headers) */
  rows: ReportCell[][];
  /** Number of frozen rows at the top (e.g. 1 for header) */
  freezeRows?: number;
  /** Optional column widths in characters */
  columnWidths?: number[];
  /** Optional number formats per 1-based column index */
  numberFormats?: ReportNumberFormat[];
  /** Optional top-left placement cell (e.g. "A1", "B3") */
  startCell?: string;
  /** Optional target named range for template mapping */
  targetNamedRange?: string;
  /** Optional flag to clear managed range before writing */
  clearManagedRange?: boolean;
  /** Optional sheet-level metadata */
  metadata?: Record<string, unknown>;
}

export interface ReportProvenance {
  /** Template identifier */
  templateId: string;
  /** Template version */
  templateVersion: string | number;
  /** Deterministic SHA-256 hash of normalized input data */
  sourceHash: string;
  /** ISO 8601 generation timestamp */
  generatedAt?: string;
  /** High-level summary metrics */
  summary?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface ReportDocument {
  /** Ordered list of worksheets */
  sheets: ReportSheet[];
  /** Provenance and audit metadata */
  provenance: ReportProvenance;
}

// ============================================================================
// Input Types
// ============================================================================

export interface ReadInputOptions {
  /** Inline JSON string or data payload */
  data?: string;
  /** File path, or '-' for stdin */
  file?: string;
  /** Explicit format: 'json' or 'csv' */
  format?: 'json' | 'csv';
  /** Max allowed bytes to read (default 20MB) */
  maxBytes?: number;
}

export interface ParseInputOptions {
  /** Whether CSV should be parsed as 2D matrix array instead of objects */
  matrix?: boolean;
  /** CSV delimiter (default ',') */
  delimiter?: string;
  /** Strict column count matching header */
  strictColumnCount?: boolean;
  /** Preserve leading zeros in numeric strings */
  preserveLeadingZeros?: boolean;
}

// ============================================================================
// Schema Normalization Types
// ============================================================================

export type FieldType = 'string' | 'decimal' | 'integer' | 'boolean' | 'date';

export interface FieldSchema {
  /** Field / column name */
  name: string;
  /** Expected data type */
  type: FieldType;
  /** Whether the field is mandatory (cannot be undefined or missing) */
  required?: boolean;
  /** Whether null is an acceptable value */
  nullable?: boolean;
  /** Convert empty strings "" to null before type validation */
  emptyAsNull?: boolean;
  /** Explicit date format for parsing (e.g. "YYYY-MM-DD", "DD/MM/YYYY") */
  dateFormat?: string;
  /** Explicit decimal separator (e.g. "." or ",") */
  decimalSeparator?: string;
  /** Explicit thousand grouping separator (e.g. "," or ".") */
  thousandSeparator?: string;
  /** Decimal places for rounding (banker's / half-up) */
  decimalPlaces?: number;
  /** Allowed enum values */
  enum?: (string | number | boolean)[];
  /** Minimum value for numbers / min length for strings */
  min?: number | string;
  /** Maximum value for numbers / max length for strings */
  max?: number | string;
  /** Human-readable field description */
  description?: string;
}

export interface TableSchema {
  /** List of field schemas */
  fields: FieldSchema[];
  /** Allow columns not explicitly listed in schema */
  allowExtraColumns?: boolean;
  /** Require strictly preserved leading zeros in string types */
  strictLeadingZeros?: boolean;
}

export interface ValidationError {
  /** 1-based row index */
  row: number;
  /** 1-based column index if available */
  column?: number;
  /** Field / property name */
  field: string;
  /** The invalid value */
  value: unknown;
  /** Descriptive error message */
  message: string;
  /** Error code identifier */
  code: string;
}

// ============================================================================
// Declarative Transforms Types
// ============================================================================

export type FilterOperator =
  | 'eq'
  | 'neq'
  | 'gt'
  | 'gte'
  | 'lt'
  | 'lte'
  | 'in'
  | 'nin'
  | 'between'
  | 'contains'
  | 'startsWith'
  | 'endsWith'
  | 'isNull'
  | 'isNotNull'
  | 'and'
  | 'or'
  | 'not';

export interface FilterCondition {
  field?: string;
  op: FilterOperator;
  value?: unknown;
  values?: unknown[];
  conditions?: FilterCondition[];
}

export type MathOperator =
  | 'add'
  | 'sub'
  | 'mul'
  | 'div'
  | 'mod'
  | 'abs'
  | 'neg'
  | 'round'
  | 'ceil'
  | 'floor'
  | 'min'
  | 'max';

export type StringOperator =
  | 'concat'
  | 'trim'
  | 'lower'
  | 'upper'
  | 'substring'
  | 'replace';

export type LogicOperator = 'if' | 'coalesce' | 'switch';

export type TransformOperator = MathOperator | StringOperator | LogicOperator | 'lookup' | 'fxConvert' | 'custom';

export interface ComputedExpression {
  op: TransformOperator | string;
  args?: (string | number | boolean | ComputedExpression | null)[];
  condition?: FilterCondition;
  then?: unknown;
  else?: unknown;
  field?: string;
  format?: string;
  places?: number;
}

export interface FilterTransformStep {
  kind: 'filter';
  condition: FilterCondition;
}

export interface ComputedTransformStep {
  kind: 'computed' | 'map';
  targetField: string;
  expression: ComputedExpression;
}

export interface LookupJoinStep {
  kind: 'lookupJoin';
  sourceKey: string;
  lookupTable: string;
  lookupKey: string;
  select: Record<string, string>;
  joinType?: 'inner' | 'left';
  defaultValues?: Record<string, unknown>;
}

export interface AggregationDef {
  field: string;
  op: 'sum' | 'count' | 'countDistinct' | 'avg' | 'min' | 'max';
  as: string;
  round?: number;
}

export interface GroupSumStep {
  kind: 'groupSum' | 'aggregate';
  groupBy: string[];
  aggregations: AggregationDef[];
}

export interface PivotStep {
  kind: 'pivot';
  rowGroupBy: string[];
  columnPivot: string;
  valueField: string;
  aggOp?: 'sum' | 'count';
  columnValues?: string[];
}

export interface SortFieldDef {
  field: string;
  order?: 'asc' | 'desc';
}

export interface SortStep {
  kind: 'sort';
  fields: SortFieldDef[];
}

export interface LimitStep {
  kind: 'limit';
  limit: number;
  offset?: number;
}

export type TransformStep =
  | FilterTransformStep
  | ComputedTransformStep
  | LookupJoinStep
  | GroupSumStep
  | PivotStep
  | SortStep
  | LimitStep;

// ============================================================================
// Output Sheet Mapping Types
// ============================================================================

export interface OutputColumn {
  /** Header label in row 1 */
  header: string;
  /** Source field name from data */
  field?: string;
  /** Explicit Excel formula template (e.g. "D{row}*E{row}") without '=' */
  formulaTemplate?: string;
  /** Optional pre-computed result field name */
  resultField?: string;
  /** Static column width */
  width?: number;
  /** Number format */
  format?: string;
}

export interface OutputSummaryCell {
  /** Field or value */
  field?: string;
  value?: ReportScalar;
  /** Formula expression without '=' */
  formula?: string;
  /** Result value */
  result?: ReportScalar;
  format?: string;
}

export interface OutputSummaryRow {
  /** Label for first column */
  label?: string;
  cells: OutputSummaryCell[];
}

export interface ReportSheetTemplate {
  name: string;
  sourceDataset?: string;
  columns: OutputColumn[];
  freezeRows?: number;
  summaryRows?: OutputSummaryRow[];
  startCell?: string;
  targetNamedRange?: string;
  clearManagedRange?: boolean;
}

// ============================================================================
// Template Schema (Version 1)
// ============================================================================

export type TemplateKind = 'table' | 'finance' | 'manpower';

export interface ReportPolicy {
  /** Strictly enforce schema types */
  strictTypes?: boolean;
  /** Never treat missing values or errors as 0 */
  treatMissingAsZero?: boolean;
  /** Policy when duplicate keys are encountered */
  duplicateKeyPolicy?: 'reject' | 'keepFirst' | 'keepLast';
  /** Max input rows allowed */
  maxInputRows?: number;
}

export interface ReportTemplateV1 {
  /** Unique template identifier */
  id: string;
  /** Version number, strictly 1 */
  version: 1 | number | string;
  /** Report kind */
  kind: TemplateKind;
  /** Human-readable title */
  title?: string;
  /** Template description */
  description?: string;
  /** Schema definitions for inputs */
  schema?: TableSchema | Record<string, TableSchema>;
  /** Pipeline of declarative transformation steps */
  transforms?: TransformStep[];
  /** Output worksheet definitions */
  outputs?: ReportSheetTemplate[];
  /** Preset-specific configuration */
  config?: Record<string, unknown>;
  /** Execution policy */
  policy?: ReportPolicy;
}

// ============================================================================
// Finance Preset Types
// ============================================================================

export interface FinanceTransaction {
  id: string;
  date: string;
  account: string;
  direction: 'in' | 'out' | 'IN' | 'OUT';
  amount: string | number;
  currency: string;
  description?: string;
  category?: string;
  isRefund?: boolean;
  [key: string]: unknown;
}

export interface FxRate {
  date: string;
  from: string;
  to: string;
  rate: string | number;
  source?: string;
}

export interface FinanceOpeningBalance {
  account: string;
  balance: string | number;
  currency?: string;
}

export interface FinanceConfig {
  /** Target base currency for reporting (e.g. "USD", "VND", "EUR") */
  baseCurrency: string;
  /** Available FX exchange rates */
  fxRates?: FxRate[];
  /** Opening balances per account */
  openingBalances?: Record<string, string | number> | FinanceOpeningBalance[];
  /** Optional expected closing balances for reconciliation audit */
  expectedClosingBalances?: Record<string, string | number>;
  /** Optional date filter window */
  dateWindow?: {
    from?: string;
    to?: string;
  };
  /** Discrepancy tolerance for balance reconciliation */
  tolerance?: number | string;
}

export interface FinanceReportSummary {
  baseCurrency: string;
  totalInflow: string;
  totalOutflow: string;
  netCashFlow: string;
  transactionCount: number;
  accountCount: number;
  reconciled: boolean;
  totalDiscrepancy: string;
}

// ============================================================================
// Manpower Preset Types
// ============================================================================

export interface ManpowerTask {
  id: string;
  module: string;
  role: string;
  quantity: number | string;
  unitEffort: number | string;
  complexity?: number | string;
  contingency?: number | string;
  description?: string;
  [key: string]: unknown;
}

export interface RoleRate {
  role: string;
  dailyRate: number | string;
  unit?: string;
}

export interface ScenarioConfig {
  name: string;
  complexityMultiplier?: number | string;
  contingencyPercent?: number | string;
  roleRateMultipliers?: Record<string, number | string>;
  /** Optional flag marking this scenario as the baseline for delta comparisons */
  isBase?: boolean;
}

export interface ManpowerConfig {
  /** Currency code for rates (e.g. "VND", "USD") */
  currency?: string;
  /** Unit of effort (default "person-days") */
  effortUnit?: 'hours' | 'days' | 'person-days';
  /** Mapping of role names to standard daily rates */
  roleRates: Record<string, number | string> | RoleRate[];
  /** Global contingency percentage (e.g. 20 for 20%) */
  globalContingencyPercent?: number | string;
  /** Optional explicit baseline scenario name for delta comparisons */
  baselineScenario?: string;
  /** Multi-scenario simulations */
  scenarios?: ScenarioConfig[];
  /** Working days per person per month */
  baseWorkingDaysPerMonth?: number;
}

export interface ManpowerReportSummary {
  effortUnit: string;
  currency: string;
  totalBaseEffort: string;
  totalEffort: string;
  totalCost: string;
  moduleCount: number;
  roleCount: number;
  taskCount: number;
}
