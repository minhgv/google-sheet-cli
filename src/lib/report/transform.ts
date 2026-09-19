/**
 * Secure declarative transformation engine for reports.
 * Pure AST-based expression and filter evaluation with zero eval / code execution.
 */

import Decimal from 'decimal.js';
import {
  AggregationDef,
  ComputedExpression,
  FilterCondition,
  GroupSumStep,
  LookupJoinStep,
  PivotStep,
  SortStep,
  TransformStep,
} from './types';

/**
 * Executes a pipeline of declarative transformation steps on an array of records.
 */
export function executeTransforms(
  data: Record<string, unknown>[],
  transforms?: TransformStep[],
  context: Record<string, unknown> = {}
): Record<string, unknown>[] {
  if (!transforms || transforms.length === 0) {
    return data;
  }

  let current = data;

  for (let s = 0; s < transforms.length; s++) {
    const step = transforms[s];
    switch (step.kind) {
      case 'filter': {
        current = current.filter((row) => evaluateCondition(row, step.condition, context));
        break;
      }

      case 'computed':
      case 'map': {
        current = current.map((row) => {
          const updated: Record<string, unknown> = Object.create(null);
          Object.assign(updated, row);
          const computedVal = evaluateExpression(row, step.expression, context);
          updated[step.targetField] = computedVal;
          return updated;
        });
        break;
      }

      case 'lookupJoin': {
        current = applyLookupJoin(current, step, context);
        break;
      }

      case 'groupSum':
      case 'aggregate': {
        current = applyGroupSum(current, step);
        break;
      }

      case 'pivot': {
        current = applyPivot(current, step);
        break;
      }

      case 'sort': {
        current = applySort(current, step);
        break;
      }

      case 'limit': {
        const offset = step.offset || 0;
        const limit = step.limit;
        current = current.slice(offset, offset + limit);
        break;
      }

      default: {
        const unknownStep = step as { kind?: string };
        throw new Error(`Unsupported transform step kind: "${unknownStep.kind || 'unknown'}"`);
      }
    }
  }

  return current;
}

/**
 * Evaluates a boolean FilterCondition against a record.
 */
export function evaluateCondition(
  row: Record<string, unknown>,
  cond: FilterCondition,
  context: Record<string, unknown> = {}
): boolean {
  if (!cond) return true;

  switch (cond.op) {
    case 'and': {
      if (!cond.conditions || cond.conditions.length === 0) return true;
      return cond.conditions.every((c) => evaluateCondition(row, c, context));
    }

    case 'or': {
      if (!cond.conditions || cond.conditions.length === 0) return true;
      return cond.conditions.some((c) => evaluateCondition(row, c, context));
    }

    case 'not': {
      if (!cond.conditions || cond.conditions.length === 0) return true;
      return !evaluateCondition(row, cond.conditions[0], context);
    }

    case 'isNull': {
      const val = cond.field ? row[cond.field] : undefined;
      return val === null || val === undefined;
    }

    case 'isNotNull': {
      const val = cond.field ? row[cond.field] : undefined;
      return val !== null && val !== undefined;
    }

    case 'eq': {
      const val = cond.field ? row[cond.field] : undefined;
      const target = resolveOperandValue(row, cond.value, context);
      return compareEquality(val, target);
    }

    case 'neq': {
      const val = cond.field ? row[cond.field] : undefined;
      const target = resolveOperandValue(row, cond.value, context);
      return !compareEquality(val, target);
    }

    case 'gt': {
      const val = cond.field ? row[cond.field] : undefined;
      const target = resolveOperandValue(row, cond.value, context);
      return compareOrder(val, target) > 0;
    }

    case 'gte': {
      const val = cond.field ? row[cond.field] : undefined;
      const target = resolveOperandValue(row, cond.value, context);
      return compareOrder(val, target) >= 0;
    }

    case 'lt': {
      const val = cond.field ? row[cond.field] : undefined;
      const target = resolveOperandValue(row, cond.value, context);
      return compareOrder(val, target) < 0;
    }

    case 'lte': {
      const val = cond.field ? row[cond.field] : undefined;
      const target = resolveOperandValue(row, cond.value, context);
      return compareOrder(val, target) <= 0;
    }

    case 'in': {
      const val = cond.field ? row[cond.field] : undefined;
      const allowed = Array.isArray(cond.values) ? cond.values : [];
      return allowed.some((target) => compareEquality(val, target));
    }

    case 'nin': {
      const val = cond.field ? row[cond.field] : undefined;
      const allowed = Array.isArray(cond.values) ? cond.values : [];
      return !allowed.some((target) => compareEquality(val, target));
    }

    case 'between': {
      const val = cond.field ? row[cond.field] : undefined;
      const range = Array.isArray(cond.values) ? cond.values : [];
      if (range.length < 2) return true;
      const min = resolveOperandValue(row, range[0], context);
      const max = resolveOperandValue(row, range[1], context);
      return compareOrder(val, min) >= 0 && compareOrder(val, max) <= 0;
    }

    case 'contains': {
      const val = cond.field ? row[cond.field] : undefined;
      if (val === null || val === undefined) return false;
      const str = String(val);
      const target = String(resolveOperandValue(row, cond.value, context) || '');
      return str.includes(target);
    }

    case 'startsWith': {
      const val = cond.field ? row[cond.field] : undefined;
      if (val === null || val === undefined) return false;
      const str = String(val);
      const target = String(resolveOperandValue(row, cond.value, context) || '');
      return str.startsWith(target);
    }

    case 'endsWith': {
      const val = cond.field ? row[cond.field] : undefined;
      if (val === null || val === undefined) return false;
      const str = String(val);
      const target = String(resolveOperandValue(row, cond.value, context) || '');
      return str.endsWith(target);
    }

    default:
      throw new Error(`Unknown filter operator "${(cond as FilterCondition).op}"`);
  }
}

/**
 * Resolves an expression or raw literal/field reference to its concrete evaluated value.
 */
export function evaluateExpression(
  row: Record<string, unknown>,
  expr: unknown,
  context: Record<string, unknown> = {}
): unknown {
  if (expr === null || expr === undefined) {
    return null;
  }

  // If scalar primitive (not expression object), return as-is
  if (typeof expr !== 'object') {
    return expr;
  }

  const exp = expr as ComputedExpression;
  if (!exp.op) {
    return expr;
  }

  const rawArgs = exp.args || [];
  const evaluatedArgs = rawArgs.map((arg) => resolveOperandValue(row, arg, context));

  switch (exp.op) {
    // Math operations (Decimal precision)
    case 'add': {
      let sum = new Decimal(0);
      for (const a of evaluatedArgs) {
        if (a !== null && a !== undefined) {
          sum = sum.plus(toDecimal(a));
        }
      }
      return toSafeNumericResult(sum);
    }

    case 'sub': {
      if (evaluatedArgs.length === 0) return 0;
      let diff = toDecimal(evaluatedArgs[0]);
      for (let i = 1; i < evaluatedArgs.length; i++) {
        diff = diff.minus(toDecimal(evaluatedArgs[i]));
      }
      return toSafeNumericResult(diff);
    }

    case 'mul': {
      if (evaluatedArgs.length === 0) return 0;
      let prod = new Decimal(1);
      for (const a of evaluatedArgs) {
        if (a === null || a === undefined) return 0;
        prod = prod.times(toDecimal(a));
      }
      return toSafeNumericResult(prod);
    }

    case 'div': {
      if (evaluatedArgs.length < 2) return null;
      const num = toDecimal(evaluatedArgs[0]);
      const denom = toDecimal(evaluatedArgs[1]);
      if (denom.isZero()) {
        throw new Error('Division by zero in transform expression');
      }
      return toSafeNumericResult(num.dividedBy(denom));
    }

    case 'mod': {
      if (evaluatedArgs.length < 2) return null;
      const num = toDecimal(evaluatedArgs[0]);
      const denom = toDecimal(evaluatedArgs[1]);
      if (denom.isZero()) {
        throw new Error('Modulo by zero in transform expression');
      }
      return toSafeNumericResult(num.mod(denom));
    }

    case 'abs': {
      if (evaluatedArgs.length === 0) return null;
      return toSafeNumericResult(toDecimal(evaluatedArgs[0]).abs());
    }

    case 'neg': {
      if (evaluatedArgs.length === 0) return null;
      return toSafeNumericResult(toDecimal(evaluatedArgs[0]).negated());
    }

    case 'round': {
      if (evaluatedArgs.length === 0) return null;
      const val = toDecimal(evaluatedArgs[0]);
      const places = exp.places !== undefined ? exp.places : (evaluatedArgs[1] !== undefined ? Number(evaluatedArgs[1]) : 0);
      return toSafeNumericResult(val.toDecimalPlaces(places, Decimal.ROUND_HALF_UP));
    }

    case 'ceil': {
      if (evaluatedArgs.length === 0) return null;
      const val = toDecimal(evaluatedArgs[0]);
      return toSafeNumericResult(val.ceil());
    }

    case 'floor': {
      if (evaluatedArgs.length === 0) return null;
      const val = toDecimal(evaluatedArgs[0]);
      return toSafeNumericResult(val.floor());
    }

    case 'min': {
      if (evaluatedArgs.length === 0) return null;
      let minVal = toDecimal(evaluatedArgs[0]);
      for (let i = 1; i < evaluatedArgs.length; i++) {
        const d = toDecimal(evaluatedArgs[i]);
        if (d.lessThan(minVal)) minVal = d;
      }
      return toSafeNumericResult(minVal);
    }

    case 'max': {
      if (evaluatedArgs.length === 0) return null;
      let maxVal = toDecimal(evaluatedArgs[0]);
      for (let i = 1; i < evaluatedArgs.length; i++) {
        const d = toDecimal(evaluatedArgs[i]);
        if (d.greaterThan(maxVal)) maxVal = d;
      }
      return toSafeNumericResult(maxVal);
    }

    // String operations
    case 'concat': {
      return evaluatedArgs.map((a) => (a === null || a === undefined ? '' : String(a))).join('');
    }

    case 'trim': {
      return evaluatedArgs.length > 0 && evaluatedArgs[0] !== null ? String(evaluatedArgs[0]).trim() : '';
    }

    case 'lower': {
      return evaluatedArgs.length > 0 && evaluatedArgs[0] !== null ? String(evaluatedArgs[0]).toLowerCase() : '';
    }

    case 'upper': {
      return evaluatedArgs.length > 0 && evaluatedArgs[0] !== null ? String(evaluatedArgs[0]).toUpperCase() : '';
    }

    case 'substring': {
      if (evaluatedArgs.length === 0 || evaluatedArgs[0] === null) return '';
      const str = String(evaluatedArgs[0]);
      const start = Number(evaluatedArgs[1] || 0);
      const end = evaluatedArgs[2] !== undefined ? Number(evaluatedArgs[2]) : undefined;
      return str.substring(start, end);
    }

    case 'replace': {
      if (evaluatedArgs.length < 3 || evaluatedArgs[0] === null) return '';
      const str = String(evaluatedArgs[0]);
      const find = String(evaluatedArgs[1]);
      const replaceWith = String(evaluatedArgs[2]);
      return str.split(find).join(replaceWith);
    }

    // Logic operations
    case 'if': {
      const condResult = exp.condition ? evaluateCondition(row, exp.condition, context) : Boolean(evaluatedArgs[0]);
      if (condResult) {
        return exp.then !== undefined ? evaluateExpression(row, exp.then, context) : evaluatedArgs[1];
      } else {
        return exp.else !== undefined ? evaluateExpression(row, exp.else, context) : (evaluatedArgs[2] !== undefined ? evaluatedArgs[2] : null);
      }
    }

    case 'coalesce': {
      for (const arg of evaluatedArgs) {
        if (arg !== null && arg !== undefined && arg !== '') {
          return arg;
        }
      }
      return null;
    }

    case 'lookup': {
      // Lookup value in a table/map in context
      const keyVal = evaluatedArgs[0];
      const tableName = exp.field || (evaluatedArgs[1] ? String(evaluatedArgs[1]) : '');
      const lookupData = context[tableName];
      if (lookupData && typeof lookupData === 'object') {
        const tableRecord = lookupData as Record<string, unknown>;
        return tableRecord[String(keyVal)] !== undefined ? tableRecord[String(keyVal)] : null;
      }
      return null;
    }

    default:
      throw new Error(`Unsupported expression operator: "${exp.op}"`);
  }
}

/**
 * Resolves an operand value which might be a field reference string, sub-expression, or literal.
 */
function resolveOperandValue(
  row: Record<string, unknown>,
  operand: unknown,
  context: Record<string, unknown>
): unknown {
  if (operand === null || operand === undefined) {
    return null;
  }

  if (typeof operand === 'object') {
    return evaluateExpression(row, operand, context);
  }

  if (typeof operand === 'string') {
    // If operand starts with '$', it's an explicit field reference or context reference
    if (operand.startsWith('$.')) {
      const fieldName = operand.slice(2);
      return row[fieldName];
    }
    if (operand.startsWith('$ctx.')) {
      const ctxKey = operand.slice(5);
      return context[ctxKey];
    }
    // If row has exact key matching operand string, resolve field value
    if (Object.prototype.hasOwnProperty.call(row, operand)) {
      return row[operand];
    }
    // Otherwise it's a literal string
    return operand;
  }

  return operand;
}

/**
 * Applies a lookup join against another dataset in context.
 */
function applyLookupJoin(
  data: Record<string, unknown>[],
  step: LookupJoinStep,
  context: Record<string, unknown>
): Record<string, unknown>[] {
  const lookupDataset = context[step.lookupTable];
  if (!Array.isArray(lookupDataset)) {
    throw new Error(`Lookup dataset "${step.lookupTable}" not found in context or is not an array`);
  }

  const lookupIndex = new Map<string, Record<string, unknown>>();
  for (const item of lookupDataset) {
    if (item && typeof item === 'object') {
      const key = String(item[step.lookupKey]);
      lookupIndex.set(key, item as Record<string, unknown>);
    }
  }

  const result: Record<string, unknown>[] = [];

  for (const row of data) {
    const sourceKeyVal = String(row[step.sourceKey]);
    const matched = lookupIndex.get(sourceKeyVal);

    if (!matched && step.joinType === 'inner') {
      // Omit row on inner join failure
      continue;
    }

    const merged: Record<string, unknown> = Object.create(null);
    Object.assign(merged, row);

    for (const [targetField, lookupField] of Object.entries(step.select)) {
      if (matched && matched[lookupField] !== undefined) {
        merged[targetField] = matched[lookupField];
      } else if (step.defaultValues && step.defaultValues[targetField] !== undefined) {
        merged[targetField] = step.defaultValues[targetField];
      } else {
        merged[targetField] = null;
      }
    }

    result.push(merged);
  }

  return result;
}

/**
 * Aggregates dataset rows by groupBy keys.
 */
function applyGroupSum(
  data: Record<string, unknown>[],
  step: GroupSumStep
): Record<string, unknown>[] {
  interface GroupBucket {
    keys: Record<string, unknown>;
    rows: Record<string, unknown>[];
  }

  const groups = new Map<string, GroupBucket>();

  for (const row of data) {
    const groupKey = step.groupBy.map((g) => String(row[g] ?? '')).join('\u0000');
    let bucket = groups.get(groupKey);
    if (!bucket) {
      const keysObj: Record<string, unknown> = Object.create(null);
      for (const g of step.groupBy) {
        keysObj[g] = row[g];
      }
      bucket = { keys: keysObj, rows: [] };
      groups.set(groupKey, bucket);
    }
    bucket.rows.push(row);
  }

  const result: Record<string, unknown>[] = [];

  for (const bucket of groups.values()) {
    const outRow: Record<string, unknown> = Object.create(null);
    Object.assign(outRow, bucket.keys);

    for (const agg of step.aggregations) {
      outRow[agg.as] = computeAggregation(bucket.rows, agg);
    }

    result.push(outRow);
  }

  return result;
}

/**
 * Computes a single aggregation metric over a group of rows using Decimal.
 */
function computeAggregation(rows: Record<string, unknown>[], agg: AggregationDef): unknown {
  if (rows.length === 0) {
    return agg.op === 'count' || agg.op === 'countDistinct' ? 0 : null;
  }

  switch (agg.op) {
    case 'count': {
      if (agg.field === '*' || !agg.field) {
        return rows.length;
      }
      return rows.filter((r) => r[agg.field] !== null && r[agg.field] !== undefined).length;
    }

    case 'countDistinct': {
      const distinct = new Set<string>();
      for (const r of rows) {
        const v = r[agg.field];
        if (v !== null && v !== undefined) {
          distinct.add(String(v));
        }
      }
      return distinct.size;
    }

    case 'sum': {
      let total = new Decimal(0);
      for (const r of rows) {
        const v = r[agg.field];
        if (v !== null && v !== undefined) {
          total = total.plus(toDecimal(v));
        }
      }
      if (agg.round !== undefined) {
        total = total.toDecimalPlaces(agg.round, Decimal.ROUND_HALF_UP);
      }
      return toSafeNumericResult(total);
    }

    case 'avg': {
      let total = new Decimal(0);
      let count = 0;
      for (const r of rows) {
        const v = r[agg.field];
        if (v !== null && v !== undefined) {
          total = total.plus(toDecimal(v));
          count++;
        }
      }
      if (count === 0) return null;
      let avg = total.dividedBy(count);
      if (agg.round !== undefined) {
        avg = avg.toDecimalPlaces(agg.round, Decimal.ROUND_HALF_UP);
      }
      return toSafeNumericResult(avg);
    }

    case 'min': {
      let minVal: Decimal | null = null;
      for (const r of rows) {
        const v = r[agg.field];
        if (v !== null && v !== undefined) {
          const d = toDecimal(v);
          if (minVal === null || d.lessThan(minVal)) {
            minVal = d;
          }
        }
      }
      return minVal !== null ? toSafeNumericResult(minVal) : null;
    }

    case 'max': {
      let maxVal: Decimal | null = null;
      for (const r of rows) {
        const v = r[agg.field];
        if (v !== null && v !== undefined) {
          const d = toDecimal(v);
          if (maxVal === null || d.greaterThan(maxVal)) {
            maxVal = d;
          }
        }
      }
      return maxVal !== null ? toSafeNumericResult(maxVal) : null;
    }

    default:
      throw new Error(`Unsupported aggregation operator "${agg.op}"`);
  }
}

/**
 * Pivots tabular data by row grouping and column pivot values.
 */
function applyPivot(
  data: Record<string, unknown>[],
  step: PivotStep
): Record<string, unknown>[] {
  const rowBuckets = new Map<string, { keys: Record<string, unknown>; rows: Record<string, unknown>[] }>();
  const discoveredColumns = new Set<string>();

  for (const row of data) {
    const rowKey = step.rowGroupBy.map((g) => String(row[g] ?? '')).join('\u0000');
    let bucket = rowBuckets.get(rowKey);
    if (!bucket) {
      const keysObj: Record<string, unknown> = Object.create(null);
      for (const g of step.rowGroupBy) {
        keysObj[g] = row[g];
      }
      bucket = { keys: keysObj, rows: [] };
      rowBuckets.set(rowKey, bucket);
    }
    bucket.rows.push(row);

    const pivotVal = row[step.columnPivot];
    if (pivotVal !== null && pivotVal !== undefined) {
      discoveredColumns.add(String(pivotVal));
    }
  }

  const pivotColumns = step.columnValues && step.columnValues.length > 0
    ? step.columnValues
    : Array.from(discoveredColumns).sort();

  const result: Record<string, unknown>[] = [];

  for (const bucket of rowBuckets.values()) {
    const outRow: Record<string, unknown> = Object.create(null);
    Object.assign(outRow, bucket.keys);

    for (const col of pivotColumns) {
      const matchingRows = bucket.rows.filter((r) => String(r[step.columnPivot] ?? '') === col);
      if (step.aggOp === 'count') {
        outRow[col] = matchingRows.length;
      } else {
        // Default sum
        let sum = new Decimal(0);
        for (const mr of matchingRows) {
          const v = mr[step.valueField];
          if (v !== null && v !== undefined) {
            sum = sum.plus(toDecimal(v));
          }
        }
        outRow[col] = toSafeNumericResult(sum);
      }
    }

    result.push(outRow);
  }

  return result;
}

/**
 * Sorts data by multiple fields with natural ordering.
 */
function applySort(
  data: Record<string, unknown>[],
  step: SortStep
): Record<string, unknown>[] {
  const sorted = [...data];

  sorted.sort((a, b) => {
    for (const f of step.fields) {
      const valA = a[f.field];
      const valB = b[f.field];
      const cmp = compareOrder(valA, valB);
      if (cmp !== 0) {
        return f.order === 'desc' ? -cmp : cmp;
      }
    }
    return 0;
  });

  return sorted;
}

/**
 * Helper to convert any value safely to Decimal.
 */
function toDecimal(val: unknown): Decimal {
  if (val instanceof Decimal) return val;
  if (typeof val === 'number') {
    if (!Number.isFinite(val)) {
      throw new Error(`Cannot convert non-finite number ${val} to Decimal`);
    }
    return new Decimal(val);
  }
  if (typeof val === 'string') {
    const clean = val.trim();
    if (clean === '') return new Decimal(0);
    return new Decimal(clean);
  }
  if (typeof val === 'boolean') {
    return new Decimal(val ? 1 : 0);
  }
  return new Decimal(0);
}

/**
 * Returns standard JS number if safe, otherwise Decimal string.
 */
function toSafeNumericResult(dec: Decimal): number | string {
  const num = dec.toNumber();
  if (Number.isFinite(num) && dec.equals(new Decimal(num))) {
    return num;
  }
  return dec.toString();
}

/**
 * Compares two values for equality.
 */
function compareEquality(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (a === null || a === undefined || b === null || b === undefined) {
    return a === b;
  }

  // If numbers or numeric strings
  if (isNumeric(a) && isNumeric(b)) {
    try {
      return toDecimal(a).equals(toDecimal(b));
    } catch {
      return false;
    }
  }

  return String(a) === String(b);
}

/**
 * Compares two values for ordering (<, ==, >).
 */
function compareOrder(a: unknown, b: unknown): number {
  if (a === b) return 0;
  if (a === null || a === undefined) return -1;
  if (b === null || b === undefined) return 1;

  if (isNumeric(a) && isNumeric(b)) {
    try {
      const decA = toDecimal(a);
      const decB = toDecimal(b);
      return decA.comparedTo(decB);
    } catch {
      // Fallback to string comparison
    }
  }

  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}

function isNumeric(val: unknown): boolean {
  if (typeof val === 'number') return Number.isFinite(val);
  if (typeof val === 'string') {
    const trimmed = val.trim();
    return trimmed !== '' && !isNaN(Number(trimmed));
  }
  return false;
}
