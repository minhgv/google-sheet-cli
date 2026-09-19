/**
 * Manpower Estimation, Role Rates, and Scenario Modeling Preset.
 */

import Decimal from 'decimal.js';
import {
  ManpowerConfig,
  ManpowerReportSummary,
  ManpowerTask,
  ReportCell,
  ReportDocument,
  ReportNumberFormat,
  ReportSheet,
  ReportTemplateV1,
  RoleRate,
  ScenarioConfig,
} from './types';
/**
 * Maximum significant digits representable without precision loss in standard spreadsheet IEEE-754 numbers.
 */
const SPREADSHEET_MAX_SIGNIFICANT_DIGITS = 15;

/**
 * Safely converts a Decimal to a spreadsheet cell numeric value.
 * Validates exact IEEE-754 double precision round-tripping and spreadsheet 15-significant-digit limit.
 * Fails explicitly if the effort/cost amount exceeds spreadsheet precision limits to prevent formula calculation drift.
 */
export function toSafeNumericCell(dec: Decimal, decimalPlaces = 2, context?: string): number {
  if (!dec.isFinite()) {
    const ctxStr = context ? ` for ${context}` : '';
    throw new Error(`Cannot output non-finite Decimal (NaN or Infinity)${ctxStr} to report cell`);
  }
  const rounded = dec.toDecimalPlaces(decimalPlaces, Decimal.ROUND_HALF_UP);
  const num = rounded.toNumber();

  if (rounded.precision(true) > SPREADSHEET_MAX_SIGNIFICANT_DIGITS || !new Decimal(num).equals(rounded)) {
    const ctxStr = context ? ` for ${context}` : '';
    throw new Error(
      `Amount "${rounded.toString()}"${ctxStr} exceeds standard spreadsheet numeric precision (${SPREADSHEET_MAX_SIGNIFICANT_DIGITS} significant digits) and cannot be represented exactly in spreadsheet calculations.`
    );
  }

  return num;
}

/**
 * Builds a complete multi-sheet Manpower & Effort Estimation Report.
 */
export function buildManpowerReport(
  template: ReportTemplateV1,
  input: unknown,
  sourceHash: string
): ReportDocument {
  const config = (template.config || {}) as unknown as ManpowerConfig;
  const currency = config.currency || 'USD';
  const effortUnit = config.effortUnit || 'person-days';
  const globalContingency = new Decimal(config.globalContingencyPercent !== undefined ? config.globalContingencyPercent : 20);

  // Validate and parse role rates mapping
  const roleRatesMap = parseRoleRates(config.roleRates);

  // Validate and parse tasks
  const tasks = parseAndValidateManpowerTasks(input, template.policy?.maxInputRows);

  // Process tasks with Decimal arithmetic
  interface ProcessedTask {
    id: string;
    module: string;
    role: string;
    description: string;
    quantity: Decimal;
    unitEffort: Decimal;
    complexity: Decimal;
    baseEffort: Decimal;
    contingencyPercent: Decimal;
    totalEffort: Decimal;
    dailyRate: Decimal;
    totalCost: Decimal;
  }

  const processedTasks: ProcessedTask[] = [];
  const modulesSet = new Set<string>();
  const rolesSet = new Set<string>();

  for (let i = 0; i < tasks.length; i++) {
    const t = tasks[i];
    modulesSet.add(t.module);
    rolesSet.add(t.role);

    const rate = roleRatesMap.get(t.role);
    if (rate === undefined) {
      throw new Error(`Missing role rate for role "${t.role}" (task ID: "${t.id}")`);
    }

    const qty = new Decimal(t.quantity);
    const unitEff = new Decimal(t.unitEffort);
    const comp = new Decimal(t.complexity !== undefined ? t.complexity : 1);
    const contPercent = new Decimal(t.contingency !== undefined ? t.contingency : globalContingency);

    const baseEff = qty.times(unitEff).times(comp).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
    const contMultiplier = new Decimal(1).plus(contPercent.dividedBy(100));
    const totalEff = baseEff.times(contMultiplier).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
    const rateDec = new Decimal(rate);
    const totalCost = totalEff.times(rateDec).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);

    processedTasks.push({
      id: t.id,
      module: t.module,
      role: t.role,
      description: t.description || '',
      quantity: qty,
      unitEffort: unitEff,
      complexity: comp,
      baseEffort: baseEff,
      contingencyPercent: contPercent,
      totalEffort: totalEff,
      dailyRate: rateDec,
      totalCost,
    });
  }

  // Grand totals
  let grandTotalBaseEffort = new Decimal(0);
  let grandTotalEffort = new Decimal(0);
  let grandTotalCost = new Decimal(0);

  for (const pt of processedTasks) {
    grandTotalBaseEffort = grandTotalBaseEffort.plus(pt.baseEffort);
    grandTotalEffort = grandTotalEffort.plus(pt.totalEffort);
    grandTotalCost = grandTotalCost.plus(pt.totalCost);
  }

  // Aggregate by Module
  const sortedModules = Array.from(modulesSet).sort();
  interface ModuleAgg {
    module: string;
    taskCount: number;
    baseEffort: Decimal;
    totalEffort: Decimal;
    totalCost: Decimal;
    pctCost: Decimal;
  }

  const moduleData: ModuleAgg[] = sortedModules.map((m) => {
    const mTasks = processedTasks.filter((pt) => pt.module === m);
    let mBase = new Decimal(0);
    let mTotalEff = new Decimal(0);
    let mCost = new Decimal(0);
    for (const t of mTasks) {
      mBase = mBase.plus(t.baseEffort);
      mTotalEff = mTotalEff.plus(t.totalEffort);
      mCost = mCost.plus(t.totalCost);
    }
    const pct = grandTotalCost.isZero()
      ? new Decimal(0)
      : mCost.dividedBy(grandTotalCost).times(100).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);

    return {
      module: m,
      taskCount: mTasks.length,
      baseEffort: mBase,
      totalEffort: mTotalEff,
      totalCost: mCost,
      pctCost: pct,
    };
  });

  // Aggregate by Role
  const sortedRoles = Array.from(rolesSet).sort();
  interface RoleAgg {
    role: string;
    taskCount: number;
    totalEffort: Decimal;
    dailyRate: Decimal;
    totalCost: Decimal;
    pctEffort: Decimal;
  }

  const roleData: RoleAgg[] = sortedRoles.map((r) => {
    const rTasks = processedTasks.filter((pt) => pt.role === r);
    let rTotalEff = new Decimal(0);
    let rCost = new Decimal(0);
    const rRate = new Decimal(roleRatesMap.get(r) || 0);

    for (const t of rTasks) {
      rTotalEff = rTotalEff.plus(t.totalEffort);
      rCost = rCost.plus(t.totalCost);
    }
    const pct = grandTotalEffort.isZero()
      ? new Decimal(0)
      : rTotalEff.dividedBy(grandTotalEffort).times(100).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);

    return {
      role: r,
      taskCount: rTasks.length,
      totalEffort: rTotalEff,
      dailyRate: rRate,
      totalCost: rCost,
      pctEffort: pct,
    };
  });

  // Scenarios simulation
  const defaultScenarios: ScenarioConfig[] = [
    { name: 'Optimistic (-10% complexity, 10% contingency)', complexityMultiplier: 0.9, contingencyPercent: 10 },
    { name: 'Base Plan (Standard complexity, 20% contingency)', complexityMultiplier: 1.0, contingencyPercent: 20, isBase: true },
    { name: 'Pessimistic (+25% complexity, 35% contingency)', complexityMultiplier: 1.25, contingencyPercent: 35 },
  ];

  const scenariosList = config.scenarios && config.scenarios.length > 0 ? config.scenarios : defaultScenarios;

  // Validate scenarios: non-empty and unique names
  const seenScenarioNames = new Set<string>();
  for (let i = 0; i < scenariosList.length; i++) {
    const s = scenariosList[i];
    if (!s || typeof s !== 'object' || !s.name || String(s.name).trim() === '') {
      throw new Error(`Scenario at index ${i} must have a non-empty "name" string`);
    }
    const nameStr = String(s.name).trim();
    if (seenScenarioNames.has(nameStr)) {
      throw new Error(`Duplicate scenario name "${nameStr}". Scenario names must be unique.`);
    }
    seenScenarioNames.add(nameStr);
  }

  // Resolve baseline scenario name
  let baselineScenarioName: string | undefined = undefined;
  if (config.baselineScenario) {
    const targetName = String(config.baselineScenario).trim();
    const matched = scenariosList.filter((s) => String(s.name).trim() === targetName);
    if (matched.length === 0) {
      throw new Error(`Configured baselineScenario "${config.baselineScenario}" was not found in scenarios list.`);
    }
    baselineScenarioName = targetName;
  } else {
    const explicitBase = scenariosList.filter((s) => s.isBase === true);
    if (explicitBase.length > 1) {
      throw new Error('Multiple scenarios are marked with isBase: true. At most one baseline scenario is allowed.');
    }
    if (explicitBase.length === 1) {
      baselineScenarioName = String(explicitBase[0].name).trim();
    }
  }

  interface ScenarioResult {
    name: string;
    complexityMultiplier: Decimal;
    contingencyPercent: Decimal;
    totalBaseEffort: Decimal;
    totalEffort: Decimal;
    totalCost: Decimal;
    costDelta: Decimal;
    pctDelta: Decimal;
  }

  const scenarioResults: ScenarioResult[] = [];
  let baseScenarioCost: Decimal = grandTotalCost;

  for (const s of scenariosList) {
    const compMult = new Decimal(s.complexityMultiplier !== undefined ? s.complexityMultiplier : 1);
    const contPct = new Decimal(s.contingencyPercent !== undefined ? s.contingencyPercent : globalContingency);

    let sBaseEff = new Decimal(0);
    let sTotalEff = new Decimal(0);
    let sTotalCost = new Decimal(0);

    for (const pt of processedTasks) {
      const taskBase = pt.quantity.times(pt.unitEffort).times(pt.complexity).times(compMult).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
      const taskTotalEff = taskBase.times(new Decimal(1).plus(contPct.dividedBy(100))).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
      const roleRateMult = s.roleRateMultipliers && s.roleRateMultipliers[pt.role] !== undefined
        ? new Decimal(s.roleRateMultipliers[pt.role])
        : new Decimal(1);
      const effRate = pt.dailyRate.times(roleRateMult);
      const taskCost = taskTotalEff.times(effRate).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);

      sBaseEff = sBaseEff.plus(taskBase);
      sTotalEff = sTotalEff.plus(taskTotalEff);
      sTotalCost = sTotalCost.plus(taskCost);
    }

    if (baselineScenarioName && String(s.name).trim() === baselineScenarioName) {
      baseScenarioCost = sTotalCost;
    }

    scenarioResults.push({
      name: s.name,
      complexityMultiplier: compMult,
      contingencyPercent: contPct,
      totalBaseEffort: sBaseEff,
      totalEffort: sTotalEff,
      totalCost: sTotalCost,
      costDelta: new Decimal(0),
      pctDelta: new Decimal(0),
    });
  }

  // Update deltas vs base
  for (const sr of scenarioResults) {
    sr.costDelta = sr.totalCost.minus(baseScenarioCost);
    sr.pctDelta = baseScenarioCost.isZero()
      ? new Decimal(0)
      : sr.costDelta.dividedBy(baseScenarioCost).times(100).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);
  }

  // 1. Build Assumptions Sheet
  const assumptionsSheet = buildAssumptionsSheet(config, roleRatesMap, scenariosList, currency, effortUnit, globalContingency);

  // 2. Build Task Details Sheet
  const detailsSheet = buildTaskDetailsSheet(processedTasks, currency, effortUnit, grandTotalBaseEffort, grandTotalEffort, grandTotalCost);

  // 3. Build Module Summary Sheet
  const moduleSheet = buildModuleSummarySheet(moduleData, currency, effortUnit, grandTotalBaseEffort, grandTotalEffort, grandTotalCost);

  // 4. Build Role Summary Sheet
  const roleSheet = buildRoleSummarySheet(roleData, currency, effortUnit, grandTotalEffort, grandTotalCost);

  // 5. Build Scenarios Sheet
  const scenarioSheet = buildScenariosSheet(scenarioResults, currency, effortUnit);

  const summary: ManpowerReportSummary & { baselineScenario: string } = {
    effortUnit,
    currency,
    totalBaseEffort: grandTotalBaseEffort.toFixed(2),
    totalEffort: grandTotalEffort.toFixed(2),
    totalCost: grandTotalCost.toFixed(2),
    moduleCount: sortedModules.length,
    roleCount: sortedRoles.length,
    taskCount: processedTasks.length,
    baselineScenario: baselineScenarioName || 'Original Plan',
  };

  return {
    sheets: [assumptionsSheet, detailsSheet, moduleSheet, roleSheet, scenarioSheet],
    provenance: {
      templateId: template.id,
      templateVersion: template.version,
      sourceHash,
      generatedAt: new Date().toISOString(),
      summary: summary as unknown as Record<string, unknown>,
    },
  };
}

/**
 * Validates and parses raw manpower task records.
 */
function parseAndValidateManpowerTasks(
  input: unknown,
  maxRows?: number
): ManpowerTask[] {
  if (!Array.isArray(input)) {
    throw new Error('Manpower report input must be an array of task records');
  }

  const limit = maxRows || 100000;
  if (input.length > limit) {
    throw new Error(`Input tasks count (${input.length}) exceeds maximum allowed limit of ${limit}`);
  }

  const result: ManpowerTask[] = [];

  for (let i = 0; i < input.length; i++) {
    const raw = input[i];
    const rowNum = i + 1;

    if (!raw || typeof raw !== 'object') {
      throw new Error(`Row ${rowNum}: Expected task record object`);
    }

    const item = raw as Record<string, unknown>;

    // id
    if (!item.id || String(item.id).trim() === '') {
      throw new Error(`Row ${rowNum}: Task "id" is required`);
    }
    const id = String(item.id).trim();

    // module
    if (!item.module || String(item.module).trim() === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Task "module" is required`);
    }
    const modName = String(item.module).trim();

    // role
    if (!item.role || String(item.role).trim() === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Task "role" is required`);
    }
    const roleName = String(item.role).trim();

    // quantity
    if (item.quantity === undefined || item.quantity === null || item.quantity === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Task "quantity" is required`);
    }
    let quantityNum: number;
    try {
      const qDec = new Decimal(String(item.quantity));
      if (qDec.lessThanOrEqualTo(0)) {
        throw new Error('Quantity must be greater than zero');
      }
      quantityNum = qDec.toNumber();
    } catch {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Invalid positive quantity "${item.quantity}"`);
    }

    // unitEffort
    if (item.unitEffort === undefined || item.unitEffort === null || item.unitEffort === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Task "unitEffort" is required`);
    }
    let unitEffNum: number;
    try {
      const ueDec = new Decimal(String(item.unitEffort));
      if (ueDec.lessThan(0)) {
        throw new Error('Unit effort cannot be negative');
      }
      unitEffNum = ueDec.toNumber();
    } catch {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Invalid unit effort "${item.unitEffort}"`);
    }

    // complexity
    let compNum: number | undefined;
    if (item.complexity !== undefined && item.complexity !== null && item.complexity !== '') {
      try {
        const cDec = new Decimal(String(item.complexity));
        if (cDec.lessThanOrEqualTo(0)) {
          throw new Error('Complexity must be positive');
        }
        compNum = cDec.toNumber();
      } catch {
        throw new Error(`Row ${rowNum} (ID: "${id}"): Invalid complexity multiplier "${item.complexity}"`);
      }
    }

    // contingency
    let contNum: number | undefined;
    if (item.contingency !== undefined && item.contingency !== null && item.contingency !== '') {
      try {
        const contDec = new Decimal(String(item.contingency));
        if (contDec.lessThan(0)) {
          throw new Error('Contingency percent cannot be negative');
        }
        contNum = contDec.toNumber();
      } catch {
        throw new Error(`Row ${rowNum} (ID: "${id}"): Invalid contingency percent "${item.contingency}"`);
      }
    }

    result.push({
      id,
      module: modName,
      role: roleName,
      quantity: quantityNum,
      unitEffort: unitEffNum,
      complexity: compNum !== undefined ? compNum : 1.0,
      contingency: contNum,
      description: item.description ? String(item.description).trim() : undefined,
    });
  }

  return result;
}

/**
 * Parses role rates from Record or Array.
 */
function parseRoleRates(
  ratesConfig: Record<string, number | string> | RoleRate[] | undefined
): Map<string, number> {
  const map = new Map<string, number>();

  if (!ratesConfig) {
    return map;
  }

  if (Array.isArray(ratesConfig)) {
    for (const r of ratesConfig) {
      if (r && r.role) {
        map.set(r.role, new Decimal(r.dailyRate || 0).toNumber());
      }
    }
  } else if (typeof ratesConfig === 'object') {
    for (const [role, rate] of Object.entries(ratesConfig)) {
      map.set(role, new Decimal(rate || 0).toNumber());
    }
  }

  return map;
}

/**
 * Builds Assumptions worksheet.
 */
function buildAssumptionsSheet(
  config: ManpowerConfig,
  roleRatesMap: Map<string, number>,
  scenariosList: ScenarioConfig[],
  currency: string,
  effortUnit: string,
  globalContingency: Decimal
): ReportSheet {
  const rows: ReportCell[][] = [
    ['Manpower Planning Parameters & Assumptions', null, null],
    ['Currency', currency, null],
    ['Effort Unit', effortUnit, null],
    ['Global Contingency Buffer (%)', globalContingency.toNumber() / 100, null],
    ['Standard Working Days/Month', config.baseWorkingDaysPerMonth || 22, null],
    [null, null, null],
    ['Standard Role Rates', null, null],
    ['Role', `Daily Rate (${currency})`, 'Effort Unit'],
  ];

  for (const [role, rate] of roleRatesMap.entries()) {
    rows.push([role, rate, effortUnit]);
  }

  rows.push([null, null, null]);
  rows.push(['Scenario Simulations Configuration', null, null]);
  rows.push(['Scenario Name', 'Complexity Multiplier', 'Contingency %']);

  for (const s of scenariosList) {
    const comp = new Decimal(s.complexityMultiplier !== undefined ? s.complexityMultiplier : 1).toNumber();
    const cont = new Decimal(s.contingencyPercent !== undefined ? s.contingencyPercent : globalContingency).toNumber() / 100;
    rows.push([s.name, comp, cont]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '#,##0.00' },
    { column: 3, format: '0.0%' },
  ];

  return {
    name: 'Assumptions',
    rows,
    freezeRows: 1,
    columnWidths: [32, 28, 20],
    numberFormats,
  };
}

/**
 * Builds Task Details worksheet with embedded formulas and pre-calculated results.
 */
function buildTaskDetailsSheet(
  tasks: {
    id: string;
    module: string;
    role: string;
    description: string;
    quantity: Decimal;
    unitEffort: Decimal;
    complexity: Decimal;
    baseEffort: Decimal;
    contingencyPercent: Decimal;
    totalEffort: Decimal;
    dailyRate: Decimal;
    totalCost: Decimal;
  }[],
  currency: string,
  effortUnit: string,
  grandBaseEffort: Decimal,
  grandTotalEffort: Decimal,
  grandTotalCost: Decimal
): ReportSheet {
  const headers: ReportCell[] = [
    'Task ID',
    'Module',
    'Role',
    'Description',
    'Quantity',
    `Unit Effort (${effortUnit})`,
    'Complexity',
    `Base Effort (${effortUnit})`,
    'Contingency %',
    `Total Effort (${effortUnit})`,
    `Daily Rate (${currency})`,
    `Total Cost (${currency})`,
  ];

  const rows: ReportCell[][] = [headers];

  for (let i = 0; i < tasks.length; i++) {
    const t = tasks[i];
    const excelRow = i + 2;

    const baseEffNum = toSafeNumericCell(t.baseEffort, 2);
    const totalEffNum = toSafeNumericCell(t.totalEffort, 2);
    const totalCostNum = toSafeNumericCell(t.totalCost, 2);
    const contDecimalRatio = t.contingencyPercent.toNumber() / 100;

    rows.push([
      t.id,
      t.module,
      t.role,
      t.description,
      toSafeNumericCell(t.quantity, 2),
      toSafeNumericCell(t.unitEffort, 2),
      toSafeNumericCell(t.complexity, 2),
      { formula: `E${excelRow}*F${excelRow}*G${excelRow}`, result: baseEffNum },
      contDecimalRatio,
      { formula: `H${excelRow}*(1+I${excelRow})`, result: totalEffNum },
      toSafeNumericCell(t.dailyRate, 2),
      { formula: `J${excelRow}*K${excelRow}`, result: totalCostNum },
    ]);
  }

  if (tasks.length > 0) {
    const lastDataRow = tasks.length + 1;
    rows.push([
      'Total',
      null,
      null,
      null,
      null,
      null,
      null,
      { formula: `SUM(H2:H${lastDataRow})`, result: toSafeNumericCell(grandBaseEffort, 2) },
      null,
      { formula: `SUM(J2:J${lastDataRow})`, result: toSafeNumericCell(grandTotalEffort, 2) },
      null,
      { formula: `SUM(L2:L${lastDataRow})`, result: toSafeNumericCell(grandTotalCost, 2) },
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 5, format: '#,##0.00' },
    { column: 6, format: '#,##0.00' },
    { column: 7, format: '0.00' },
    { column: 8, format: '#,##0.00' },
    { column: 9, format: '0.0%' },
    { column: 10, format: '#,##0.00' },
    { column: 11, format: '#,##0.00' },
    { column: 12, format: '#,##0.00' },
  ];

  return {
    name: 'Task Details',
    rows,
    freezeRows: 1,
    columnWidths: [16, 18, 18, 30, 12, 18, 14, 20, 16, 20, 20, 22],
    numberFormats,
  };
}

/**
 * Builds Module Summary worksheet.
 */
function buildModuleSummarySheet(
  modules: {
    module: string;
    taskCount: number;
    baseEffort: Decimal;
    totalEffort: Decimal;
    totalCost: Decimal;
    pctCost: Decimal;
  }[],
  currency: string,
  effortUnit: string,
  grandBaseEffort: Decimal,
  grandTotalEffort: Decimal,
  grandTotalCost: Decimal
): ReportSheet {
  const headers: ReportCell[] = [
    'Module',
    'Task Count',
    `Base Effort (${effortUnit})`,
    `Total Effort (${effortUnit})`,
    `Total Cost (${currency})`,
    '% of Total Cost',
  ];

  const rows: ReportCell[][] = [headers];
  const totalRowIndex = modules.length + 2;

  for (let i = 0; i < modules.length; i++) {
    const m = modules[i];
    const excelRow = i + 2;

    rows.push([
      m.module,
      m.taskCount,
      toSafeNumericCell(m.baseEffort, 2),
      toSafeNumericCell(m.totalEffort, 2),
      toSafeNumericCell(m.totalCost, 2),
      { formula: `E${excelRow}/$E$${totalRowIndex}`, result: m.pctCost.toNumber() / 100 },
    ]);
  }

  if (modules.length > 0) {
    const lastDataRow = modules.length + 1;
    rows.push([
      'Total',
      { formula: `SUM(B2:B${lastDataRow})`, result: modules.reduce((acc, m) => acc + m.taskCount, 0) },
      { formula: `SUM(C2:C${lastDataRow})`, result: toSafeNumericCell(grandBaseEffort, 2) },
      { formula: `SUM(D2:D${lastDataRow})`, result: toSafeNumericCell(grandTotalEffort, 2) },
      { formula: `SUM(E2:E${lastDataRow})`, result: toSafeNumericCell(grandTotalCost, 2) },
      1.0,
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '#,##0' },
    { column: 3, format: '#,##0.00' },
    { column: 4, format: '#,##0.00' },
    { column: 5, format: '#,##0.00' },
    { column: 6, format: '0.0%' },
  ];

  return {
    name: 'Module Summary',
    rows,
    freezeRows: 1,
    columnWidths: [24, 14, 22, 22, 24, 18],
    numberFormats,
  };
}

/**
 * Builds Role Summary worksheet.
 */
function buildRoleSummarySheet(
  roles: {
    role: string;
    taskCount: number;
    totalEffort: Decimal;
    dailyRate: Decimal;
    totalCost: Decimal;
    pctEffort: Decimal;
  }[],
  currency: string,
  effortUnit: string,
  grandTotalEffort: Decimal,
  grandTotalCost: Decimal
): ReportSheet {
  const headers: ReportCell[] = [
    'Role',
    'Task Count',
    `Total Effort (${effortUnit})`,
    `Daily Rate (${currency})`,
    `Total Cost (${currency})`,
    '% of Total Effort',
  ];

  const rows: ReportCell[][] = [headers];
  const totalRowIndex = roles.length + 2;

  for (let i = 0; i < roles.length; i++) {
    const r = roles[i];
    const excelRow = i + 2;

    rows.push([
      r.role,
      r.taskCount,
      toSafeNumericCell(r.totalEffort, 2),
      toSafeNumericCell(r.dailyRate, 2),
      toSafeNumericCell(r.totalCost, 2),
      { formula: `C${excelRow}/$C$${totalRowIndex}`, result: r.pctEffort.toNumber() / 100 },
    ]);
  }

  if (roles.length > 0) {
    const lastDataRow = roles.length + 1;
    rows.push([
      'Total',
      { formula: `SUM(B2:B${lastDataRow})`, result: roles.reduce((acc, r) => acc + r.taskCount, 0) },
      { formula: `SUM(C2:C${lastDataRow})`, result: toSafeNumericCell(grandTotalEffort, 2) },
      null,
      { formula: `SUM(E2:E${lastDataRow})`, result: toSafeNumericCell(grandTotalCost, 2) },
      1.0,
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '#,##0' },
    { column: 3, format: '#,##0.00' },
    { column: 4, format: '#,##0.00' },
    { column: 5, format: '#,##0.00' },
    { column: 6, format: '0.0%' },
  ];

  return {
    name: 'Role Summary',
    rows,
    freezeRows: 1,
    columnWidths: [22, 14, 22, 22, 24, 18],
    numberFormats,
  };
}

/**
 * Builds Scenarios comparison worksheet.
 */
function buildScenariosSheet(
  scenarios: {
    name: string;
    complexityMultiplier: Decimal;
    contingencyPercent: Decimal;
    totalBaseEffort: Decimal;
    totalEffort: Decimal;
    totalCost: Decimal;
    costDelta: Decimal;
    pctDelta: Decimal;
  }[],
  currency: string,
  effortUnit: string
): ReportSheet {
  const headers: ReportCell[] = [
    'Scenario',
    'Complexity Multiplier',
    'Contingency %',
    `Base Effort (${effortUnit})`,
    `Total Effort (${effortUnit})`,
    `Total Cost (${currency})`,
    `Cost Delta vs Base (${currency})`,
    '% Delta vs Base',
  ];

  const rows: ReportCell[][] = [headers];

  for (let i = 0; i < scenarios.length; i++) {
    const s = scenarios[i];

    rows.push([
      s.name,
      toSafeNumericCell(s.complexityMultiplier, 2),
      s.contingencyPercent.toNumber() / 100,
      toSafeNumericCell(s.totalBaseEffort, 2),
      toSafeNumericCell(s.totalEffort, 2),
      toSafeNumericCell(s.totalCost, 2),
      toSafeNumericCell(s.costDelta, 2),
      s.pctDelta.toNumber() / 100,
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '0.00' },
    { column: 3, format: '0.0%' },
    { column: 4, format: '#,##0.00' },
    { column: 5, format: '#,##0.00' },
    { column: 6, format: '#,##0.00' },
    { column: 7, format: '+#,##0.00;-#,##0.00;0.00' },
    { column: 8, format: '+0.0%;-0.0%;0.0%' },
  ];

  return {
    name: 'Scenarios',
    rows,
    freezeRows: 1,
    columnWidths: [32, 22, 18, 22, 22, 24, 26, 18],
    numberFormats,
  };
}
