/**
 * Financial Cash Flow, FX Conversion, and Account Reconciliation Preset.
 */

import Decimal from 'decimal.js';
import {
  FinanceConfig,
  FinanceOpeningBalance,
  FinanceReportSummary,
  FinanceTransaction,
  FxRate,
  ReportCell,
  ReportDocument,
  ReportNumberFormat,
  ReportSheet,
  ReportTemplateV1,
} from './types';
/**
 * Maximum significant digits representable without precision loss in standard spreadsheet IEEE-754 numbers.
 */
const SPREADSHEET_MAX_SIGNIFICANT_DIGITS = 15;

/**
 * Safely converts a Decimal to a spreadsheet cell numeric value.
 * Validates exact IEEE-754 double precision round-tripping and spreadsheet 15-significant-digit limit.
 * Fails explicitly if the monetary amount exceeds spreadsheet precision limits to prevent formula calculation drift.
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
      `Monetary amount "${rounded.toString()}"${ctxStr} exceeds standard spreadsheet numeric precision (${SPREADSHEET_MAX_SIGNIFICANT_DIGITS} significant digits) and cannot be represented exactly in spreadsheet calculations.`
    );
  }

  return num;
}

/**
 * Builds a complete multi-sheet Financial Report from transactions and configuration.
 */
export function buildFinanceReport(
  template: ReportTemplateV1,
  input: unknown,
  sourceHash: string
): ReportDocument {
  const config = (template.config || {}) as unknown as FinanceConfig;
  const baseCurrency = config.baseCurrency;
  if (!baseCurrency || typeof baseCurrency !== 'string') {
    throw new Error('Finance report template must specify "config.baseCurrency"');
  }

  // Validate and extract transactions
  const transactions = parseAndValidateFinanceTransactions(input, template.policy?.maxInputRows);

  // Check for duplicate transaction IDs
  const seenIds = new Map<string, number>();
  for (let i = 0; i < transactions.length; i++) {
    const tx = transactions[i];
    if (seenIds.has(tx.id)) {
      const prevRow = seenIds.get(tx.id)!;
      throw new Error(`Duplicate transaction ID: "${tx.id}" found at row ${i + 1} (previously at row ${prevRow})`);
    }
    seenIds.set(tx.id, i + 1);
  }

  // Filter by date window if configured
  const dateWindow = config.dateWindow;
  const filteredTransactions = transactions.filter((tx) => {
    if (dateWindow?.from && tx.date < dateWindow.from) return false;
    if (dateWindow?.to && tx.date > dateWindow.to) return false;
    return true;
  });

  // Prepare FX rate lookup table
  const fxMap = buildFxRateIndex(config.fxRates || []);

  // Opening balances mapping
  const openingBalances = normalizeOpeningBalances(config.openingBalances);
  const expectedClosingBalances = config.expectedClosingBalances || {};

  // Process transactions with Decimal arithmetic
  interface ProcessedTx {
    id: string;
    date: string;
    account: string;
    description: string;
    category: string;
    direction: 'in' | 'out';
    originalAmount: Decimal;
    currency: string;
    fxRate: Decimal;
    baseAmount: Decimal;
    inflow: Decimal;
    outflow: Decimal;
    netFlow: Decimal;
  }

  const processedList: ProcessedTx[] = [];
  const accountsSet = new Set<string>();
  const monthsSet = new Set<string>();

  for (let i = 0; i < filteredTransactions.length; i++) {
    const tx = filteredTransactions[i];
    accountsSet.add(tx.account);

    const monthStr = tx.date.substring(0, 7);
    if (monthStr.length === 7) {
      monthsSet.add(monthStr);
    }

    const txCurrency = tx.currency.toUpperCase();
    let fxRate: Decimal;

    if (txCurrency === baseCurrency.toUpperCase()) {
      fxRate = new Decimal(1);
    } else {
      const lookupRate = findFxRate(fxMap, tx.date, txCurrency, baseCurrency.toUpperCase());
      if (lookupRate === null) {
        throw new Error(
          `Missing FX rate for currency pair ${txCurrency}->${baseCurrency} on date ${tx.date} (transaction ID: "${tx.id}")`
        );
      }
      fxRate = new Decimal(lookupRate);
    }

    const originalAmount = new Decimal(tx.amount);
    const baseAmount = originalAmount.times(fxRate).toDecimalPlaces(2, Decimal.ROUND_HALF_UP);

    const isRefund = tx.isRefund === true || String(tx.isRefund).toLowerCase() === 'true';
    const directionNorm = tx.direction.toLowerCase() === 'in' ? 'in' : 'out';

    let inflow = new Decimal(0);
    let outflow = new Decimal(0);

    if (directionNorm === 'in') {
      if (isRefund) {
        inflow = baseAmount.negated();
      } else {
        inflow = baseAmount;
      }
    } else {
      if (isRefund) {
        outflow = baseAmount.negated();
      } else {
        outflow = baseAmount;
      }
    }

    const netFlow = inflow.minus(outflow);

    processedList.push({
      id: tx.id,
      date: tx.date,
      account: tx.account,
      description: tx.description || '',
      category: tx.category || 'General',
      direction: directionNorm,
      originalAmount,
      currency: txCurrency,
      fxRate,
      baseAmount,
      inflow,
      outflow,
      netFlow,
    });
  }

  // Include accounts from openingBalances that may have had no transactions
  for (const acc of Object.keys(openingBalances)) {
    accountsSet.add(acc);
  }

  // Aggregate by month
  const sortedMonths = Array.from(monthsSet).sort();
  interface MonthlyAgg {
    month: string;
    inflow: Decimal;
    outflow: Decimal;
    netFlow: Decimal;
  }
  const monthlyData: MonthlyAgg[] = sortedMonths.map((m) => {
    const txsInMonth = processedList.filter((tx) => tx.date.startsWith(m));
    let monthInflow = new Decimal(0);
    let monthOutflow = new Decimal(0);
    for (const t of txsInMonth) {
      monthInflow = monthInflow.plus(t.inflow);
      monthOutflow = monthOutflow.plus(t.outflow);
    }
    return {
      month: m,
      inflow: monthInflow,
      outflow: monthOutflow,
      netFlow: monthInflow.minus(monthOutflow),
    };
  });

  // Aggregate by account
  const sortedAccounts = Array.from(accountsSet).sort();
  interface AccountAgg {
    account: string;
    openingBalance: Decimal;
    inflow: Decimal;
    outflow: Decimal;
    netFlow: Decimal;
    calculatedClosing: Decimal;
    expectedClosing?: Decimal;
    discrepancy?: Decimal;
    reconciled: boolean;
  }

  const tolerance = new Decimal(config.tolerance !== undefined ? config.tolerance : '0.01');
  let overallDiscrepancyTotal = new Decimal(0);
  let allReconciled = true;

  const accountData: AccountAgg[] = sortedAccounts.map((acc) => {
    const openBal = new Decimal(openingBalances[acc] || 0);
    const txs = processedList.filter((t) => t.account === acc);
    let accInflow = new Decimal(0);
    let accOutflow = new Decimal(0);
    for (const t of txs) {
      accInflow = accInflow.plus(t.inflow);
      accOutflow = accOutflow.plus(t.outflow);
    }
    const accNet = accInflow.minus(accOutflow);
    const calcClosing = openBal.plus(accNet);

    let expClosing: Decimal | undefined;
    let discrepancy: Decimal | undefined;
    let isReconciled = true;

    if (expectedClosingBalances[acc] !== undefined) {
      expClosing = new Decimal(expectedClosingBalances[acc]);
      discrepancy = calcClosing.minus(expClosing);
      if (discrepancy.abs().greaterThan(tolerance)) {
        isReconciled = false;
        allReconciled = false;
        overallDiscrepancyTotal = overallDiscrepancyTotal.plus(discrepancy.abs());
      }
    }

    return {
      account: acc,
      openingBalance: openBal,
      inflow: accInflow,
      outflow: accOutflow,
      netFlow: accNet,
      calculatedClosing: calcClosing,
      expectedClosing: expClosing,
      discrepancy,
      reconciled: isReconciled,
    };
  });

  // Calculate grand totals across all transactions
  let grandTotalInflow = new Decimal(0);
  let grandTotalOutflow = new Decimal(0);
  for (const t of processedList) {
    grandTotalInflow = grandTotalInflow.plus(t.inflow);
    grandTotalOutflow = grandTotalOutflow.plus(t.outflow);
  }
  const grandNetCashFlow = grandTotalInflow.minus(grandTotalOutflow);

  // 1. Build Transactions Sheet
  const txSheet = buildTransactionsSheet(processedList, baseCurrency, grandTotalInflow, grandTotalOutflow);

  // 2. Build Monthly Summary Sheet
  const monthlySheet = buildMonthlySummarySheet(monthlyData, baseCurrency, grandTotalInflow, grandTotalOutflow, grandNetCashFlow);

  // 3. Build Account Summary Sheet
  const accountSheet = buildAccountSummarySheet(accountData, baseCurrency);

  // 4. Build Reconciliation Sheet
  const reconSheet = buildReconciliationSheet(
    accountData,
    baseCurrency,
    grandTotalInflow,
    grandTotalOutflow,
    grandNetCashFlow,
    allReconciled,
    overallDiscrepancyTotal
  );

  // 5. Build FX Rates Sheet
  const fxSheet = buildFxRatesSheet(config.fxRates || []);

  const summary: FinanceReportSummary = {
    baseCurrency,
    totalInflow: grandTotalInflow.toFixed(2),
    totalOutflow: grandTotalOutflow.toFixed(2),
    netCashFlow: grandNetCashFlow.toFixed(2),
    transactionCount: processedList.length,
    accountCount: sortedAccounts.length,
    reconciled: allReconciled,
    totalDiscrepancy: overallDiscrepancyTotal.toFixed(2),
  };

  return {
    sheets: [txSheet, monthlySheet, accountSheet, reconSheet, fxSheet],
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
 * Validates and parses raw finance transaction records.
 */
function parseAndValidateFinanceTransactions(
  input: unknown,
  maxRows?: number
): FinanceTransaction[] {
  if (!Array.isArray(input)) {
    throw new Error('Finance report input must be an array of transaction records');
  }

  const limit = maxRows || 100000;
  if (input.length > limit) {
    throw new Error(`Input transactions count (${input.length}) exceeds maximum allowed limit of ${limit}`);
  }

  const result: FinanceTransaction[] = [];

  for (let i = 0; i < input.length; i++) {
    const raw = input[i];
    const rowNum = i + 1;

    if (!raw || typeof raw !== 'object') {
      throw new Error(`Row ${rowNum}: Expected transaction record object`);
    }

    const item = raw as Record<string, unknown>;

    // id
    if (!item.id || String(item.id).trim() === '') {
      throw new Error(`Row ${rowNum}: Transaction "id" is required and cannot be empty`);
    }
    const id = String(item.id).trim();

    // date
    if (!item.date || String(item.date).trim() === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Transaction "date" is required`);
    }
    const dateStr = String(item.date).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Invalid date format "${dateStr}" (expected YYYY-MM-DD)`);
    }

    // account
    if (!item.account || String(item.account).trim() === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Transaction "account" is required`);
    }
    const account = String(item.account).trim();

    // direction
    if (!item.direction) {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Transaction "direction" is required (in/out)`);
    }
    const dirStr = String(item.direction).trim().toLowerCase();
    if (dirStr !== 'in' && dirStr !== 'out') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Direction must be "in" or "out", received "${item.direction}"`);
    }

    // amount
    if (item.amount === undefined || item.amount === null || item.amount === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Transaction "amount" is required`);
    }
    let amountDec: Decimal;
    try {
      amountDec = new Decimal(String(item.amount));
      if (amountDec.lessThan(0)) {
        throw new Error('negative amount');
      }
    } catch {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Invalid non-negative transaction amount "${item.amount}"`);
    }

    // currency
    if (!item.currency || String(item.currency).trim() === '') {
      throw new Error(`Row ${rowNum} (ID: "${id}"): Transaction "currency" is required`);
    }
    const currency = String(item.currency).trim().toUpperCase();

    result.push({
      id,
      date: dateStr,
      account,
      direction: dirStr as 'in' | 'out',
      amount: amountDec.toNumber(),
      currency,
      description: item.description ? String(item.description).trim() : undefined,
      category: item.category ? String(item.category).trim() : undefined,
      isRefund: item.isRefund === true || String(item.isRefund).toLowerCase() === 'true',
    });
  }

  return result;
}

/**
 * Builds the Transactions worksheet.
 */
function buildTransactionsSheet(
  txs: {
    id: string;
    date: string;
    account: string;
    description: string;
    category: string;
    direction: 'in' | 'out';
    originalAmount: Decimal;
    currency: string;
    fxRate: Decimal;
    baseAmount: Decimal;
    inflow: Decimal;
    outflow: Decimal;
    netFlow: Decimal;
  }[],
  baseCurrency: string,
  totalInflow: Decimal,
  totalOutflow: Decimal
): ReportSheet {
  const headers: ReportCell[] = [
    'Transaction ID',
    'Date',
    'Account',
    'Description',
    'Category',
    'Direction',
    'Original Amount',
    'Currency',
    'FX Rate',
    `Base Amount (${baseCurrency})`,
    `Inflow (${baseCurrency})`,
    `Outflow (${baseCurrency})`,
  ];

  const rows: ReportCell[][] = [headers];

  for (let i = 0; i < txs.length; i++) {
    const tx = txs[i];
    const excelRow = i + 2; // Row 1 is header

    const origAmtNum = toSafeNumericCell(tx.originalAmount, 2);
    const fxRateNum = toSafeNumericCell(tx.fxRate, 4);
    const baseAmtNum = toSafeNumericCell(tx.baseAmount, 2);
    const inflowNum = toSafeNumericCell(tx.inflow, 2);
    const outflowNum = toSafeNumericCell(tx.outflow, 2);

    rows.push([
      tx.id,
      tx.date,
      tx.account,
      tx.description,
      tx.category,
      tx.direction,
      origAmtNum,
      tx.currency,
      fxRateNum,
      { formula: `G${excelRow}*I${excelRow}`, result: baseAmtNum },
      { formula: `IF(F${excelRow}="in",J${excelRow},0)`, result: inflowNum },
      { formula: `IF(F${excelRow}="out",J${excelRow},0)`, result: outflowNum },
    ]);
  }

  // Summary row at the bottom if transactions exist
  if (txs.length > 0) {
    const lastDataRow = txs.length + 1;
    rows.push([
      'Total',
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      null,
      { formula: `SUM(K2:K${lastDataRow})`, result: toSafeNumericCell(totalInflow, 2) },
      { formula: `SUM(L2:L${lastDataRow})`, result: toSafeNumericCell(totalOutflow, 2) },
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 7, format: '#,##0.00' },
    { column: 9, format: '#,##0.0000' },
    { column: 10, format: '#,##0.00' },
    { column: 11, format: '#,##0.00' },
    { column: 12, format: '#,##0.00' },
  ];

  return {
    name: 'Transactions',
    rows,
    freezeRows: 1,
    columnWidths: [18, 12, 16, 28, 16, 12, 16, 10, 10, 20, 20, 20],
    numberFormats,
  };
}

/**
 * Builds Monthly Summary worksheet.
 */
function buildMonthlySummarySheet(
  months: { month: string; inflow: Decimal; outflow: Decimal; netFlow: Decimal }[],
  baseCurrency: string,
  grandInflow: Decimal,
  grandOutflow: Decimal,
  grandNet: Decimal
): ReportSheet {
  const headers: ReportCell[] = [
    'Month',
    `Total Inflow (${baseCurrency})`,
    `Total Outflow (${baseCurrency})`,
    `Net Cash Flow (${baseCurrency})`,
  ];

  const rows: ReportCell[][] = [headers];

  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    const excelRow = i + 2;

    rows.push([
      m.month,
      toSafeNumericCell(m.inflow, 2),
      toSafeNumericCell(m.outflow, 2),
      { formula: `B${excelRow}-C${excelRow}`, result: toSafeNumericCell(m.netFlow, 2) },
    ]);
  }

  if (months.length > 0) {
    const lastRow = months.length + 1;
    rows.push([
      'Total',
      { formula: `SUM(B2:B${lastRow})`, result: toSafeNumericCell(grandInflow, 2) },
      { formula: `SUM(C2:C${lastRow})`, result: toSafeNumericCell(grandOutflow, 2) },
      { formula: `SUM(D2:D${lastRow})`, result: toSafeNumericCell(grandNet, 2) },
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '#,##0.00' },
    { column: 3, format: '#,##0.00' },
    { column: 4, format: '#,##0.00' },
  ];

  return {
    name: 'Monthly Summary',
    rows,
    freezeRows: 1,
    columnWidths: [14, 24, 24, 24],
    numberFormats,
  };
}

/**
 * Builds Account Summary worksheet.
 */
function buildAccountSummarySheet(
  accounts: {
    account: string;
    openingBalance: Decimal;
    inflow: Decimal;
    outflow: Decimal;
    netFlow: Decimal;
    calculatedClosing: Decimal;
    expectedClosing?: Decimal;
    discrepancy?: Decimal;
    reconciled: boolean;
  }[],
  baseCurrency: string
): ReportSheet {
  const headers: ReportCell[] = [
    'Account',
    `Opening Balance (${baseCurrency})`,
    `Total Inflow (${baseCurrency})`,
    `Total Outflow (${baseCurrency})`,
    `Net Flow (${baseCurrency})`,
    `Calculated Closing (${baseCurrency})`,
    `Expected Closing (${baseCurrency})`,
    `Discrepancy (${baseCurrency})`,
    'Reconciliation Status',
  ];

  const rows: ReportCell[][] = [headers];

  for (let i = 0; i < accounts.length; i++) {
    const a = accounts[i];
    const excelRow = i + 2;

    const expVal = a.expectedClosing !== undefined ? toSafeNumericCell(a.expectedClosing, 2) : null;
    const discVal = a.discrepancy !== undefined ? toSafeNumericCell(a.discrepancy, 2) : null;

    rows.push([
      a.account,
      toSafeNumericCell(a.openingBalance, 2),
      toSafeNumericCell(a.inflow, 2),
      toSafeNumericCell(a.outflow, 2),
      { formula: `C${excelRow}-D${excelRow}`, result: toSafeNumericCell(a.netFlow, 2) },
      { formula: `B${excelRow}+E${excelRow}`, result: toSafeNumericCell(a.calculatedClosing, 2) },
      expVal,
      expVal !== null ? { formula: `F${excelRow}-G${excelRow}`, result: discVal } : null,
      a.reconciled ? 'Reconciled' : 'Discrepancy Detected',
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '#,##0.00' },
    { column: 3, format: '#,##0.00' },
    { column: 4, format: '#,##0.00' },
    { column: 5, format: '#,##0.00' },
    { column: 6, format: '#,##0.00' },
    { column: 7, format: '#,##0.00' },
    { column: 8, format: '#,##0.00' },
  ];

  return {
    name: 'Account Summary',
    rows,
    freezeRows: 1,
    columnWidths: [20, 22, 20, 20, 20, 24, 24, 20, 22],
    numberFormats,
  };
}

/**
 * Builds Reconciliation audit worksheet.
 */
function buildReconciliationSheet(
  accounts: {
    account: string;
    openingBalance: Decimal;
    inflow: Decimal;
    outflow: Decimal;
    netFlow: Decimal;
    calculatedClosing: Decimal;
    expectedClosing?: Decimal;
    discrepancy?: Decimal;
    reconciled: boolean;
  }[],
  baseCurrency: string,
  totalInflow: Decimal,
  totalOutflow: Decimal,
  netCashFlow: Decimal,
  allReconciled: boolean,
  totalDiscrepancy: Decimal
): ReportSheet {
  let totalOpening = new Decimal(0);
  let totalCalculatedClosing = new Decimal(0);
  for (const a of accounts) {
    totalOpening = totalOpening.plus(a.openingBalance);
    totalCalculatedClosing = totalCalculatedClosing.plus(a.calculatedClosing);
  }

  const rows: ReportCell[][] = [
    ['Audit Reconciliation Summary', null],
    ['Base Currency', baseCurrency],
    ['Total Accounts Count', accounts.length],
    ['Total Opening Balance', toSafeNumericCell(totalOpening, 2)],
    ['Total Inflow', toSafeNumericCell(totalInflow, 2)],
    ['Total Outflow', toSafeNumericCell(totalOutflow, 2)],
    ['Net Cash Flow', toSafeNumericCell(netCashFlow, 2)],
    ['Total Calculated Closing Balance', toSafeNumericCell(totalCalculatedClosing, 2)],
    ['Audit Mathematical Integrity Check', { formula: 'B4+B7=B8', result: totalOpening.plus(netCashFlow).equals(totalCalculatedClosing) }],
    ['Overall Status', allReconciled ? 'PASSED (Reconciled)' : 'FAILED (Discrepancies Detected)'],
    ['Total Discrepancy Amount', toSafeNumericCell(totalDiscrepancy, 2)],
  ];

  const numberFormats: ReportNumberFormat[] = [
    { column: 2, format: '#,##0.00' },
  ];

  return {
    name: 'Reconciliation',
    rows,
    freezeRows: 1,
    columnWidths: [36, 28],
    numberFormats,
  };
}

/**
 * Builds FX Rates reference worksheet.
 */
function buildFxRatesSheet(rates: FxRate[]): ReportSheet {
  const headers: ReportCell[] = [
    'Date',
    'From Currency',
    'To Currency',
    'Exchange Rate',
    'Source',
  ];

  const rows: ReportCell[][] = [headers];

  for (const r of rates) {
    rows.push([
      r.date,
      r.from.toUpperCase(),
      r.to.toUpperCase(),
      new Decimal(r.rate).toNumber(),
      r.source || 'Manual/Preset',
    ]);
  }

  const numberFormats: ReportNumberFormat[] = [
    { column: 4, format: '#,##0.0000' },
  ];

  return {
    name: 'FX Rates',
    rows,
    freezeRows: 1,
    columnWidths: [14, 16, 16, 18, 22],
    numberFormats,
  };
}

/**
 * Builds an FX lookup index by currency pair and date.
 */
function buildFxRateIndex(rates: FxRate[]): Map<string, { date: string; rate: number }[]> {
  const map = new Map<string, { date: string; rate: number }[]>();

  for (const r of rates) {
    const pairKey = `${r.from.toUpperCase()}->${r.to.toUpperCase()}`;
    let list = map.get(pairKey);
    if (!list) {
      list = [];
      map.set(pairKey, list);
    }
    list.push({
      date: r.date,
      rate: new Decimal(r.rate).toNumber(),
    });
  }

  // Sort each list descending by date
  for (const list of map.values()) {
    list.sort((a, b) => b.date.localeCompare(a.date));
  }

  return map;
}

/**
 * Finds the exchange rate for a pair on a given date (exact or latest preceding).
 */
function findFxRate(
  fxMap: Map<string, { date: string; rate: number }[]>,
  date: string,
  from: string,
  to: string
): number | null {
  const directKey = `${from}->${to}`;
  const directList = fxMap.get(directKey);

  if (directList && directList.length > 0) {
    // Exact match
    const exact = directList.find((r) => r.date === date);
    if (exact) return exact.rate;

    // Latest preceding date
    const preceding = directList.find((r) => r.date <= date);
    if (preceding) return preceding.rate;
  }

  // Check inverse pair (to -> from)
  const inverseKey = `${to}->${from}`;
  const inverseList = fxMap.get(inverseKey);
  if (inverseList && inverseList.length > 0) {
    const exact = inverseList.find((r) => r.date === date);
    if (exact && exact.rate !== 0) {
      return new Decimal(1).dividedBy(new Decimal(exact.rate)).toNumber();
    }
    const preceding = inverseList.find((r) => r.date <= date);
    if (preceding && preceding.rate !== 0) {
      return new Decimal(1).dividedBy(new Decimal(preceding.rate)).toNumber();
    }
  }

  return null;
}

/**
 * Normalizes openingBalances from Record or Array format.
 */
function normalizeOpeningBalances(
  balances?: Record<string, string | number> | FinanceOpeningBalance[]
): Record<string, number> {
  const result: Record<string, number> = Object.create(null);

  if (!balances) return result;

  if (Array.isArray(balances)) {
    for (const b of balances) {
      if (b && b.account) {
        result[b.account] = new Decimal(b.balance || 0).toNumber();
      }
    }
  } else if (typeof balances === 'object') {
    for (const [acc, val] of Object.entries(balances)) {
      result[acc] = new Decimal(val || 0).toNumber();
    }
  }

  return result;
}
