import { expect } from 'chai';
import * as fs from 'fs';
import * as path from 'path';
import { FakeSheets } from './fake-sheets';
import GoogleSheet, { GoogleSheetCli } from '../src/lib/google-sheet';
import { XlsxWorkbook } from '../src/lib/xlsx';
import {
  buildReport,
  buildFinanceReport,
  buildManpowerReport,
  buildTableReport,
  parseInput,
  computeSourceHash,
  ReportDocument,
} from '../src/lib/report';
import { runReport } from '../src/lib/report/runner';

const TMP_DIR = path.join(__dirname, '..', 'tmp', 'integration-tests');

describe('Report & Spreadsheet Integration Suite (Wave B)', () => {
  let fakeSheets: FakeSheets;
  const SPREADSHEET_ID = 'integration-sheet-1';

  before(() => {
    if (!fs.existsSync(TMP_DIR)) {
      fs.mkdirSync(TMP_DIR, { recursive: true });
    }
  });

  after(() => {
    if (fs.existsSync(TMP_DIR)) {
      fs.rmSync(TMP_DIR, { recursive: true, force: true });
    }
  });

  beforeEach(() => {
    fakeSheets = new FakeSheets();
    fakeSheets.install();
    fakeSheets.addSpreadsheet(SPREADSHEET_ID, 'Report Automation Test', [
      { title: 'Sheet1', rowCount: 100, columnCount: 26 },
    ]);
  });

  afterEach(() => {
    fakeSheets.uninstall();
    fakeSheets.reset();
  });

  describe('Google Sheet Backend Integration', () => {
    it('executes 12 bounded ranges in a single batch read request without repeated metadata reads', async () => {
      const sheet = new GoogleSheet(SPREADSHEET_ID);
      await sheet.authorize(fakeSheets.credentials);

      // Populate test cells in fake
      for (let i = 1; i <= 12; i++) {
        fakeSheets.setCells(SPREADSHEET_ID, 'Sheet1', `A${i}:B${i}`, [[`Key${i}`, i * 10]]);
      }

      const ranges = Array.from({ length: 12 }, (_, idx) => `Sheet1!A${idx + 1}:B${idx + 1}`);
      const results = await sheet.getDataBatch(ranges, { valueRenderOption: GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE }, SPREADSHEET_ID);

      expect(results).to.have.lengthOf(12);
      expect(results[0].values).to.deep.equal([['Key1', 10]]);
      expect(results[11].values).to.deep.equal([['Key12', 120]]);

      // Verify that values.batchGet was called and no metadata roundtrips per range
      const batchGetCalls = fakeSheets.requests.filter((r) => r.url.includes('/values:batchGet'));
      expect(batchGetCalls).to.have.lengthOf(1);
    });

    it('applies mixed typed cells (strings, numbers, booleans, formulas, literal "=" text) cleanly', async () => {
      const sheet = new GoogleSheet(SPREADSHEET_ID);
      await sheet.authorize(fakeSheets.credentials);

      const doc: ReportDocument = {
        sheets: [
          {
            name: 'TypedSheet',
            rows: [
              ['Header String', 100, true, { formula: 'SUM(B2:B5)', result: 400 }, "='literal text with equals"],
            ],
            freezeRows: 1,
            columnWidths: [20, 15, 10, 15, 25],
            numberFormats: [{ column: 2, format: '#,##0.00' }],
          },
        ],
        provenance: {
          templateId: 'test-typed',
          templateVersion: 1,
          sourceHash: 'hash-123',
        },
      };

      const receipt = await sheet.applyReport(doc, {}, SPREADSHEET_ID);
      expect(receipt.sheetsApplied).to.have.lengthOf(1);
      expect(receipt.sheetsApplied[0].name).to.equal('TypedSheet');

      // Verify final stored cells in fake
      expect(fakeSheets.cell(SPREADSHEET_ID, 'TypedSheet', 'A1')).to.equal('Header String');
      expect(fakeSheets.cell(SPREADSHEET_ID, 'TypedSheet', 'B1')).to.equal('100');
      expect(fakeSheets.cell(SPREADSHEET_ID, 'TypedSheet', 'C1')).to.equal('true');
      expect(fakeSheets.cell(SPREADSHEET_ID, 'TypedSheet', 'D1')).to.equal('=SUM(B2:B5)');
      expect(fakeSheets.cell(SPREADSHEET_ID, 'TypedSheet', 'E1')).to.equal("='literal text with equals");
    });

    it('protects existing formulas against un-opted overwrite and allows dryRun preview', async () => {
      const sheet = new GoogleSheet(SPREADSHEET_ID);
      await sheet.authorize(fakeSheets.credentials);

      // Seed an existing formula cell
      fakeSheets.setCells(SPREADSHEET_ID, 'Sheet1', 'B2', [['=AVERAGE(A1:A10)']]);

      const doc: ReportDocument = {
        sheets: [
          {
            name: 'Sheet1',
            startCell: 'A1',
            rows: [
              ['Header1', 'Header2'],
              ['Data1', 999], // Target B2 has an existing formula!
            ],
          },
        ],
        provenance: {
          templateId: 'test-prot',
          templateVersion: 1,
          sourceHash: 'hash-prot',
        },
      };

      // Dry run should report preview without throwing or mutating
      const previewReceipt = await sheet.applyReport(doc, { dryRun: true }, SPREADSHEET_ID);
      expect(previewReceipt.dryRun).to.be.true;
      expect(fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B2')).to.equal('=AVERAGE(A1:A10)');

      // Sheet1 is pre-populated and unowned by this report: the run must be
      // rejected, judged by observable behavior (rejection plus retained prior
      // formula), never by error wording.
      let rejected = false;
      try {
        await sheet.applyReport(doc, {}, SPREADSHEET_ID);
      } catch {
        rejected = true;
      }
      expect(rejected, 'populated unowned worksheet must require overwrite=true').to.be.true;
      expect(fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B2')).to.equal('=AVERAGE(A1:A10)');

      // overwrite:true consents to touching the unowned worksheet but does not
      // opt into formula replacement: the formula guard must still reject.
      rejected = false;
      try {
        await sheet.applyReport(doc, { overwrite: true }, SPREADSHEET_ID);
      } catch {
        rejected = true;
      }
      expect(rejected, 'existing formulas must still require overwriteFormulas=true').to.be.true;
      expect(fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B2')).to.equal('=AVERAGE(A1:A10)');

      // Intended replacement: granting both overwrite and overwriteFormulas proceeds.
      const overwriteReceipt = await sheet.applyReport(
        doc,
        { overwrite: true, overwriteFormulas: true },
        SPREADSHEET_ID
      );
      expect(overwriteReceipt.dryRun).to.be.false;
      expect(fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B2')).to.equal('999');
    });

    it('performs concurrent native table appends and assigns distinct non-colliding rows', async () => {
      const sheet = new GoogleSheet(SPREADSHEET_ID);
      await sheet.authorize(fakeSheets.credentials);

      // Initial table header
      fakeSheets.setCells(SPREADSHEET_ID, 'Sheet1', 'A1:C1', [['ID', 'Name', 'Score']]);

      const appends = [
        sheet.appendTableData([['1', 'Alice', '90']], { worksheetTitle: 'Sheet1' }, SPREADSHEET_ID),
        sheet.appendTableData([['2', 'Bob', '85']], { worksheetTitle: 'Sheet1' }, SPREADSHEET_ID),
        sheet.appendTableData([['3', 'Charlie', '95']], { worksheetTitle: 'Sheet1' }, SPREADSHEET_ID),
      ];

      const results = await Promise.all(appends);
      expect(results).to.have.lengthOf(3);

      // Verify all rows are distinctly written across rows 2, 3, 4
      const r2 = [fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'A2'), fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B2')];
      const r3 = [fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'A3'), fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B3')];
      const r4 = [fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'A4'), fakeSheets.cell(SPREADSHEET_ID, 'Sheet1', 'B4')];

      const ids = [r2[0], r3[0], r4[0]].sort();
      expect(ids).to.deep.equal(['1', '2', '3']);
    });
  });

  describe('XLSX Backend Integration & Formula Safety', () => {
    it('blocks replacing an existing formula with a DIFFERENT formula without overwriteFormulas', () => {
      const workbook = XlsxWorkbook.create();
      const ws = workbook.workbook.addWorksheet('Formulas');
      ws.getCell('B2').value = { formula: 'SUM(A1:A10)', result: 50 };

      const doc: ReportDocument = {
        sheets: [
          {
            name: 'Formulas',
            startCell: 'A1',
            rows: [
              ['Col A', 'Col B'],
              ['Val1', { formula: 'AVERAGE(A1:A10)', result: 25 }], // Different formula!
            ],
          },
        ],
        provenance: { templateId: 't-form', templateVersion: 1, sourceHash: 'h-1' },
      };

      // Preview should detect conflict
      const diff = workbook.preview(doc, { overwriteFormulas: false });
      expect(diff.totalConflicts).to.equal(1);
      expect(diff.conflicts).to.have.lengthOf(1);
      expect(diff.conflicts[0].address).to.equal('B2');

      // apply() without overwriteFormulas should throw
      expect(() => workbook.apply(doc, { overwriteFormulas: false })).to.throw(/formula/i);

      // apply() with overwriteFormulas: true succeeds
      const result = workbook.apply(doc, { overwriteFormulas: true });
      expect(result.success).to.be.true;

      const readBack = workbook.read('Formulas!B2:B2', { mode: 'formula' });
      expect(readBack.values[0][0]).to.deep.equal({ formula: 'AVERAGE(A1:A10)', result: 25 });
    });
    it('shorter rerun with clearManagedRange clears trailing owned rows and columns without touching unrelated template cells', () => {
      const workbook = XlsxWorkbook.create();
      const ws = workbook.workbook.addWorksheet('Data');

      // Unrelated template cells nearby
      ws.getCell('E1').value = 'Template Notes (DO NOT TOUCH)';
      ws.getCell('A10').value = 'Footer / Signatures (DO NOT TOUCH)';

      // First run: 5 rows x 3 cols (A1..C5)
      const doc1: ReportDocument = {
        sheets: [
          {
            name: 'Data',
            startCell: 'A1',
            clearManagedRange: true,
            rows: [
              ['H1', 'H2', 'H3'],
              ['1', '2', '3'],
              ['4', '5', '6'],
              ['7', '8', '9'],
              ['10', '11', '12'],
            ],
          },
        ],
        provenance: { templateId: 't-shorter', templateVersion: 1, sourceHash: 'h-run1' },
      };
      workbook.apply(doc1);

      // Verify run 1 cells
      expect(ws.getCell('C5').value).to.equal('12');
      expect(ws.getCell('E1').value).to.equal('Template Notes (DO NOT TOUCH)');
      expect(ws.getCell('A10').value).to.equal('Footer / Signatures (DO NOT TOUCH)');

      // Second run: shorter: 2 rows x 2 cols (A1..B2)
      const doc2: ReportDocument = {
        sheets: [
          {
            name: 'Data',
            startCell: 'A1',
            clearManagedRange: true,
            rows: [
              ['H1_New', 'H2_New'],
              ['100', '200'],
            ],
          },
        ],
        provenance: { templateId: 't-shorter', templateVersion: 1, sourceHash: 'h-run2' },
      };
      workbook.apply(doc2);

      // Verify owned leftovers in rows 3..5 and col C are cleared (null or empty)
      expect(ws.getCell('A1').value).to.equal('H1_New');
      expect(ws.getCell('B2').value).to.equal('200');
      expect(ws.getCell('C3').value).to.be.oneOf([null, undefined, '']);
      expect(ws.getCell('C5').value).to.be.oneOf([null, undefined, '']);
      expect(ws.getCell('A4').value).to.be.oneOf([null, undefined, '']);

      // Verify unrelated template cells remain intact
      expect(ws.getCell('E1').value).to.equal('Template Notes (DO NOT TOUCH)');
      expect(ws.getCell('A10').value).to.equal('Footer / Signatures (DO NOT TOUCH)');
    });

    it('reports formula freshness as unknown when cached result is missing and never assumes recalculation', () => {
      const workbook = XlsxWorkbook.create();
      const ws = workbook.workbook.addWorksheet('Freshness');
      ws.getCell('A1').value = 10;
      ws.getCell('A2').value = 20;
      // Formula with cached result
      ws.getCell('A3').value = { formula: 'A1+A2', result: 30 };
      // Formula without cached result (raw formula string)
      ws.getCell('A4').value = { formula: 'A1*2' };

      const readResult = workbook.read('Freshness!A1:A4', { mode: 'formula' });
      expect(readResult.formulaCount).to.equal(2);
      expect(readResult.freshnessSummary.cachedCount).to.equal(1);
      expect(readResult.freshnessSummary.unknownCount).to.equal(1);
      expect(readResult.freshnessSummary.notFormulaCount).to.equal(2);

      const inspection = workbook.inspect();
      expect(inspection.capabilities.canRecalculateFormulas).to.be.false;
    });
  });

  describe('Business Report Presets & Calculations', () => {
    it('calculates Finance report with multi-currency FX, refunds as signed flows, and balance reconciliation', () => {
      const template = {
        id: 'finance-monthly',
        version: 1,
        kind: 'finance' as const,
        config: {
          baseCurrency: 'VND',
          fxRates: [
            { date: '2026-09-01', from: 'USD', to: 'VND', rate: 25000 },
          ],
          openingBalances: {
            acc_main: 100000000, // 100m VND
          },
          expectedClosingBalances: {
            acc_main: 180000000, // 180m VND
          },
        },
      };

      const transactions = [
        {
          id: 'tx_01',
          date: '2026-09-01',
          account: 'acc_main',
          direction: 'in' as const,
          category: 'Consulting',
          amount: 50000000,
          currency: 'VND',
          description: 'Software Architecture Consulting',
        },
        {
          id: 'tx_02',
          date: '2026-09-05',
          account: 'acc_main',
          direction: 'in' as const,
          category: 'Overseas Contract',
          amount: 2000, // USD -> 50m VND
          currency: 'USD',
          description: 'Cloud Migration USD',
        },
        {
          id: 'tx_03',
          date: '2026-09-10',
          account: 'acc_main',
          direction: 'out' as const,
          category: 'Infrastructure',
          amount: 15000000,
          currency: 'VND',
          description: 'Server Hosting',
        },
        {
          id: 'tx_04',
          date: '2026-09-15',
          account: 'acc_main',
          direction: 'in' as const,
          category: 'Consulting',
          amount: 5000000,
          currency: 'VND',
          isRefund: true,
          description: 'Partial refund on consulting',
        },
      ];

      const doc = buildFinanceReport(template, transactions, 'test-finance-hash');
      expect(doc.sheets).to.have.length.greaterThan(0);

      // Verify provenance
      expect(doc.provenance.templateId).to.equal('finance-monthly');
      expect(doc.provenance.sourceHash).to.equal('test-finance-hash');

      // Check summary values (refunds are signed flows: an 'in' refund reduces inflow):
      // Inflow: 50m VND + (2000 USD * 25000 = 50m VND) - 5m refund = 95m VND
      // Outflow: 15m expense
      // Net flow: 95m - 15m = 80m VND
      // Closing balance: 100m (opening) + 80m (net) = 180m VND (matches expectedClosingBalances)
      const summary = doc.provenance.summary as {
        totalInflow: string;
        totalOutflow: string;
        netCashFlow: string;
        reconciled: boolean;
      };
      expect(summary.totalInflow).to.equal('95000000.00');
      expect(summary.totalOutflow).to.equal('15000000.00');
      expect(summary.netCashFlow).to.equal('80000000.00');
      expect(summary.reconciled).to.be.true;

      // Verify materialization to XLSX
      const workbook = XlsxWorkbook.create();
      workbook.apply(doc);
      const reconWs = workbook.workbook.getWorksheet('Reconciliation');
      expect(reconWs).to.not.be.undefined;
    });

    it('rejects duplicate transaction IDs in finance report before computation', () => {
      const template = {
        id: 'finance-dup',
        version: 1,
        kind: 'finance' as const,
        config: {
          baseCurrency: 'VND',
          openingBalances: {
            acc_main: 0,
          },
        },
      };

      const dupTx = [
        { id: 'tx_dup', date: '2026-09-01', account: 'acc_main', direction: 'in' as const, amount: 100, currency: 'VND' },
        { id: 'tx_dup', date: '2026-09-02', account: 'acc_main', direction: 'in' as const, amount: 200, currency: 'VND' },
      ];

      expect(() => buildFinanceReport(template, dupTx, 'hash')).to.throw(/duplicate/i);
    });

    it('calculates Manpower report with exact 14.4 days effort and 36m cost equation and multi-scenarios', () => {
      const template = {
        id: 'manpower-standard',
        version: 1,
        kind: 'manpower' as const,
        config: {
          currency: 'VND',
          effortUnit: 'person-days' as const,
          roleRates: {
            senior_eng: 2500000, // 2.5m / day
            qa_lead: 1800000,
          },
          globalContingencyPercent: 20, // 20% -> 1.2 multiplier
          scenarios: [
            { name: 'Optimistic', contingencyPercent: 10 },
            { name: 'Base', contingencyPercent: 20, isBase: true },
            { name: 'Pessimistic', contingencyPercent: 40 },
          ],
        },
      };

      const tasks = [
        {
          id: 'TASK-01',
          module: 'Core',
          role: 'senior_eng',
          quantity: 2,
          unitEffort: 5,
          complexity: 1.2,
          description: 'Batch Engine & Pipeline',
        },
      ];

      // Calculation:
      // Quantity(2) * UnitEffort(5) * Complexity(1.2) = 12 base person-days
      // BaseEffort(12) * Contingency(1 + 0.2 = 1.2) = 14.4 total person-days
      // Total Cost = 14.4 days * 2,500,000 VND/day = 36,000,000 VND (36m)
      const doc = buildManpowerReport(template, tasks, 'test-manpower-hash');

      expect(doc.sheets.length).to.be.greaterThan(1);

      // Verify task details sheet
      const taskSheet = doc.sheets.find((s) => s.name === 'Task Details');
      expect(taskSheet).to.not.be.undefined;

      // Find the row containing TASK-01 and verify numbers
      const taskRow = taskSheet!.rows.find((r) => r.some((c) => c === 'TASK-01'));
      expect(taskRow).to.not.be.undefined;

      // Check that 14.4 and 36000000 are present in the calculated cells or formula results
      const hasEffort14_4 = taskRow!.some(
        (c) => c === 14.4 || (typeof c === 'object' && c !== null && 'result' in c && (c as { result: number }).result === 14.4)
      );
      const hasCost36m = taskRow!.some(
        (c) => c === 36000000 || (typeof c === 'object' && c !== null && 'result' in c && (c as { result: number }).result === 36000000)
      );

      expect(hasEffort14_4, 'Must compute 14.4 person-days').to.be.true;
      expect(hasCost36m, 'Must compute 36m VND cost').to.be.true;

      // Apply to actual XLSX and verify final cells directly
      const workbook = XlsxWorkbook.create();
      workbook.apply(doc);
      const xlsxTaskSheet = workbook.workbook.getWorksheet('Task Details')!;
      expect(xlsxTaskSheet).to.not.be.undefined;

      // Row 2 is TASK-01: Col J (10) is Total Effort, Col L (12) is Total Cost
      const cellTotalEffort = xlsxTaskSheet.getCell('J2').value;
      const cellTotalCost = xlsxTaskSheet.getCell('L2').value;

      const effortVal = typeof cellTotalEffort === 'object' && cellTotalEffort !== null && 'result' in cellTotalEffort
        ? (cellTotalEffort as { result: number }).result
        : cellTotalEffort;
      const costVal = typeof cellTotalCost === 'object' && cellTotalCost !== null && 'result' in cellTotalCost
        ? (cellTotalCost as { result: number }).result
        : cellTotalCost;

      expect(effortVal).to.equal(14.4);
      expect(costVal).to.equal(36000000);
    });
  });

  describe('Dual-Backend Parity & Idempotence', () => {
    it('produces identical summary values on Google Sheets and XLSX backends from the same input and template', async () => {
      const template = {
        id: 'sales-summary',
        version: 1,
        kind: 'table' as const,
        table: {
          title: 'Sales Q3',
          columns: [
            { key: 'region', header: 'Region', type: 'string' },
            { key: 'units', header: 'Units', type: 'integer' },
            { key: 'revenue', header: 'Revenue', type: 'decimal' },
          ],
        },
      };

      const inputData = [
        { region: 'North', units: 100, revenue: 50000 },
        { region: 'South', units: 250, revenue: 125000 },
        { region: 'Central', units: 150, revenue: 75000 },
      ];

      const doc = buildReport(template, inputData);

      // 1. Apply to Google Sheets
      const sheet = new GoogleSheet(SPREADSHEET_ID);
      await sheet.authorize(fakeSheets.credentials);
      await sheet.applyReport(doc, {}, SPREADSHEET_ID);

      const googleValRegion = fakeSheets.cell(SPREADSHEET_ID, doc.sheets[0].name, 'A2');
      const googleValUnits = fakeSheets.cell(SPREADSHEET_ID, doc.sheets[0].name, 'B2');
      const googleValRev = fakeSheets.cell(SPREADSHEET_ID, doc.sheets[0].name, 'C2');

      // 2. Apply to XLSX
      const workbook = XlsxWorkbook.create();
      workbook.apply(doc);
      const xlsxSheet = workbook.workbook.getWorksheet(doc.sheets[0].name);

      expect(xlsxSheet).to.not.be.undefined;
      const xlsxValRegion = xlsxSheet!.getCell('A2').value;
      const xlsxValUnits = xlsxSheet!.getCell('B2').value;
      const xlsxValRev = xlsxSheet!.getCell('C2').value;

      // Assert parity across backends
      expect(String(googleValRegion)).to.equal(String(xlsxValRegion));
      expect(Number(googleValUnits)).to.equal(Number(xlsxValUnits));
      expect(Number(googleValRev)).to.equal(Number(xlsxValRev));
    });
  });

  describe('Report CLI Runner End-to-End', () => {
    it('executes report:run locally without any Google credentials, outputting to a new XLSX file', async () => {
      const templatePath = path.join(TMP_DIR, 'local-template.json');
      const dataPath = path.join(TMP_DIR, 'local-data.csv');
      const outputPath = path.join(TMP_DIR, 'output-report.xlsx');

      fs.writeFileSync(
        templatePath,
        JSON.stringify({
          id: 'cli-test',
          version: 1,
          kind: 'table',
          title: 'CLI Test',
          schema: {
            fields: [
              { name: 'id', type: 'string', required: true },
              { name: 'val', type: 'decimal', required: true },
            ],
          },
          outputs: [
            {
              name: 'CLI Test',
              columns: [
                { header: 'ID', field: 'id', width: 14 },
                { header: 'Value', field: 'val', width: 16, format: '#,##0.00' },
              ],
              freezeRows: 1,
            },
          ],
        }),
        'utf8'
      );

      fs.writeFileSync(dataPath, 'id,val\nITEM-001,150.75\nITEM-002,300.25\n', 'utf8');

      // Ensure no Google env vars
      const receipt = await runReport({
        template: templatePath,
        source: {
          input: {
            file: dataPath,
            format: 'csv',
          },
        },
        target: {
          outputFile: {
            filePath: outputPath,
            overwrite: true,
          },
        },
        dryRun: false,
      });

      expect(receipt.targetKind).to.equal('outputFile');
      expect(receipt.sheetsGenerated).to.equal(1);
      expect(fs.existsSync(outputPath)).to.be.true;

      // Inspect written workbook
      const loaded = await XlsxWorkbook.load(outputPath);
      const readResult = loaded.read('CLI Test!A2:B3', { mode: 'unformatted' });
      expect(readResult.values).to.deep.equal([
        ['ITEM-001', 150.75],
        ['ITEM-002', 300.25],
      ]);
    });
    it('fails before output write when input data is malformed or invalid', async () => {
      const templatePath = path.join(TMP_DIR, 'invalid-test-template.json');
      const dataPath = path.join(TMP_DIR, 'bad-data.json');
      const outputPath = path.join(TMP_DIR, 'should-not-exist.xlsx');

      fs.writeFileSync(
        templatePath,
        JSON.stringify({
          id: 'invalid-test',
          version: 1,
          kind: 'table',
          title: 'CLI Test',
          schema: {
            fields: [
              { name: 'id', type: 'string', required: true },
              { name: 'count', type: 'integer', required: true },
            ],
          },
          outputs: [
            {
              name: 'CLI Test',
              columns: [
                { header: 'ID', field: 'id', width: 14 },
                { header: 'Count', field: 'count', width: 14, format: '#,##0' },
              ],
              freezeRows: 1,
            },
          ],
        }),
        'utf8'
      );

      // Invalid data: the required integer field "count" carries a non-numeric
      // string, so the declared schema must reject it during document build,
      // before any output file is written. The sibling "id" field is valid,
      // pinning the rejection to the numeric type violation alone.
      fs.writeFileSync(
        dataPath,
        JSON.stringify([{ id: 'ITEM-001', count: 'not-a-number' }]),
        'utf8'
      );

      let failed = false;
      try {
        await runReport({
          template: templatePath,
          source: {
            input: {
              file: dataPath,
              format: 'json',
            },
          },
          target: {
            outputFile: {
              filePath: outputPath,
              overwrite: true,
            },
          },
        });
      } catch {
        failed = true;
      }

      expect(failed, 'Should fail validation before write').to.be.true;
      expect(fs.existsSync(outputPath), 'Output file must not be created on validation failure').to.be.false;
    });
  });
});
