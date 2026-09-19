import { expect } from 'chai';
import * as path from 'path';
import {
  buildFinanceReport,
  buildManpowerReport,
  buildReport,
  buildTableReport,
  computeSourceHash,
  evaluateCondition,
  evaluateExpression,
  executeTransforms,
  normalizeTable,
  parseInput,
  readInput,
  ReportDocument,
  ReportSchemaError,
  ReportTemplateV1,
  TableSchema,
} from '../src/lib/report';

describe('ReportCore', () => {
  describe('Input Reading & Parsing (T-04, AC-02)', () => {
    it('enforces mutual exclusivity between data and file options', async () => {
      let threw = false;
      try {
        await readInput({ data: '[]', file: 'test.csv' });
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('Mutually exclusive');
      }
      expect(threw).to.be.true;
    });

    it('rejects call when neither data nor file is provided', async () => {
      let threw = false;
      try {
        await readInput({});
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('Must provide either "data" or "file"');
      }
      expect(threw).to.be.true;
    });

    it('rejects input exceeding maximum byte limits', async () => {
      const largeData = 'a'.repeat(200);
      let threw = false;
      try {
        await readInput({ data: largeData, maxBytes: 100 });
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('exceeds limit');
      }
      expect(threw).to.be.true;
    });

    it('parses JSON datasets preserving types and sanitizing prototype pollution', () => {
      const rawJson = `[
        {"id": 1, "name": "Item A", "active": true, "__proto__": {"admin": true}},
        {"id": 2, "name": "Item B", "active": false}
      ]`;
      const result = parseInput(rawJson, 'json') as Record<string, unknown>[];
      expect(result).to.have.lengthOf(2);
      expect(result[0].id).to.equal(1);
      expect(result[0].active).to.be.true;
      expect(({} as Record<string, unknown>).admin).to.be.undefined;
    });

    it('parses CSV preserving leading zeros in IDs and codes', () => {
      const csvData = `order_id,code,amount\n00123,00045,99.50\n00999,01000,150.00`;
      const result = parseInput(csvData, 'csv') as Record<string, unknown>[];
      expect(result).to.have.lengthOf(2);
      expect(result[0].order_id).to.equal('00123');
      expect(result[0].code).to.equal('00045');
      expect(result[1].order_id).to.equal('00999');
      expect(result[1].code).to.equal('01000');
    });

    it('parses CSV 2D matrix format correctly', () => {
      const csvData = `A,B,C\n1,2,3\n4,5,6`;
      const result = parseInput(csvData, 'csv', { matrix: true }) as (string | null)[][];
      expect(result).to.deep.equal([
        ['A', 'B', 'C'],
        ['1', '2', '3'],
        ['4', '5', '6'],
      ]);
    });

    it('handles quoted fields, escaped commas, newlines, and UTF-8 characters', () => {
      const csvData = `id,name,notes\n1,"Doe, John","Multi-line\nDescription"\n2,"Nguyễn Văn A","Tiếng Việt có dấu"`;
      const result = parseInput(csvData, 'csv') as Record<string, unknown>[];
      expect(result).to.have.lengthOf(2);
      expect(result[0].name).to.equal('Doe, John');
      expect(result[0].notes).to.equal('Multi-line\nDescription');
      expect(result[1].name).to.equal('Nguyễn Văn A');
      expect(result[1].notes).to.equal('Tiếng Việt có dấu');
    });
  });

  describe('Schema Normalization & Type Validation (T-04, AC-02)', () => {
    const schema: TableSchema = {
      fields: [
        { name: 'id', type: 'string', required: true, min: 2 },
        { name: 'amount', type: 'decimal', required: true, decimalPlaces: 2 },
        { name: 'quantity', type: 'integer', required: true, min: 1 },
        { name: 'active', type: 'boolean', required: true },
        { name: 'date', type: 'date', required: true, dateFormat: 'YYYY-MM-DD' },
        { name: 'optional_note', type: 'string', nullable: true, emptyAsNull: true },
      ],
      allowExtraColumns: false,
    };

    it('normalizes valid records matching schema', () => {
      const raw = [
        {
          id: '001',
          amount: '125.456',
          quantity: '5',
          active: 'true',
          date: '2026-02-15',
          optional_note: '  ',
        },
      ];

      const normalized = normalizeTable(raw, schema);
      expect(normalized).to.have.lengthOf(1);
      expect(normalized[0].id).to.equal('001');
      expect(normalized[0].amount).to.equal(125.46);
      expect(normalized[0].quantity).to.equal(5);
      expect(normalized[0].active).to.be.true;
      expect(normalized[0].date).to.equal('2026-02-15');
      expect(normalized[0].optional_note).to.be.null;
    });

    it('rejects invalid decimal strings with clear row/column errors', () => {
      const raw = [
        {
          id: '001',
          amount: 'invalid_num',
          quantity: 5,
          active: true,
          date: '2026-01-01',
        },
      ];

      let error: ReportSchemaError | null = null;
      try {
        normalizeTable(raw, schema);
      } catch (err: unknown) {
        error = err as ReportSchemaError;
      }

      expect(error).to.not.be.null;
      expect(error!.errors).to.have.lengthOf(1);
      expect(error!.errors[0].row).to.equal(1);
      expect(error!.errors[0].field).to.equal('amount');
      expect(error!.errors[0].code).to.equal('DECIMAL_PARSE_ERROR');
    });

    it('handles explicit decimal and thousand separators (e.g. EU locale 1.234,56)', () => {
      const euSchema: TableSchema = {
        fields: [
          {
            name: 'price',
            type: 'decimal',
            required: true,
            decimalSeparator: ',',
            thousandSeparator: '.',
            decimalPlaces: 2,
          },
        ],
      };

      const raw = [{ price: '1.250,75' }, { price: '3.400.500,20' }];
      const normalized = normalizeTable(raw, euSchema);
      expect(normalized[0].price).to.equal(1250.75);
      expect(normalized[1].price).to.equal(3400500.2);
    });

    it('validates calendar dates strictly rejecting invalid dates like Feb 30', () => {
      const raw = [
        { id: '001', amount: 10, quantity: 1, active: true, date: '2026-02-30' },
      ];

      let error: ReportSchemaError | null = null;
      try {
        normalizeTable(raw, schema);
      } catch (err: unknown) {
        error = err as ReportSchemaError;
      }

      expect(error).to.not.be.null;
      expect(error!.errors[0].field).to.equal('date');
      expect(error!.errors[0].code).to.equal('INVALID_DATE_FORMAT');
    });
    it('normalizes Date instances and valid Excel/Sheets serial numbers (>= 61) to canonical YYYY-MM-DD', () => {
      const dateSchema: TableSchema = {
        fields: [
          { name: 'id', type: 'string', required: true },
          { name: 'date', type: 'date', required: true },
        ],
      };

      const raw = [
        { id: '001', date: new Date('2026-09-19T12:00:00Z') },
        { id: '002', date: 45000 }, // 2023-03-15
        { id: '003', date: 61 },    // 1900-03-01
        { id: '004', date: '2026-09-19' },
      ];

      const normalized = normalizeTable(raw, dateSchema);
      expect(normalized[0].date).to.equal('2026-09-19');
      expect(normalized[1].date).to.equal('2023-03-15');
      expect(normalized[2].date).to.equal('1900-03-01');
      expect(normalized[3].date).to.equal('2026-09-19');
      // Ensure string ID was preserved as string, not parsed as date
      expect(normalized[0].id).to.equal('001');
    });

    it('rejects ambiguous pre-1900-03-01 serial dates (< 61) explicitly to avoid epoch divergence', () => {
      const dateSchema: TableSchema = {
        fields: [{ name: 'date', type: 'date', required: true }],
      };
      let error: ReportSchemaError | null = null;
      try {
        normalizeTable([{ date: 60 }], dateSchema);
      } catch (err: unknown) {
        error = err as ReportSchemaError;
      }
      expect(error).to.not.be.null;
      expect(error!.errors[0].code).to.equal('AMBIGUOUS_SERIAL_DATE');
      expect(error!.errors[0].message).to.include('pre-1900-03-01 epoch ambiguity band');
    });

    it('rejects unexpected extra columns when allowExtraColumns is false', () => {
      const raw = [
        {
          id: '001',
          amount: 10,
          quantity: 1,
          active: true,
          date: '2026-01-01',
          unexpected_extra: 'boom',
        },
      ];

      let error: ReportSchemaError | null = null;
      try {
        normalizeTable(raw, schema);
      } catch (err: unknown) {
        error = err as ReportSchemaError;
      }

      expect(error).to.not.be.null;
      expect(error!.errors[0].field).to.equal('unexpected_extra');
      expect(error!.errors[0].code).to.equal('EXTRA_COLUMN_DISALLOWED');
    });
  });

  describe('Declarative Transforms & Arithmetic (T-11, AC-06)', () => {
    it('evaluates comparison and logical filter conditions securely', () => {
      const row = { status: 'completed', amount: 500, region: 'APAC' };

      expect(evaluateCondition(row, { field: 'status', op: 'eq', value: 'completed' })).to.be.true;
      expect(evaluateCondition(row, { field: 'amount', op: 'gt', value: 400 })).to.be.true;
      expect(evaluateCondition(row, { field: 'amount', op: 'between', values: [100, 1000] })).to.be.true;
      expect(evaluateCondition(row, { field: 'region', op: 'in', values: ['APAC', 'EMEA'] })).to.be.true;
      expect(
        evaluateCondition(row, {
          op: 'and',
          conditions: [
            { field: 'status', op: 'eq', value: 'completed' },
            { field: 'amount', op: 'gte', value: 500 },
          ],
        })
      ).to.be.true;
    });

    it('computes exact decimal expressions (e.g. 5 * 2 * 1.2 = 12.0)', () => {
      const row = { quantity: 5, unitEffort: 2, complexity: 1.2 };
      const expr = {
        op: 'mul',
        args: ['quantity', 'unitEffort', 'complexity'],
      };

      const result = evaluateExpression(row, expr);
      expect(result).to.equal(12);
    });

    it('executes groupSum aggregations with Decimal precision', () => {
      const data = [
        { department: 'Engineering', cost: 1000.5 },
        { department: 'Engineering', cost: 2500.25 },
        { department: 'Marketing', cost: 800.0 },
      ];

      const result = executeTransforms(data, [
        {
          kind: 'groupSum',
          groupBy: ['department'],
          aggregations: [
            { field: 'cost', op: 'sum', as: 'total_cost', round: 2 },
            { field: '*', op: 'count', as: 'headcount' },
          ],
        },
      ]);

      expect(result).to.have.lengthOf(2);
      const eng = result.find((r) => r.department === 'Engineering');
      expect(eng).to.not.be.undefined;
      expect(eng!.total_cost).to.equal(3500.75);
      expect(eng!.headcount).to.equal(2);
    });

    it('executes pivot transformations', () => {
      const data = [
        { region: 'North', quarter: 'Q1', sales: 100 },
        { region: 'North', quarter: 'Q2', sales: 150 },
        { region: 'South', quarter: 'Q1', sales: 200 },
        { region: 'South', quarter: 'Q2', sales: 250 },
      ];

      const result = executeTransforms(data, [
        {
          kind: 'pivot',
          rowGroupBy: ['region'],
          columnPivot: 'quarter',
          valueField: 'sales',
          columnValues: ['Q1', 'Q2'],
        },
      ]);

      expect(result).to.have.lengthOf(2);
      expect(result[0].region).to.equal('North');
      expect(result[0].Q1).to.equal(100);
      expect(result[0].Q2).to.equal(150);
    });

    it('executes lookup joins securely without eval', () => {
      const data = [
        { role: 'FE', hours: 40 },
        { role: 'BE', hours: 50 },
      ];
      const context = {
        ratesTable: [
          { role: 'FE', rate: 75 },
          { role: 'BE', rate: 90 },
        ],
      };

      const result = executeTransforms(
        data,
        [
          {
            kind: 'lookupJoin',
            sourceKey: 'role',
            lookupTable: 'ratesTable',
            lookupKey: 'role',
            select: { hourlyRate: 'rate' },
          },
        ],
        context
      );

      expect(result[0].hourlyRate).to.equal(75);
      expect(result[1].hourlyRate).to.equal(90);
    });
  });
  describe('Table Report & Formula Result Validation', () => {
    it('rejects nested formula objects and non-scalars in formula results', () => {
      const tableTemplate: ReportTemplateV1 = {
        id: 'nested-formula-reject-v1',
        version: 1,
        kind: 'table',
        outputs: [
          {
            name: 'Test',
            columns: [
              { header: 'ID', field: 'id' },
              { header: 'Calc', formulaTemplate: 'A{row}*2', resultField: 'nested' },
            ],
          },
        ],
      };

      const badInput = [
        { id: 1, nested: { formula: 'B2+1', result: 5 } }, // Nested formula object
      ];

      expect(() => buildReport(tableTemplate, badInput)).to.throw(/must be a scalar/);
    });
  });
  describe('Finance Report Preset (T-12, AC-06)', () => {
    const template: ReportTemplateV1 = {
      id: 'finance-test-v1',
      version: 1,
      kind: 'finance',
      title: 'Finance Audit Report',
      config: {
        baseCurrency: 'USD',
        openingBalances: {
          'Chase Primary': 10000.0,
          'Euro Account': 5000.0,
        },
        expectedClosingBalances: {
          'Chase Primary': 14000.0,
          'Euro Account': 10450.0,
        },
        fxRates: [
          { date: '2026-01-01', from: 'EUR', to: 'USD', rate: 1.09, source: 'ECB' },
          { date: '2026-02-01', from: 'EUR', to: 'USD', rate: 1.1, source: 'ECB' },
        ],
      },
    };

    it('builds multi-sheet finance report with FX conversion, refunds, and reconciliation', () => {
      const transactions = [
        {
          id: 'TX-01',
          date: '2026-01-10',
          account: 'Chase Primary',
          direction: 'in',
          amount: 5000.0,
          currency: 'USD',
          description: 'Client Invoice',
        },
        {
          id: 'TX-02',
          date: '2026-01-15',
          account: 'Chase Primary',
          direction: 'out',
          amount: 1000.0,
          currency: 'USD',
          description: 'Software License',
        },
        {
          id: 'TX-03',
          date: '2026-01-20',
          account: 'Euro Account',
          direction: 'in',
          amount: 5000.0,
          currency: 'EUR',
          description: 'EU Sales',
        },
        {
          id: 'TX-04',
          date: '2026-01-25',
          account: 'Euro Account',
          direction: 'in',
          amount: 0.0,
          currency: 'EUR',
          description: 'Zero fee check',
        },
      ];

      const doc = buildReport(template, transactions);
      expect(doc.sheets).to.have.lengthOf(5);
      expect(doc.sheets.map((s) => s.name)).to.deep.equal([
        'Transactions',
        'Monthly Summary',
        'Account Summary',
        'Reconciliation',
        'FX Rates',
      ]);

      // Check transactions formulas and pre-calculated results
      const txSheet = doc.sheets[0];
      expect(txSheet.rows).to.have.lengthOf(6); // 1 header + 4 tx + 1 total
      const row2 = txSheet.rows[1]; // TX-01
      expect(row2[9]).to.deep.equal({ formula: 'G2*I2', result: 5000 }); // Base amount
      expect(row2[10]).to.deep.equal({ formula: 'IF(F2="in",J2,0)', result: 5000 }); // Inflow

      // Check provenance
      expect(doc.provenance.templateId).to.equal('finance-test-v1');
      expect(doc.provenance.sourceHash).to.be.a('string');
      expect((doc.provenance.summary as Record<string, unknown>).baseCurrency).to.equal('USD');
      expect((doc.provenance.summary as Record<string, unknown>).transactionCount).to.equal(4);
    });

    it('rejects duplicate transaction IDs with clear error message', () => {
      const duplicates = [
        { id: 'DUP-01', date: '2026-01-01', account: 'Chase Primary', direction: 'in', amount: 100, currency: 'USD' },
        { id: 'DUP-01', date: '2026-01-02', account: 'Chase Primary', direction: 'in', amount: 200, currency: 'USD' },
      ];

      let threw = false;
      try {
        buildFinanceReport(template, duplicates, 'hash123');
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('Duplicate transaction ID: "DUP-01"');
      }
      expect(threw).to.be.true;
    });

    it('rejects missing FX rate for non-base currency transaction', () => {
      const missingFx = [
        { id: 'TX-JPY', date: '2026-01-01', account: 'Chase Primary', direction: 'in', amount: 10000, currency: 'JPY' },
      ];

      let threw = false;
      try {
        buildFinanceReport(template, missingFx, 'hash123');
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('Missing FX rate for currency pair JPY->USD');
      }
      expect(threw).to.be.true;
    });
    it('fails explicitly when a monetary amount exceeds spreadsheet precision limit (15 significant digits)', () => {
      const hugeTx = [
        { id: 'TX-HUGE', date: '2026-01-01', account: 'Chase Primary', direction: 'in', amount: '9007199254740990.01', currency: 'USD' },
      ];

      expect(() => buildFinanceReport(template, hugeTx, 'hash123')).to.throw(
        /exceeds standard spreadsheet numeric precision/
      );
    });
  });

  describe('Manpower Preset (T-13, AC-06)', () => {
    const template: ReportTemplateV1 = {
      id: 'manpower-test-v1',
      version: 1,
      kind: 'manpower',
      title: 'Engineering Capacity Plan',
      config: {
        currency: 'USD',
        effortUnit: 'person-days',
        globalContingencyPercent: 20,
        roleRates: {
          'Senior Backend': 2500000,
          'Tech Lead': 3500000,
        },
      },
    };

    it('computes exact 5 * 2 * 1.2 then 20% = 14.4 days at 2.5m/day = 36m (AC-06 formula)', () => {
      const tasks = [
        {
          id: 'T-01',
          module: 'API Core',
          role: 'Senior Backend',
          quantity: 5,
          unitEffort: 2,
          complexity: 1.2,
          contingency: 20,
          description: 'Specification example task',
        },
      ];

      const doc = buildReport(template, tasks);
      expect(doc.sheets).to.have.lengthOf(5);
      expect(doc.sheets.map((s) => s.name)).to.deep.equal([
        'Assumptions',
        'Task Details',
        'Module Summary',
        'Role Summary',
        'Scenarios',
      ]);

      const detailsSheet = doc.sheets[1];
      const taskRow = detailsSheet.rows[1];

      // Base effort: 5 * 2 * 1.2 = 12
      expect(taskRow[7]).to.deep.equal({ formula: 'E2*F2*G2', result: 12 });

      // Total effort: 12 * (1 + 0.20) = 14.4
      expect(taskRow[9]).to.deep.equal({ formula: 'H2*(1+I2)', result: 14.4 });

      // Total cost: 14.4 * 2,500,000 = 36,000,000
      expect(taskRow[11]).to.deep.equal({ formula: 'J2*K2', result: 36000000 });

      // Check summary provenance
      const summary = doc.provenance.summary as Record<string, unknown>;
      expect(summary.totalBaseEffort).to.equal('12.00');
      expect(summary.totalEffort).to.equal('14.40');
      expect(summary.totalCost).to.equal('36000000.00');
    });

    it('rejects missing role rates without inventing default rates', () => {
      const unknownRoleTask = [
        {
          id: 'T-99',
          module: 'Design',
          role: 'Unknown Role Nonexistent',
          quantity: 1,
          unitEffort: 1,
        },
      ];

      let threw = false;
      try {
        buildManpowerReport(template, unknownRoleTask, 'hash123');
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('Missing role rate for role "Unknown Role Nonexistent"');
      }
      expect(threw).to.be.true;
    });
    it('maintains delta arithmetic regardless of scenario display names (rename-invariance)', () => {
      const tasks = [
        {
          id: 'T-01',
          module: 'Core',
          role: 'Senior Backend',
          quantity: 2,
          unitEffort: 5,
          complexity: 1.2,
          contingency: 20,
        },
      ];

      // Template with custom scenario names not containing "base"
      const customNamedTemplate: ReportTemplateV1 = {
        id: 'manpower-custom-names-v1',
        version: 1,
        kind: 'manpower',
        config: {
          roleRates: { 'Senior Backend': 2500000 },
          globalContingencyPercent: 20,
          scenarios: [
            { name: 'Plan Alpha', complexityMultiplier: 1.0, contingencyPercent: 20 },
            { name: 'Plan Beta', complexityMultiplier: 1.25, contingencyPercent: 35 },
          ],
        },
      };

      const doc = buildReport(customNamedTemplate, tasks);
      const scenarioSheet = doc.sheets[4];

      // Plan Alpha has complexity 1.0, cont 20% -> identical to base (cost delta = 0)
      const alphaRow = scenarioSheet.rows[1];
      expect(alphaRow[0]).to.equal('Plan Alpha');
      expect(alphaRow[5]).to.equal(36000000); // Total cost
      expect(alphaRow[6]).to.equal(0);        // Cost Delta vs Base
      expect(alphaRow[7]).to.equal(0);        // % Delta vs Base
    });

    it('supports explicit baselineScenario configuration and validates existence', () => {
      const tasks = [
        { id: 'T-01', module: 'Core', role: 'Senior Backend', quantity: 2, unitEffort: 5, complexity: 1.2, contingency: 20 },
      ];

      const templateWithBaseline: ReportTemplateV1 = {
        id: 'manpower-baseline-config-v1',
        version: 1,
        kind: 'manpower',
        config: {
          roleRates: { 'Senior Backend': 2500000 },
          baselineScenario: 'Custom Target Baseline',
          scenarios: [
            { name: 'Custom Target Baseline', complexityMultiplier: 1.0, contingencyPercent: 20 },
            { name: 'Pessimistic Scenario', complexityMultiplier: 1.25, contingencyPercent: 35 },
          ],
        },
      };

      const doc = buildReport(templateWithBaseline, tasks);
      const summary = doc.provenance.summary as Record<string, unknown>;
      expect(summary.baselineScenario).to.equal('Custom Target Baseline');

      // Invalid baselineScenario name throws clear error
      const invalidBaselineTemplate: ReportTemplateV1 = {
        ...templateWithBaseline,
        config: {
          ...templateWithBaseline.config,
          baselineScenario: 'Nonexistent Scenario Name',
        },
      };
      expect(() => buildReport(invalidBaselineTemplate, tasks)).to.throw(/Configured baselineScenario "Nonexistent Scenario Name" was not found/);
    });
  });

  describe('Deterministic Provenance & Security (AC-02, AC-06, AC-07)', () => {
    it('produces identical sourceHash for identical data across reruns', () => {
      const inputA = [{ b: 2, a: 1 }, { c: 3 }];
      const inputB = [{ a: 1, b: 2 }, { c: 3 }];

      const hashA = computeSourceHash(inputA);
      const hashB = computeSourceHash(inputB);

      expect(hashA).to.equal(hashB);
    });

    it('rejects unknown template kinds at runtime', () => {
      const badTemplate = {
        id: 'bad-v1',
        version: 1,
        kind: 'unsupported_custom_kind',
      };

      let threw = false;
      try {
        buildReport(badTemplate, []);
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('Unknown template kind');
      }
      expect(threw).to.be.true;
    });
  });
});
