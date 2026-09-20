import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import GoogleSheet from '../src/lib/google-sheet';
import { FakeSheets } from './fake-sheets';

const SPREADSHEET_ID = 'schema-transport';

interface SchemaCommandReceipt {
  operation: string;
  worksheetTitle: string;
  sheetId: number;
  sample: { range: string; startRow: number; startColumn: number; rows: number; columns: number };
  columns: {
    header: string;
    headerA1: string;
    column: number;
    letter: string;
    inferredTypes: Record<string, number>;
    formulaCells: number;
    validations: { condition: string; count: number }[];
  }[];
  warnings: string[];
  namedRanges: { name: string; sheetId?: number; range: string }[];
}

interface ValidateCommandReceipt {
  operation: string;
  range: string;
  rowsChecked: number;
  valid: boolean;
  issues: { field?: string; row?: number; column?: number; a1?: string; code?: string; message?: string; value?: unknown }[];
  warnings: string[];
}

/**
 * data:schema and data:validate, end to end over the real client and the fake transport.
 *
 * The schema path is the only GoogleSheet call that reads spreadsheets.get gridData, and the
 * offline command suite stubs that metadata layer entirely - so a wrong ranges encoding, field
 * mask or gridData interpretation would pass offline and only fail on a live run. These cases
 * drive the real oclif commands against the fake's GridData and named-range rendering: offset
 * samples, typed entered/effective values, formulas, formatted text, validation rules,
 * named-range scoping, and violation coordinates reported at absolute sheet positions for
 * invalid required/type/unique data.
 */
describe('data:schema and data:validate over the real client and fake transport', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;
  let dataSheetId: number;

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    const spreadsheet = fake.addSpreadsheet(SPREADSHEET_ID, 'Schema Transport', [
      { title: 'Data', rowCount: 50, columnCount: 10 },
      { title: 'Aux', rowCount: 10, columnCount: 5 },
      { title: 'Offset', rowCount: 20, columnCount: 6 },
    ]);
    const sheetIdOf = (title: string): number => spreadsheet.sheets.filter((sheet) => sheet.title === title)[0].sheetId;
    dataSheetId = sheetIdOf('Data');

    fake.setCells(SPREADSHEET_ID, 'Data', 'A1:F3', [
      ['ID', 'Name', 'Score', 'Active', 'Computed', 'Rate'],
      [1, 'Widget', 90, 'TRUE', '=C2*2', 0.15],
      [2, 'Gadget', 85.5, 'FALSE', '=C3*2', 0.25],
    ]);
    fake.setCells(SPREADSHEET_ID, 'Offset', 'A3:B5', [
      ['Key', 'Qty'],
      ['k1', 5],
      ['k2', 6],
    ]);
    fake.setCells(SPREADSHEET_ID, 'Aux', 'A1', [['elsewhere']]);

    // A real server evaluates formulas and applies number formats; the fake cannot compute, so
    // the fixture states what the transport should carry back as effective/formatted values.
    fake.setCellMeta(SPREADSHEET_ID, 'Data', 'E2', { effectiveValue: { numberValue: 180 } });
    fake.setCellMeta(SPREADSHEET_ID, 'Data', 'E3', { effectiveValue: { numberValue: 171 } });
    fake.setCellMeta(SPREADSHEET_ID, 'Data', 'F2', { formattedValue: '15%' });
    fake.setCellMeta(SPREADSHEET_ID, 'Data', 'F3', { formattedValue: '25%' });
    const trueFalse = {
      condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'TRUE' }, { userEnteredValue: 'FALSE' }] },
      showCustomUi: true,
    };
    fake.setCellMeta(SPREADSHEET_ID, 'Data', 'D2', { dataValidation: trueFalse });
    fake.setCellMeta(SPREADSHEET_ID, 'Data', 'D3', { dataValidation: trueFalse });

    // One named range over the sampled sheet, one over another sheet the client filters out.
    spreadsheet.namedRanges = [
      { namedRangeId: 'nr-scores', name: 'Scores', range: { sheetId: dataSheetId, startRowIndex: 0, endRowIndex: 3, startColumnIndex: 2, endColumnIndex: 3 } },
      { namedRangeId: 'nr-aux', name: 'AuxCells', range: { sheetId: sheetIdOf('Aux'), startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 } },
    ];

    gsheet = new GoogleSheet(SPREADSHEET_ID);
    await gsheet.authorize(fake.credentials);
  });

  describe('worksheet metadata transport', () => {
    it('issues one scoped spreadsheets.get with the bounded range and the cell-level field mask', async () => {
      const metadata = await gsheet.getWorksheetMetadata({ worksheetTitle: 'Data', range: 'A1:F3' });

      const gets = fake.requests.filter((request) => request.method === 'GET' && request.url.includes(`/v4/spreadsheets/${SPREADSHEET_ID}?`));
      expect(gets, 'exactly one spreadsheets.get').to.have.length(1);
      const url = decodeURIComponent(gets[0].url);
      expect(url).to.contain('ranges=Data!A1:F3');
      // a wrong field mask would 400 on live Google, so the request itself is part of the contract
      for (const maskPart of ['startRow', 'startColumn', 'userEnteredValue', 'effectiveValue', 'formattedValue', 'dataValidation', 'namedRanges']) {
        expect(url, `field mask should ask for ${maskPart}`).to.contain(maskPart);
      }

      expect(metadata.properties.title).to.equal('Data');
      expect(metadata.properties.sheetId).to.equal(dataSheetId);
      // named ranges come back spreadsheet-wide; the client keeps only this sheet's
      expect(metadata.namedRanges.map((namedRange) => namedRange.name)).to.eql(['Scores']);
    });

    it('renders typed cell detail: entered values, formula results, formatted text, validation', async () => {
      const metadata = await gsheet.getWorksheetMetadata({ worksheetTitle: 'Data', range: 'A1:F3' });

      expect(metadata.gridData).to.have.length(1);
      const grid = metadata.gridData[0];
      expect(grid.startRow).to.equal(0);
      expect(grid.startColumn).to.equal(0);
      const rows = grid.rowData ?? [];
      expect(rows).to.have.length(3);

      // the header row keeps its text positionally
      expect(rows[0].values?.[0]?.formattedValue).to.equal('ID');
      // stored text arrives typed, not as strings
      expect(rows[1].values?.[0]?.userEnteredValue).to.eql({ numberValue: 1 });
      expect(rows[1].values?.[2]?.userEnteredValue).to.eql({ numberValue: 90 });
      expect(rows[1].values?.[3]?.userEnteredValue).to.eql({ boolValue: true });
      expect(rows[2].values?.[2]?.userEnteredValue).to.eql({ numberValue: 85.5 });
      expect(rows[1].values?.[1]?.userEnteredValue).to.eql({ stringValue: 'Widget' });
      // the formula cell carries its formula plus the fixture-provided evaluated value
      expect(rows[1].values?.[4]?.userEnteredValue).to.eql({ formulaValue: '=C2*2' });
      expect(rows[1].values?.[4]?.effectiveValue).to.eql({ numberValue: 180 });
      expect(rows[2].values?.[4]?.effectiveValue).to.eql({ numberValue: 171 });
      // the percent-formatted rate is a number with display text
      expect(rows[1].values?.[5]?.effectiveValue).to.eql({ numberValue: 0.15 });
      expect(rows[1].values?.[5]?.formattedValue).to.equal('15%');
      // the validation rule rides on the cell
      expect(rows[1].values?.[3]).to.have.nested.property('dataValidation.condition.type', 'ONE_OF_LIST');
    });

    it('honors an offset sample window with 0-based gridData anchors', async () => {
      const metadata = await gsheet.getWorksheetMetadata({ worksheetTitle: 'Data', range: 'B2:C3' });

      const grid = metadata.gridData[0];
      expect(grid.startRow).to.equal(1);
      expect(grid.startColumn).to.equal(1);
      const rows = grid.rowData ?? [];
      expect(rows).to.have.length(2);
      // positional values restart at the requested anchor: values[0] is column B
      expect(rows[0].values?.[0]?.userEnteredValue).to.eql({ stringValue: 'Widget' });
      expect(rows[0].values?.[1]?.userEnteredValue).to.eql({ numberValue: 90 });
    });

    it('refuses a requested range naming an unknown worksheet the way Google does', async () => {
      try {
        await gsheet.getWorksheetMetadata({ worksheetTitle: 'Ghost', range: 'A1:B2' });
        expect.fail('should have thrown');
      } catch (error) {
        expect(error).to.have.property('message').that.contains('Unable to parse range');
      }
    });
  });

  describe('commands', () => {
    let savedEnv: Record<string, string | undefined>;

    // The key cannot travel on the command line (the PEM header has spaces), so it rides the
    // environment exactly the way test/commands/offline.test.ts and copy-export manage theirs.
    const MANAGED_ENV = ['GSHEET_CLIENT_EMAIL', 'GSHEET_PRIVATE_KEY', 'GSHEET_CREDENTIALS_FILE', 'GSHEET_USE_OAUTH', 'GSHEET_CLIENT_SECRET_FILE', 'SPREADSHEET_ID', 'WORKSHEET_TITLE'];
    const authFlags = [`--clientEmail=${fake.credentials.client_email}`];

    beforeEach(() => {
      savedEnv = Object.fromEntries(MANAGED_ENV.map((name) => [name, process.env[name]]));
      for (const name of MANAGED_ENV) delete process.env[name];
      process.env.GSHEET_PRIVATE_KEY = fake.credentials.private_key;
    });

    afterEach(() => {
      for (const [name, value] of Object.entries(savedEnv)) {
        if (value === undefined) delete process.env[name];
        else process.env[name] = value;
      }
    });

    describe('data:schema', () => {
      it('reports headers, inferred types, formulas, validations and scoped named ranges', async () => {
        const { error, result } = await runCommand<SchemaCommandReceipt>([
          'data:schema',
          `--spreadsheetId=${SPREADSHEET_ID}`,
          '--worksheetTitle=Data',
          '--maxRow=3',
          '--maxCol=6',
          ...authFlags,
          '--rawOutput',
        ]);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(result.sample).to.include({ startRow: 1, startColumn: 1, rows: 3, columns: 6 });
        const byHeader = new Map<string, SchemaCommandReceipt['columns'][number]>();
        for (const column of result.columns) byHeader.set(column.header, column);
        expect(byHeader.get('Score')?.inferredTypes).to.include({ number: 2 });
        expect(byHeader.get('Rate')?.inferredTypes).to.include({ number: 2 });
        expect(byHeader.get('Active')?.inferredTypes).to.include({ boolean: 2 });
        expect(byHeader.get('Computed')?.formulaCells).to.equal(2);
        expect(byHeader.get('Active')?.validations).to.eql([{ condition: 'ONE_OF_LIST', count: 2 }]);
        // the range named over the other sheet is not this sheet's business
        expect(result.namedRanges).to.eql([{ name: 'Scores', sheetId: dataSheetId, range: 'C1:C3' }]);
        expect(result.warnings).to.eql([]);
      });

      it('resolves an offset header row to absolute coordinates', async () => {
        const { error, result } = await runCommand<SchemaCommandReceipt>([
          'data:schema',
          `--spreadsheetId=${SPREADSHEET_ID}`,
          '--worksheetTitle=Offset',
          '--minRow=3',
          '--maxRow=5',
          '--maxCol=2',
          ...authFlags,
          '--rawOutput',
        ]);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(result.sample.startRow).to.equal(3);
        expect(result.sample.startColumn).to.equal(1);
        expect(result.columns.map((column) => column.header)).to.eql(['Key', 'Qty']);
        expect(result.columns[0].headerA1).to.equal('A3');
        expect(result.columns[1].inferredTypes).to.include({ number: 2 });
      });
    });

    describe('data:validate', () => {
      const validateArgs = (schema: string, extra: string[] = []): string[] => [
        'data:validate',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        '--worksheetTitle=Data',
        '--maxRow=3',
        '--maxCol=4',
        `--schema=${schema}`,
        ...authFlags,
        ...extra,
      ];

      it('passes well-formed data and checks the sampled rows', async () => {
        const schema = '{"fields":[{"name":"ID","type":"integer","required":true},{"name":"Name","type":"string"},{"name":"Score","type":"decimal"},{"name":"Active","type":"boolean"}]}';
        const { error, result } = await runCommand<ValidateCommandReceipt>([...validateArgs(schema), '--rawOutput']);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(result.valid).to.equal(true);
        expect(result.rowsChecked).to.equal(2);
        expect(result.issues).to.eql([]);
      });

      it('reports required and type violations at absolute coordinates through the envelope', async () => {
        // row 3: the ID cell is blanked ('' renders as an empty cell, and setCells skips null),
        // so ID trips required while Score and Active hold wrong types
        fake.setCells(SPREADSHEET_ID, 'Data', 'A3:D3', [['', 'Fine', 'oops', 'MAYBE']]);
        const schema = '{"fields":[{"name":"ID","type":"integer","required":true},{"name":"Name","type":"string"},{"name":"Score","type":"decimal"},{"name":"Active","type":"boolean"}]}';

        const { error, stderr } = await runCommand([...validateArgs(schema), '--json']);
        expect(error).to.exist;
        const envelope = JSON.parse(stderr);
        expect(envelope.error.code).to.equal('DATA_INVALID');
        const issues = envelope.error.issues;
        expect(issues.find((issue: { field?: string }) => issue.field === 'ID')).to.include({ row: 3, column: 1, a1: 'A3' });
        expect(issues.find((issue: { field?: string }) => issue.field === 'Score')).to.include({ row: 3, column: 3, a1: 'C3' });
        expect(issues.find((issue: { field?: string }) => issue.field === 'Active')).to.include({ row: 3, column: 4, a1: 'D3' });
        // every coordinate is a real sheet position, never 0
        for (const issue of issues) expect(issue.row).to.be.above(0);
      });

      it('reports a unique violation with the first-seen row and exact coordinates', async () => {
        fake.setCells(SPREADSHEET_ID, 'Data', 'A3', [[1]]);
        const schema = '{"fields":[{"name":"ID","type":"integer","required":true,"unique":true},{"name":"Name","type":"string"}]}';

        const { error, stderr } = await runCommand([
          'data:validate',
          `--spreadsheetId=${SPREADSHEET_ID}`,
          '--worksheetTitle=Data',
          '--maxRow=3',
          '--maxCol=2',
          `--schema=${schema}`,
          ...authFlags,
          '--json',
        ]);
        expect(error).to.exist;
        const envelope = JSON.parse(stderr);
        expect(envelope.error.code).to.equal('DATA_INVALID');
        const duplicates = envelope.error.issues.filter((issue: { code?: string }) => issue.code === 'DUPLICATE_VALUE');
        expect(duplicates).to.have.lengthOf(1);
        expect(duplicates[0]).to.include({ row: 3, column: 1, a1: 'A3', value: 1 });
        expect(duplicates[0].message).to.contain('first seen at sheet row 2');
      });

      it('maps violations to absolute sheet rows when the header sits below row 1', async () => {
        fake.setCells(SPREADSHEET_ID, 'Offset', 'B5', [['not-a-number']]);
        const schema = '{"fields":[{"name":"Key","type":"string","required":true},{"name":"Qty","type":"integer"}]}';

        const { error, stderr } = await runCommand([
          'data:validate',
          `--spreadsheetId=${SPREADSHEET_ID}`,
          '--worksheetTitle=Offset',
          '--minRow=3',
          '--maxRow=5',
          '--maxCol=2',
          `--schema=${schema}`,
          ...authFlags,
          '--json',
        ]);
        expect(error).to.exist;
        const envelope = JSON.parse(stderr);
        expect(envelope.error.code).to.equal('DATA_INVALID');
        // absolute sheet row 5, not the third row of the sample
        expect(envelope.error.issues.find((issue: { field?: string }) => issue.field === 'Qty')).to.include({ row: 5, column: 2, a1: 'B5' });
      });
    });
  });
});
