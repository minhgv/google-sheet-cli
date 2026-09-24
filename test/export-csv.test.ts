import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { generateKeyPairSync } from 'crypto';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import * as factory from '../src/lib/factory';
import GoogleSheet, { GoogleSheetCli } from '../src/lib/google-sheet';
import { XlsxWorkbook } from '../src/lib/xlsx';

/**
 * Offline proof for `data:export-csv` (T-08): the command layer over the locked CSV
 * serializer, against the same factory seam the offline command suite uses - no credentials,
 * no network, no real Sheets client. The XLSX cases run against real workbooks built with
 * XlsxWorkbook and saved to a tmpdir, exactly like test/xlsx.test.ts builds its fixtures.
 */

const SPREADSHEET_ID = 'export-csv-spreadsheet';
const WORKSHEET_TITLE = 'ExportSheet';
const CLIENT_EMAIL = 'export-csv@example.iam.gserviceaccount.com';

// `normalizeCredentials` parses the key with `createPrivateKey`, so a placeholder would fail
// before the command is reached. Throwaway key that exists only inside the test process.
const PRIVATE_KEY = generateKeyPairSync('rsa', {
  modulusLength: 2048,
  privateKeyEncoding: { format: 'pem', type: 'pkcs8' },
  publicKeyEncoding: { format: 'pem', type: 'spki' },
}).privateKey;

/** The env every flag with an `env:` binding reads, so the suite is the same run to run. */
const MANAGED_ENV = [
  'GSHEET_CLIENT_EMAIL',
  'GSHEET_PRIVATE_KEY',
  'GSHEET_CREDENTIALS_FILE',
  'SPREADSHEET_ID',
  'WORKSHEET_TITLE',
];

interface StubCall {
  method: string;
  args: unknown[];
}

class StubGoogleSheet {
  public readonly calls: StubCall[] = [];
  public data: GoogleSheetCli.SheetData = {
    rawData: [
      ['A1', 'B1', 'C1'],
      ['A2', '', 'C2'],
    ],
    formatted: [],
    header: [],
    range: `${WORKSHEET_TITLE}!A1:C2`,
  };

  async authorize(credentials: GoogleSheetCli.Credentials): Promise<void> {
    this.calls.push({ method: 'authorize', args: [credentials] });
  }

  async getData(options: GoogleSheetCli.QueryOptions, spreadsheetId?: string): Promise<GoogleSheetCli.SheetData> {
    this.calls.push({ method: 'getData', args: [options, spreadsheetId] });
    return this.data;
  }
}

const mutableFactory = factory as { createGoogleSheet: () => GoogleSheet };
const realCreateGoogleSheet = factory.createGoogleSheet;

/** Formula with a cached result, formula without one, and a literal string starting with "=". */
async function buildFixtureWorkbook(dir: string): Promise<string> {
  const workbook = XlsxWorkbook.create();
  const ws = workbook.workbook.addWorksheet('Data');
  ws.getCell('A1').value = 'Name';
  ws.getCell('B1').value = 1500000;
  ws.getCell('A2').value = 'Total';
  ws.getCell('B2').value = { formula: 'B1*2', result: 3000000 };
  ws.getCell('C2').value = { formula: 'SUM(A1:B2)', result: undefined };
  ws.getCell('D2').value = '=LiteralEqual';
  const filePath = path.join(dir, 'fixture.xlsx');
  await workbook.save(filePath);
  return filePath;
}

describe('data:export-csv', () => {
  let stub: StubGoogleSheet;
  let savedEnv: Record<string, string | undefined>;
  let tmpDir: string;

  beforeEach(async () => {
    // Same env hygiene as the offline command suite: flags with env bindings must see a
    // controlled environment, and the private key cannot travel on a command line.
    savedEnv = Object.fromEntries(MANAGED_ENV.map((name) => [name, process.env[name]]));
    for (const name of MANAGED_ENV) delete process.env[name];
    process.env.GSHEET_PRIVATE_KEY = PRIVATE_KEY;

    stub = new StubGoogleSheet();
    mutableFactory.createGoogleSheet = () => stub as unknown as GoogleSheet;

    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'export-csv-'));
  });

  afterEach(async () => {
    mutableFactory.createGoogleSheet = realCreateGoogleSheet;
    for (const [name, value] of Object.entries(savedEnv)) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
    await fs.promises.rm(tmpDir, { recursive: true, force: true });
  });

  const sheetsArgv = (extra: string[]): string[] => [
    'data:export-csv',
    `--spreadsheetId=${SPREADSHEET_ID}`,
    `--worksheetTitle=${WORKSHEET_TITLE}`,
    `--clientEmail=${CLIENT_EMAIL}`,
    ...extra,
  ];

  describe('Sheets source', () => {
    it('writes the serialized rawData to stdout, exactly and without any receipt', async () => {
      const { error, stdout } = await runCommand(sheetsArgv([]));
      if (error) throw error;
      expect(stdout).to.equal('A1,B1,C1\r\nA2,,C2\r\n');
    });

    it('keeps stdout pure CSV even under --rawOutput (no JSON receipt on the data channel)', async () => {
      const { error, stdout } = await runCommand(sheetsArgv(['--rawOutput']));
      if (error) throw error;
      expect(stdout).to.equal('A1,B1,C1\r\nA2,,C2\r\n');
    });

    it('queries the spreadsheet through the factory client with the source and mode mapped onto render options', async () => {
      const { error } = await runCommand(sheetsArgv(['--mode=formula']));
      if (error) throw error;

      const getDataCalls = stub.calls.filter(({ method }) => method === 'getData');
      expect(getDataCalls).to.have.lengthOf(1);
      const [options, spreadsheetId] = getDataCalls[0].args as [GoogleSheetCli.QueryOptions, string];
      expect(spreadsheetId).to.equal(SPREADSHEET_ID);
      expect(options.worksheetTitle).to.equal(WORKSHEET_TITLE);
      expect(options.valueRenderOption).to.equal(GoogleSheetCli.ValueRenderOption.FORMULA);
    });

    it('maps raw and formatted modes onto UNFORMATTED_VALUE and FORMATTED_VALUE with matching date rendering', async () => {
      const raw = await runCommand(sheetsArgv(['--mode=raw']));
      if (raw.error) throw raw.error;
      const formatted = await runCommand(sheetsArgv(['--mode=formatted']));
      if (formatted.error) throw formatted.error;

      const calls = stub.calls.filter(({ method }) => method === 'getData');
      expect(calls).to.have.lengthOf(2);
      const [rawOptions] = calls[0].args as [GoogleSheetCli.QueryOptions];
      expect(rawOptions.valueRenderOption).to.equal(GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE);
      expect(rawOptions.dateTimeRenderOption).to.equal(GoogleSheetCli.DateTimeRenderOption.SERIAL_NUMBER);
      const [formattedOptions] = calls[1].args as [GoogleSheetCli.QueryOptions];
      expect(formattedOptions.valueRenderOption).to.equal(GoogleSheetCli.ValueRenderOption.FORMATTED_VALUE);
      expect(formattedOptions.dateTimeRenderOption).to.equal(GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING);
    });

    it('safe policy prefixes dangerous string cells but never typed numbers, and the receipt says so', async () => {
      stub.data = {
        rawData: [
          ['=SUM(A1)', '+1', 'plain'],
          ['@cmd', -2, '-3'],
        ],
        formatted: [],
        header: [],
        range: `${WORKSHEET_TITLE}!A1:C2`,
      };
      const output = path.join(tmpDir, 'safe.csv');
      const { error, result } = await runCommand(sheetsArgv(['--output', output, '--rawOutput']));
      if (error) throw error;

      const receipt = result as Record<string, unknown>;
      expect(receipt.source).to.equal('sheets');
      expect(receipt.injectionPolicy).to.equal('safe');
      expect(receipt.transformedCells).to.equal(4);

      const csv = await fs.promises.readFile(output, 'utf8');
      expect(csv).to.equal("'=SUM(A1),'+1,plain\r\n'@cmd,-2,'-3\r\n");
    });

    it('preserve policy passes dangerous cells byte-faithfully and warns on stderr in stdout mode', async () => {
      stub.data = {
        rawData: [['=SUM(A1)', 'plain']],
        formatted: [],
        header: [],
        range: `${WORKSHEET_TITLE}!A1:B1`,
      };
      const { error, stdout, stderr } = await runCommand(sheetsArgv(['--injection=preserve']));
      if (error) throw error;
      expect(stdout).to.equal('=SUM(A1),plain\r\n');
      expect(stderr).to.contain('preserve');
      expect(stderr).to.contain('formula-dangerous');
    });

    it('pads short rows with empty fields up to the widest row and reports it in the receipt', async () => {
      stub.data = {
        rawData: [
          ['A', 'B', 'C'],
          ['D'],
        ],
        formatted: [],
        header: [],
        range: `${WORKSHEET_TITLE}!A1:C2`,
      };
      const output = path.join(tmpDir, 'ragged.csv');
      const { error, result } = await runCommand(sheetsArgv(['--output', output, '--rawOutput']));
      if (error) throw error;

      const csv = await fs.promises.readFile(output, 'utf8');
      expect(csv).to.equal('A,B,C\r\nD,,\r\n');
      const warnings = (result as { warnings: string[] }).warnings;
      expect(warnings.some((warning) => warning.includes('padded with empty fields'))).to.be.true;
    });

    it('writes the output atomically as a 0600 file and refuses an existing one without --overwrite', async () => {
      const output = path.join(tmpDir, 'out.csv');
      const first = await runCommand(sheetsArgv(['--output', output, '--rawOutput']));
      if (first.error) throw first.error;

      const receipt = first.result as { savedPath: string; bytesWritten: number; rows: number; columns: number; mode: string };
      expect(receipt.savedPath).to.equal(path.resolve(output));
      expect(receipt.bytesWritten).to.equal('A1,B1,C1\r\nA2,,C2\r\n'.length);
      expect(receipt.rows).to.equal(2);
      expect(receipt.columns).to.equal(3);
      expect(receipt.mode).to.equal('raw');
      const stat = await fs.promises.stat(output);
      expect(stat.mode & 0o777).to.equal(0o600);

      const refused = await runCommand(sheetsArgv(['--output', output]));
      expect(refused.error, 'refuses to replace without --overwrite').to.exist;
      expect(String(refused.error?.message)).to.contain('--overwrite');

      const replaced = await runCommand(sheetsArgv(['--output', output, '--overwrite', '--rawOutput']));
      if (replaced.error) throw replaced.error;
      expect((replaced.result as { source: string }).source).to.equal('sheets');
    });
  });

  describe('XLSX source', () => {
    it('exports the fixture workbook to CSV bytes without touching the Sheets client', async () => {
      const workbookPath = await buildFixtureWorkbook(tmpDir);
      const output = path.join(tmpDir, 'exported.csv');
      const { error, result } = await runCommand([
        'data:export-csv',
        `--workbook=${workbookPath}`,
        '--range=Data!A1:D2',
        '--injection=preserve',
        '--output',
        output,
        '--rawOutput',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      if (error) throw error;

      // Cached formula result materializes as its number; the formula without a cache is an
      // empty field; the literal "=" string passes through under preserve.
      const csv = await fs.promises.readFile(output, 'utf8');
      expect(csv).to.equal('Name,1500000,,\r\nTotal,3000000,,=LiteralEqual\r\n');

      const receipt = result as Record<string, unknown>;
      expect(receipt.source).to.equal('xlsx');
      expect(receipt.mode).to.equal('raw');
      expect(receipt.rows).to.equal(2);
      expect(receipt.columns).to.equal(4);
      expect(String(receipt.range)).to.contain('Data');
      // only the startup authorization, never a Sheets API call
      expect(stub.calls.every(({ method }) => method === 'authorize')).to.be.true;
    });

    it('formula mode yields formula text: verbatim under preserve, apostrophe-prefixed under safe', async () => {
      const workbookPath = await buildFixtureWorkbook(tmpDir);
      const preserve = await runCommand([
        'data:export-csv',
        `--workbook=${workbookPath}`,
        '--range=Data!A2:D2',
        '--mode=formula',
        '--injection=preserve',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      if (preserve.error) throw preserve.error;
      expect(preserve.stdout).to.equal('Total,=B1*2,=SUM(A1:B2),=LiteralEqual\r\n');

      const safe = await runCommand([
        'data:export-csv',
        `--workbook=${workbookPath}`,
        '--range=Data!A2:D2',
        '--mode=formula',
        '--injection=safe',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      if (safe.error) throw safe.error;
      expect(safe.stdout).to.equal("Total,'=B1*2,'=SUM(A1:B2),'=LiteralEqual\r\n");
    });

    it('warns in the receipt that raw mode exports cached formula results that are never recalculated', async () => {
      const workbookPath = await buildFixtureWorkbook(tmpDir);
      const output = path.join(tmpDir, 'cached.csv');
      const { error, result } = await runCommand([
        'data:export-csv',
        `--workbook=${workbookPath}`,
        '--range=Data!A1:D2',
        '--output',
        output,
        '--rawOutput',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      if (error) throw error;

      const receipt = result as { warnings: string[]; transformedCells: number };
      expect(receipt.warnings.some((warning) => warning.includes('cached results') && warning.includes('never recalculated'))).to.be.true;
      expect(receipt.warnings.some((warning) => warning.includes('no cached result') && warning.includes('empty fields'))).to.be.true;
      // the literal "=..." string cell under the default safe policy
      expect(receipt.transformedCells).to.equal(1);
    });

    it('reports formula cells without a cached result when exporting formula text', async () => {
      const workbookPath = await buildFixtureWorkbook(tmpDir);
      const output = path.join(tmpDir, 'formulas.csv');
      const { error, result } = await runCommand([
        'data:export-csv',
        `--workbook=${workbookPath}`,
        '--range=Data!A1:D2',
        '--mode=formula',
        '--output',
        output,
        '--rawOutput',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      if (error) throw error;

      const warnings = (result as { warnings: string[] }).warnings;
      expect(warnings.some((warning) => warning.includes('no cached result') && warning.includes('formula text only'))).to.be.true;
    });
  });

  describe('source selection and errors', () => {
    it('refuses both sources at once', async () => {
      const { error } = await runCommand(
        sheetsArgv(['--workbook=whatever.xlsx'])
      );
      expect(error).to.exist;
      expect(String(error?.message)).to.contain('cannot be combined');
    });

    it('refuses when no source is given', async () => {
      const { error } = await runCommand(['data:export-csv', `--clientEmail=${CLIENT_EMAIL}`]);
      expect(error).to.exist;
      expect(String(error?.message)).to.contain('No source given');
    });

    it('refuses the Sheets source without --worksheetTitle', async () => {
      const { error } = await runCommand(['data:export-csv', `--spreadsheetId=${SPREADSHEET_ID}`, `--clientEmail=${CLIENT_EMAIL}`]);
      expect(error).to.exist;
      expect(String(error?.message)).to.contain('--worksheetTitle');
    });

    it('refuses the workbook source without --range', async () => {
      const workbookPath = await buildFixtureWorkbook(tmpDir);
      const { error } = await runCommand(['data:export-csv', `--workbook=${workbookPath}`, `--clientEmail=${CLIENT_EMAIL}`]);
      expect(error).to.exist;
      expect(String(error?.message)).to.contain('--range is required');
    });

    it('refuses an unsupported mode at parse time', async () => {
      const { error } = await runCommand(sheetsArgv(['--mode=bogus']));
      expect(error).to.exist;
      expect(String(error?.message)).to.contain('--mode');
    });

    it('fails loudly on a missing workbook file', async () => {
      const { error } = await runCommand([
        'data:export-csv',
        `--workbook=${path.join(tmpDir, 'absent.xlsx')}`,
        '--range=Data!A1:B2',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      expect(error).to.exist;
    });

    it('fails loudly when the workbook range names a missing sheet', async () => {
      const workbookPath = await buildFixtureWorkbook(tmpDir);
      const { error } = await runCommand([
        'data:export-csv',
        `--workbook=${workbookPath}`,
        '--range=Missing!A1:B2',
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      expect(error).to.exist;
      expect(String(error?.message)).to.contain('not found');
    });
  });
});
