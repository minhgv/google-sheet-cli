import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { generateKeyPairSync } from 'crypto';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import GoogleSheet, { GoogleSheetCli } from '../src/lib/google-sheet';
import { XlsxSaveResult } from '../src/lib/xlsx-types';
import { FakeSheets, fakeExportBytes, XLSX_MIME_TYPE } from './fake-sheets';

type CopyCommandReceipt = GoogleSheetCli.CopySpreadsheetResult & { operation?: string };
type WorksheetCopyReceipt = GoogleSheetCli.CopyWorksheetResult & { operation?: string };
type ExportCommandReceipt = Omit<GoogleSheetCli.ExportResult, 'bytes'> & XlsxSaveResult & { operation?: string };

const SOURCE_ID = 'copy-source';
const DEST_ID = 'copy-dest';
const SOURCE_TITLE = 'Quarterly Report';

/**
 * Spreadsheet/worksheet copy and PDF/XLSX export, against the fake over the real client.
 *
 * The copy paths are proven end to end: the endpoints they hit, the receipts they return, that
 * the source stays untouched while data (formulas included) really arrives at the copy, and that
 * an ambiguous failure is never replayed. Export is proven byte-exact: the bytes that reach the
 * local file are the bytes Google sent, an existing file is never clobbered without --overwrite,
 * and a failed run leaves whatever was there before in place.
 *
 * The command cases drive the real oclif command classes the way a user would - credentials on
 * flags/env, the factory untouched - so the flag surface, the receipt and the file handling are
 * all the shipped ones.
 */
describe('spreadsheet and worksheet copy, spreadsheet export', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    fake.addSpreadsheet(SOURCE_ID, SOURCE_TITLE, [{ title: 'Sheet1' }]);
    fake.setCells(SOURCE_ID, 'Sheet1', 'A1:C2', [
      ['Revenue', 'Costs', 'Margin'],
      [1000, 400, '=SUM(A2:B2)'],
    ]);
    fake.addSpreadsheet(DEST_ID, 'Destination', [{ title: 'Existing' }]);
    gsheet = new GoogleSheet(SOURCE_ID);
    await gsheet.authorize(fake.credentials);
  });

  const copyRequests = (): typeof fake.requests => fake.requests.filter((r) => r.url.includes(`/files/${SOURCE_ID}/copy`));
  const copyToRequests = (): typeof fake.requests => fake.requests.filter((r) => r.url.includes(':copyTo'));
  const exportRequests = (): typeof fake.requests => fake.requests.filter((r) => r.url.includes(`/drive/v3/files/${SOURCE_ID}/export`));

  describe('copySpreadsheet', () => {
    it('copies through the Drive files.copy endpoint and returns the created metadata', async () => {
      const result = await gsheet.copySpreadsheet({ title: 'My Copy' });

      const requests = copyRequests();
      expect(requests).to.have.length(1);
      expect(requests[0].method).to.equal('POST');
      expect(decodeURIComponent(requests[0].url)).to.contain(`/drive/v3/files/${SOURCE_ID}/copy`);
      expect(decodeURIComponent(requests[0].url)).to.contain('fields=id,name,mimeType,webViewLink');
      expect(requests[0].body).to.eql({ name: 'My Copy' });

      expect(result.sourceSpreadsheetId).to.equal(SOURCE_ID);
      expect(result.spreadsheetId).to.match(/^fake-copy-/);
      expect(result.title).to.equal('My Copy');
      expect(result.mimeType).to.equal('application/vnd.google-apps.spreadsheet');
      expect(result.webViewLink).to.equal(`https://docs.google.com/spreadsheets/d/${result.spreadsheetId}/edit`);
    });

    it('lets Drive inherit the source title when no title is asked for', async () => {
      const result = await gsheet.copySpreadsheet({});

      expect(copyRequests()[0].body).to.equal(undefined);
      expect(result.title).to.equal(SOURCE_TITLE);
    });

    it('carries data and formulas into the copy and leaves the source untouched', async () => {
      const result = await gsheet.copySpreadsheet({});

      expect(fake.cell(result.spreadsheetId, 'Sheet1', 'A1')).to.equal('Revenue');
      expect(fake.cell(result.spreadsheetId, 'Sheet1', 'C2')).to.equal('=SUM(A2:B2)');

      // the ordinary read path answers for the copy too
      const copied = await gsheet.getData({ worksheetTitle: 'Sheet1' }, result.spreadsheetId);
      expect(copied.rawData).to.eql([
        ['Revenue', 'Costs', 'Margin'],
        ['1000', '400', '=SUM(A2:B2)'],
      ]);

      // later edits to the source must not reach the copy
      fake.setCells(SOURCE_ID, 'Sheet1', 'A1', [['CHANGED']]);
      expect(fake.cell(result.spreadsheetId, 'Sheet1', 'A1')).to.equal('Revenue');
      expect(fake.cell(SOURCE_ID, 'Sheet1', 'C2')).to.equal('=SUM(A2:B2)');
    });

    it('rejects a blank title before asking Drive for anything', async () => {
      const before = fake.requests.length;

      try {
        await gsheet.copySpreadsheet({ title: '   ' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('title');
      }
      expect(fake.requests.length).to.equal(before);
    });

    it('fails visibly on an invisible file, with the drive.file explanation', async () => {
      try {
        await gsheet.copySpreadsheet({}, 'ghost-spreadsheet');
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('could not find the file');
        expect((error as Error).message).to.contain('drive.file');
      }
    });

    it('never replays the copy after a server failure', async () => {
      fake.failRequests('/copy', 500);

      try {
        await gsheet.copySpreadsheet({});
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Injected failure');
      }
      // one attempt, not a retry loop: a POST that may have succeeded must not be sent twice
      expect(copyRequests()).to.have.length(1);
    });

    it('keeps the HTTP status on a refused copy so failures can classify by status', async () => {
      fake.failRequests('/copy', 403);

      try {
        await gsheet.copySpreadsheet({});
        expect.fail('should have thrown');
      } catch (error) {
        // driveRequest wraps transport failures; the wrapper must carry the status forward
        expect(error).to.have.property('status', 403);
      }
      expect(copyRequests()).to.have.length(1);
    });

    it('reports 404 status metadata for an invisible file', async () => {
      try {
        await gsheet.copySpreadsheet({}, 'ghost-spreadsheet');
        expect.fail('should have thrown');
      } catch (error) {
        expect(error).to.have.property('status', 404);
      }
    });
  });

  describe('copyWorksheet', () => {
    it('copies through sheets.copyTo to the explicit destination and returns the new properties', async () => {
      const result = await gsheet.copyWorksheet({ worksheetTitle: 'Sheet1', destinationSpreadsheetId: DEST_ID });

      const requests = copyToRequests();
      expect(requests).to.have.length(1);
      expect(requests[0].method).to.equal('POST');
      expect(decodeURIComponent(requests[0].url)).to.contain(`/v4/spreadsheets/${SOURCE_ID}/sheets/`);
      expect(decodeURIComponent(requests[0].url)).to.contain(':copyTo');
      expect(requests[0].body).to.eql({ destinationSpreadsheetId: DEST_ID });

      expect(result.spreadsheetId).to.equal(SOURCE_ID);
      expect(result.worksheetTitle).to.equal('Sheet1');
      expect(result.destinationSpreadsheetId).to.equal(DEST_ID);
      expect(result.title).to.equal('Sheet1');
      expect(result.sheetId).to.be.a('number');
      expect(result.index).to.equal(1);
    });

    it('carries formulas into the destination worksheet', async () => {
      await gsheet.copyWorksheet({ worksheetTitle: 'Sheet1', destinationSpreadsheetId: DEST_ID });

      expect(fake.cell(DEST_ID, 'Sheet1', 'A1')).to.equal('Revenue');
      expect(fake.cell(DEST_ID, 'Sheet1', 'C2')).to.equal('=SUM(A2:B2)');

      const copied = await gsheet.getData({ worksheetTitle: 'Sheet1' }, DEST_ID);
      expect(copied.rawData[1]).to.eql(['1000', '400', '=SUM(A2:B2)']);
    });

    it('rejects an unknown source worksheet before any copyTo call', async () => {
      const before = copyToRequests().length;

      try {
        await gsheet.copyWorksheet({ worksheetTitle: 'Ghost', destinationSpreadsheetId: DEST_ID });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Ghost');
      }
      expect(copyToRequests()).to.have.length(before);
    });

    it('rejects an unknown destination spreadsheet and adds nothing anywhere', async () => {
      try {
        await gsheet.copyWorksheet({ worksheetTitle: 'Sheet1', destinationSpreadsheetId: 'ghost-destination' });
        expect.fail('should have thrown');
      } catch (error) {
        // copyTo is a Sheets call, so Google's own 404 text surfaces; driveRequest's drive.file
        // translation is only for the raw Drive transport
        expect((error as Error).message).to.contain('not found');
      }
      expect(fake.spreadsheets.has('ghost-destination')).to.be.false;
      expect(copyToRequests()).to.have.length(1);
    });

    it('assigns a unique title when the destination already uses the source title', async () => {
      const first = await gsheet.copyWorksheet({ worksheetTitle: 'Sheet1', destinationSpreadsheetId: DEST_ID });
      expect(first.title).to.equal('Sheet1');

      const second = await gsheet.copyWorksheet({ worksheetTitle: 'Sheet1', destinationSpreadsheetId: DEST_ID });
      // the server, not the caller, settles the title
      expect(second.title).to.equal('Sheet1 Copy');
      expect(second.index).to.equal(2);

      // every copy carries the data; the earlier copy and the source stay intact
      expect(fake.cell(DEST_ID, 'Sheet1 Copy', 'C2')).to.equal('=SUM(A2:B2)');
      expect(fake.cell(DEST_ID, 'Sheet1', 'C2')).to.equal('=SUM(A2:B2)');
      expect(fake.cell(SOURCE_ID, 'Sheet1', 'C2')).to.equal('=SUM(A2:B2)');

      const destination = fake.spreadsheets.get(DEST_ID);
      expect(destination?.sheets.map((s) => s.title)).to.eql(['Existing', 'Sheet1', 'Sheet1 Copy']);
    });

    it('duplicates a worksheet within the same spreadsheet under a server-assigned title', async () => {
      const result = await gsheet.copyWorksheet({ worksheetTitle: 'Sheet1', destinationSpreadsheetId: SOURCE_ID });

      expect(result.destinationSpreadsheetId).to.equal(SOURCE_ID);
      expect(result.sheetId).to.be.a('number');
      expect(result.title).to.equal('Sheet1 Copy');
      expect(fake.cell(SOURCE_ID, 'Sheet1 Copy', 'C2')).to.equal('=SUM(A2:B2)');

      const source = fake.spreadsheets.get(SOURCE_ID);
      expect(source?.sheets.map((s) => s.title)).to.eql(['Sheet1', 'Sheet1 Copy']);
    });
  });

  describe('exportSpreadsheet', () => {
    it('exports PDF through the Drive files.export endpoint with the exact bytes', async () => {
      const result = await gsheet.exportSpreadsheet({ format: 'pdf' });

      const requests = exportRequests();
      expect(requests).to.have.length(1);
      expect(requests[0].method).to.equal('GET');
      expect(decodeURIComponent(requests[0].url)).to.contain(`/drive/v3/files/${SOURCE_ID}/export?mimeType=application/pdf`);

      expect(result.spreadsheetId).to.equal(SOURCE_ID);
      expect(result.mimeType).to.equal('application/pdf');
      expect(Buffer.isBuffer(result.bytes)).to.be.true;
      expect(result.bytes.equals(fakeExportBytes(SOURCE_ID, 'application/pdf'))).to.be.true;
      expect(result.byteLength).to.equal(result.bytes.length);
    });

    it('maps the xlsx format to the OOXML spreadsheet MIME type', async () => {
      const result = await gsheet.exportSpreadsheet({ format: 'xlsx' });

      expect(decodeURIComponent(exportRequests()[0].url)).to.contain(`mimeType=${XLSX_MIME_TYPE}`);
      expect(result.mimeType).to.equal(XLSX_MIME_TYPE);
      expect(result.bytes.equals(fakeExportBytes(SOURCE_ID, XLSX_MIME_TYPE))).to.be.true;
    });

    it('rejects an unsupported format before asking Drive for anything', async () => {
      const before = fake.requests.length;
      // deliberate wrong input through the union type: the runtime check has to catch it
      const badFormat = 'csv' as unknown as GoogleSheetCli.ExportOptions['format'];

      try {
        await gsheet.exportSpreadsheet({ format: badFormat });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('pdf');
      }
      expect(fake.requests.length).to.equal(before);
    });

    it('fails visibly on an invisible file, with the drive.file explanation', async () => {
      try {
        await gsheet.exportSpreadsheet({ format: 'pdf' }, 'ghost-spreadsheet');
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('drive.file');
      }
    });

    it('keeps the HTTP status when Drive refuses the export', async () => {
      fake.failRequests('/export', 403, 5);

      try {
        await gsheet.exportSpreadsheet({ format: 'pdf' });
        expect.fail('should have thrown');
      } catch (error) {
        expect(error).to.have.property('status', 403);
      }
    });
  });

  describe('commands', () => {
    let tmpDir: string;
    let savedEnv: Record<string, string | undefined>;

    const MANAGED_ENV = ['GSHEET_CLIENT_EMAIL', 'GSHEET_PRIVATE_KEY', 'GSHEET_CREDENTIALS_FILE', 'GSHEET_USE_OAUTH', 'GSHEET_CLIENT_SECRET_FILE', 'SPREADSHEET_ID', 'WORKSHEET_TITLE'];

    const authFlags = [`--clientEmail=${fake.credentials.client_email}`];

    beforeEach(() => {
      // The key cannot travel on the command line (the PEM header has spaces), so it rides the
      // environment exactly the way test/commands/offline.test.ts manages its env.
      savedEnv = Object.fromEntries(MANAGED_ENV.map((name) => [name, process.env[name]]));
      for (const name of MANAGED_ENV) delete process.env[name];
      process.env.GSHEET_PRIVATE_KEY = fake.credentials.private_key;
      tmpDir = mkdtempSync(join(tmpdir(), 'gsheet-copy-export-'));
    });

    afterEach(() => {
      rmSync(tmpDir, { recursive: true, force: true });
      for (const [name, value] of Object.entries(savedEnv)) {
        if (value === undefined) delete process.env[name];
        else process.env[name] = value;
      }
    });

    describe('spreadsheet:copy', () => {
      it('returns the created metadata as JSON and really copies once', async () => {
        // runCommand re-splits array elements on spaces (splitString), so a spaced title must
        // carry its own embedded quotes or oclif sees a stray "Copy" positional
        const { error, result } = await runCommand<CopyCommandReceipt>(['spreadsheet:copy', `--spreadsheetId=${SOURCE_ID}`, '--title="CLI Copy"', ...authFlags, '--rawOutput']);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(result).to.include({ operation: 'spreadsheet:copy', sourceSpreadsheetId: SOURCE_ID, title: 'CLI Copy' });
        expect(copyRequests()).to.have.length(1);
      });

      it('speaks human output without --rawOutput', async () => {
        const { error, stdout } = await runCommand(['spreadsheet:copy', `--spreadsheetId=${SOURCE_ID}`, ...authFlags]);
        if (error) throw error;

        expect(stdout).to.contain(SOURCE_TITLE);
        expect(stdout).to.contain('fake-copy-');
        expect(stdout).to.not.contain('operation');
      });

      it('classifies a refused copy as FORBIDDEN in the failure envelope and sends it once', async () => {
        fake.failRequests('/copy', 403);

        const { error, stderr } = await runCommand(['spreadsheet:copy', `--spreadsheetId=${SOURCE_ID}`, ...authFlags, '--json']);
        expect(error).to.exist;
        const envelope = JSON.parse(stderr);
        expect(envelope.error.code).to.equal('FORBIDDEN');
        expect(envelope.error.retryable).to.equal(false);
        expect(copyRequests()).to.have.length(1);
      });

      it('classifies an invisible spreadsheet as NOT_FOUND in the failure envelope', async () => {
        const { error, stderr } = await runCommand(['spreadsheet:copy', '--spreadsheetId=ghost', ...authFlags, '--json']);
        expect(error).to.exist;
        const envelope = JSON.parse(stderr);
        expect(envelope.error.code).to.equal('NOT_FOUND');
        expect(envelope.error.retryable).to.equal(false);
      });
    });

    describe('worksheet:copy', () => {
      it('copies the worksheet to the destination and reports the new sheet', async () => {
        const { error, result } = await runCommand<WorksheetCopyReceipt>([
          'worksheet:copy',
          `--spreadsheetId=${SOURCE_ID}`,
          '--worksheetTitle=Sheet1',
          `--destinationSpreadsheetId=${DEST_ID}`,
          ...authFlags,
          '--rawOutput',
        ]);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(result).to.include({ operation: 'worksheet:copy', destinationSpreadsheetId: DEST_ID, worksheetTitle: 'Sheet1', title: 'Sheet1' });
        expect(result.sheetId).to.be.a('number');
        expect(fake.cell(DEST_ID, 'Sheet1', 'C2')).to.equal('=SUM(A2:B2)');
      });
    });

    describe('spreadsheet:export', () => {
      const exportArgs = (output: string, extra: string[] = []): string[] => [
        'spreadsheet:export',
        `--spreadsheetId=${SOURCE_ID}`,
        '--format=pdf',
        `--output=${output}`,
        ...authFlags,
        ...extra,
      ];

      it('writes the exact bytes to disk and reports a binary-free receipt', async () => {
        const output = join(tmpDir, 'report.pdf');
        const { error, result, stdout } = await runCommand<ExportCommandReceipt>([...exportArgs(output), '--rawOutput']);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(readFileSync(output).equals(fakeExportBytes(SOURCE_ID, 'application/pdf'))).to.be.true;
        expect(result.savedPath).to.equal(output);
        expect(result.bytesWritten).to.equal(256);
        expect(result.sha256Hash).to.be.a('string').with.length(64);
        expect(result.isOverwritten).to.be.false;

        // the receipt must parse as JSON and must not carry the bytes themselves
        const parsed = JSON.parse(stdout);
        expect(parsed).to.not.have.property('bytes');
        expect(parsed.mimeType).to.equal('application/pdf');
      });

      it('refuses an existing output file before touching the network, and leaves it alone', async () => {
        const output = join(tmpDir, 'report.pdf');
        writeFileSync(output, 'PRECIOUS', 'utf8');
        const before = exportRequests().length;

        const { error } = await runCommand(exportArgs(output));
        expect(error).to.be.an('error');
        expect((error as Error).message).to.contain('--overwrite');
        expect(readFileSync(output, 'utf8')).to.equal('PRECIOUS');
        expect(exportRequests()).to.have.length(before);
      });

      it('replaces an existing file only on explicit --overwrite', async () => {
        const output = join(tmpDir, 'report.pdf');
        writeFileSync(output, 'PRECIOUS', 'utf8');

        const { error, result } = await runCommand<ExportCommandReceipt>([...exportArgs(output), '--overwrite', '--rawOutput']);
        if (error || !result) throw error ?? new Error('the command resolved without a receipt');

        expect(result.isOverwritten).to.be.true;
        expect(readFileSync(output).equals(fakeExportBytes(SOURCE_ID, 'application/pdf'))).to.be.true;
      });

      it('protects an existing file when the export itself fails', async () => {
        const output = join(tmpDir, 'report.pdf');
        writeFileSync(output, 'PRECIOUS', 'utf8');
        // a GET is idempotent, so gaxios retries it; every attempt has to fail
        fake.failRequests('/export', 500, 10);

        const { error } = await runCommand([...exportArgs(output), '--overwrite']);
        expect(error).to.be.an('error');
        expect(readFileSync(output, 'utf8')).to.equal('PRECIOUS');
      });

      it('fails visibly on an invisible file and writes nothing', async () => {
        const output = join(tmpDir, 'report.pdf');

        const { error } = await runCommand(exportArgs(output).map((arg) => arg.replace(`--spreadsheetId=${SOURCE_ID}`, '--spreadsheetId=ghost')));
        expect(error).to.be.an('error');
        expect((error as Error).message).to.contain('drive.file');
        expect(existsSync(output)).to.be.false;
      });

      it('speaks human output without --rawOutput', async () => {
        const output = join(tmpDir, 'report.pdf');
        const { error, stdout } = await runCommand(exportArgs(output));
        if (error) throw error;

        expect(stdout).to.contain(`Exported spreadsheet ${SOURCE_ID} to ${output}`);
        expect(stdout).to.contain('application/pdf');
      });

      it('classifies a refused export as FORBIDDEN and leaves no output file', async () => {
        const output = join(tmpDir, 'report.pdf');
        fake.failRequests('/export', 403, 5);

        const { error, stderr } = await runCommand([...exportArgs(output, ['--json'])]);
        expect(error).to.exist;
        const envelope = JSON.parse(stderr);
        expect(envelope.error.code).to.equal('FORBIDDEN');
        expect(existsSync(output)).to.be.false;
      });
    });
  });
});
