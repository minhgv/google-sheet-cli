import { expect } from 'chai';
import GoogleSheet, { GoogleSheetCli } from '../src/lib/google-sheet';
import { FakeSheets } from './fake-sheets';

/**
 * Behavioral contract for `GoogleSheet.clearData` (T-05 / AC-04) and `GoogleSheet.upsert`
 * (T-06 / AC-05), driven through the in-memory fake of the Sheets REST API the same way
 * `test/regression.test.ts` drives it.
 *
 * The risk focus is the destructive edge: a clear or upsert that is refused - by the formula
 * guard, a duplicate key, an ambiguous header or a failed request - must leave the sheet
 * byte-for-byte as it was, and a repeated upsert under single-writer conditions must not
 * append duplicates.
 */

const SPREADSHEET_ID = 'fake-spreadsheet-id';
const SPREADSHEET_TITLE = 'Regression spreadsheet';
const TITLE = 'Sheet1';

/**
 * Run a call that is expected to reject and hand back whatever it threw
 *
 * @param {() => Promise<any>} fn
 * @returns {Promise<any>}
 */
const rejection = async (fn: () => Promise<any>): Promise<any> => {
  try {
    await fn();
  } catch (error) {
    return error;
  }
  throw new Error('expected the call to reject, but it resolved');
};

describe('clearData and upsert', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  /**
   * How many values-write POSTs the fake has seen so far - both the values:batchUpdate
   * writes and the single-range values/:clear POSTs; a dry run or a refused operation must
   * never move this counter.
   *
   * @returns {number}
   */
  const writePosts = (): number =>
    fake.requests.filter(
      (request) =>
        request.method === 'POST' &&
        (/\/values:batchUpdate/.test(request.url) || /\/values\/.+(?::|%3A)clear/.test(request.url))
    ).length;

  /**
   * The default table: header row plus two data rows. The fake stores everything as strings
   * and re-types on read, exactly as RAW storage does: '001' reads back a string, 7 a number.
   */
  const seedTable = (): void => {
    fake.setCells(SPREADSHEET_ID, TITLE, 'A1', [
      ['id', 'name', 'score'],
      ['001', 'alpha', 1],
      [7, 'bravo', 2],
    ]);
  };

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [{ title: TITLE }]);
    gsheet = new GoogleSheet(SPREADSHEET_ID);
    await gsheet.authorize(fake.credentials);
  });

  describe('clearData', () => {
    it('clears values in a bounded range and leaves every cell outside it', async () => {
      seedTable();
      const receipt = await gsheet.clearData({ range: 'Sheet1!B2:C3', worksheetTitle: TITLE });

      expect(receipt.dryRun).to.equal(false);
      expect(receipt.rowsCleared).to.equal(2);
      expect(receipt.columnsCleared).to.equal(2);
      expect(receipt.cellsCleared).to.equal(4);
      expect(receipt.range).to.equal('Sheet1!B2:C3');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C3')).to.equal('');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A1')).to.equal('id');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A2')).to.equal('001');
      expect(writePosts()).to.equal(1);
    });

    it('clears a single cell when the range is one cell', async () => {
      seedTable();
      const receipt = await gsheet.clearData({ range: 'Sheet1!C2', worksheetTitle: TITLE });

      expect(receipt.cellsCleared).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('alpha');
    });

    it('clears through one real values.clear request, not a staged empty matrix', async () => {
      seedTable();
      const batchUpdatesBefore = fake.requests.filter(
        (request) => request.method === 'POST' && request.url.includes('/values:batchUpdate')
      ).length;

      const receipt = await gsheet.clearData({ range: 'Sheet1!B2:C3', worksheetTitle: TITLE });

      const clearPosts = fake.requests.filter(
        (request) => request.method === 'POST' && /\/values\/.+(?::|%3A)clear/.test(request.url)
      );
      expect(clearPosts).to.have.lengthOf(1);
      expect(clearPosts[0].url).to.contain('/values/');
      // the clear no longer routes through values:batchUpdate at all
      const batchUpdatesAfter = fake.requests.filter(
        (request) => request.method === 'POST' && request.url.includes('/values:batchUpdate')
      ).length;
      expect(batchUpdatesAfter).to.equal(batchUpdatesBefore);
      expect(receipt.batchesExecuted).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('');
    });

    it('refuses to clear a formula without overwriteFormulas', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C4', [['=SUM(A2:B3)']]);
      const error = await rejection(() => gsheet.clearData({ range: 'Sheet1!A1:C4', worksheetTitle: TITLE }));

      expect(error.message).to.match(/overwriteFormulas/);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C4')).to.equal('=SUM(A2:B3)');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('alpha');
      expect(writePosts()).to.equal(0);
    });

    it('clears formulas when overwriteFormulas is set', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C4', [['=SUM(A2:B3)']]);
      const receipt = await gsheet.clearData({ range: 'Sheet1!C4', worksheetTitle: TITLE, overwriteFormulas: true });

      expect(receipt.cellsCleared).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C4')).to.equal('');
    });

    it('dry-run reports the plan and formula conflicts without mutating', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C4', [['=SUM(A2:B3)']]);
      const receipt = await gsheet.clearData({ range: 'Sheet1!A1:C4', worksheetTitle: TITLE, dryRun: true });

      expect(receipt.dryRun).to.equal(true);
      expect(receipt.rowsCleared).to.equal(4);
      expect(receipt.cellsCleared).to.equal(12);
      expect(receipt.formulasOverwritten).to.have.lengthOf(1);
      expect(receipt.formulasOverwritten?.[0]).to.contain('C4');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C4')).to.equal('=SUM(A2:B3)');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('alpha');
      expect(writePosts()).to.equal(0);
    });

    it('rejects unbounded ranges: whole columns, whole rows, bare worksheet names', async () => {
      await rejection(() => gsheet.clearData({ range: 'Sheet1!A:C', worksheetTitle: TITLE }));
      await rejection(() => gsheet.clearData({ range: 'Sheet1!1:3', worksheetTitle: TITLE }));
      await rejection(() => gsheet.clearData({ range: 'Sheet1', worksheetTitle: TITLE }));
      expect(writePosts()).to.equal(0);
    });

    it('rejects a range whose worksheet contradicts the worksheetTitle flag', async () => {
      const error = await rejection(() => gsheet.clearData({ range: 'Other!A1:B2', worksheetTitle: TITLE }));
      expect(error.message).to.match(/Conflicting worksheet/);
    });

    it('rejects a worksheet that does not exist before writing', async () => {
      const error = await rejection(() => gsheet.clearData({ range: 'A1:B2', worksheetTitle: 'Nope' }));
      expect(error.message).to.match(/not found/);
      expect(writePosts()).to.equal(0);
    });

    it('rejects a range that reaches past the grid', async () => {
      const error = await rejection(() => gsheet.clearData({ range: 'Sheet1!A1:Z2000', worksheetTitle: TITLE }));
      expect(error.message).to.match(/past the worksheet grid/);
      expect(writePosts()).to.equal(0);
    });
  });

  describe('upsert', () => {
    it('adds new keys and updates matched keys without touching unsupplied columns', async () => {
      seedTable();
      const receipt = await gsheet.upsert([['id', 'score'], ['001', 99], [42, 'delta']], {
        key: 'id',
        worksheetTitle: TITLE,
      });

      expect(receipt.existingRows).to.equal(2);
      expect(receipt.rowsUpdated).to.equal(1);
      expect(receipt.rowsAdded).to.equal(1);
      expect(receipt.rowsUnchanged).to.equal(0);
      expect(receipt.updatedRanges).to.eql(['Sheet1!C2']);
      expect(receipt.addedRanges).to.eql(['Sheet1!A4:C4']);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('99');
      // the matched row's unsupplied column and its key cell stay exactly as they were
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('alpha');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A2')).to.equal('001');
      // the appended row lands after the last existing data row
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('42');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C4')).to.equal('delta');
      expect(writePosts()).to.equal(1);
    });

    it('lists plannedRanges in actual execution order for mixed new and existing rows', async () => {
      seedTable();
      const receipt = await gsheet.upsert(
        [['id', 'score'], ['001', 99], [42, 'delta'], [7, 22]],
        { key: 'id', worksheetTitle: TITLE }
      );

      // input order drives execution: update for "001", append for the unseen 42, update for 7
      expect(receipt.plannedRanges).to.eql(['Sheet1!C2', 'Sheet1!A4:C4', 'Sheet1!C3']);
      expect(receipt.updatedRanges).to.eql(['Sheet1!C2', 'Sheet1!C3']);
      expect(receipt.addedRanges).to.eql(['Sheet1!A4:C4']);
    });

    it('treats cells a short input row omits as not supplied, never as clears', async () => {
      seedTable();
      // the row for key 7 carries only the key: its name and score must survive untouched
      const receipt = await gsheet.upsert([['id', 'name', 'score'], [7]], { key: 'id', worksheetTitle: TITLE });

      expect(receipt.rowsUpdated).to.equal(0);
      expect(receipt.rowsUnchanged).to.equal(1);
      expect(receipt.plannedRanges).to.eql([]);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B3')).to.equal('bravo');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C3')).to.equal('2');
      expect(writePosts()).to.equal(0);
    });

    it('is stable on repeat invocation: the same input again writes nothing and appends nothing', async () => {
      seedTable();
      const input: GoogleSheetCli.RawData = [['id', 'score'], ['001', 99], [42, 'delta']];
      await gsheet.upsert(input, { key: 'id', worksheetTitle: TITLE });
      const postsAfterFirst = writePosts();

      const second = await gsheet.upsert(input, { key: 'id', worksheetTitle: TITLE });

      expect(second.rowsAdded).to.equal(0);
      expect(second.rowsUpdated).to.equal(0);
      expect(second.rowsUnchanged).to.equal(2);
      expect(second.plannedRanges).to.eql([]);
      expect(writePosts()).to.equal(postsAfterFirst);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('42');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A5')).to.equal('');
    });

    it('keeps the string key "001" distinct from the number 1 and stable across reruns', async () => {
      seedTable();

      // a string key matching the existing string cell is a match, not a new row
      const match = await gsheet.upsert([['id'], ['001']], { key: 'id', worksheetTitle: TITLE });
      expect(match.rowsUnchanged).to.equal(1);
      expect(match.rowsAdded).to.equal(0);

      // the number 1 is a different key from the string "001": it appends
      const added = await gsheet.upsert([['id'], [1]], { key: 'id', worksheetTitle: TITLE });
      expect(added.rowsAdded).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('1');

      // rerunning the same numeric input matches the row it just wrote instead of appending again
      const repeat = await gsheet.upsert([['id'], [1]], { key: 'id', worksheetTitle: TITLE });
      expect(repeat.rowsAdded).to.equal(0);
      expect(repeat.rowsUnchanged).to.equal(1);

      // and the string "1" is again a different key from the number 1
      const stringOne = await gsheet.upsert([['id'], ['1']], { key: 'id', worksheetTitle: TITLE });
      expect(stringOne.rowsAdded).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A5')).to.equal('1');
    });

    it('refuses duplicate keys in the input before any write', async () => {
      seedTable();
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], [7, 5], [7, 6]], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Duplicate key 7 in the input at rows 2 and 3/);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C3')).to.equal('2');
      expect(writePosts()).to.equal(0);
    });

    it('refuses an empty key in the input', async () => {
      seedTable();
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['', 5]], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Empty key in the input at row 2/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses duplicate keys that already exist in the table', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'A4', [['dup', 'x', 9]]);
      fake.setCells(SPREADSHEET_ID, TITLE, 'A5', [['dup', 'y', 8]]);
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['dup', 1]], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Duplicate key "dup" in "Sheet1" at rows 4 and 5/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses an empty key that already exists in the table', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'A4', [['', 'ghost', 3]]);
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['001', 5]], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Empty key in "Sheet1" at row 4/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses unknown input columns', async () => {
      seedTable();
      const error = await rejection(() =>
        gsheet.upsert([['id', 'bogus'], [7, 1]], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Unknown input column "bogus"/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses non-scalar input cells', async () => {
      seedTable();
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], [7, ['nested']]] as unknown as GoogleSheetCli.RawData, {
          key: 'id',
          worksheetTitle: TITLE,
        })
      );

      expect(error.message).to.match(/Upsert cells must be scalars/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses a duplicate header in the existing table', async () => {
      fake.addSpreadsheet('ss-dup-header', 'Dup header', [{ title: 'T' }]);
      fake.setCells('ss-dup-header', 'T', 'A1', [['id', 'name', 'id']]);
      const error = await rejection(() =>
        gsheet.upsert([['id'], ['x']], { key: 'id', worksheetTitle: 'T' }, 'ss-dup-header')
      );

      expect(error.message).to.match(/Duplicate header "id" in "T" \(columns A and C\)/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses an empty header inside the existing table span', async () => {
      fake.addSpreadsheet('ss-gap-header', 'Gap header', [{ title: 'T' }]);
      fake.setCells('ss-gap-header', 'T', 'A1', [['id', '', 'score']]);
      const error = await rejection(() =>
        gsheet.upsert([['id'], ['x']], { key: 'id', worksheetTitle: 'T' }, 'ss-gap-header')
      );

      expect(error.message).to.match(/Empty header at B1 in "T"/);
    });

    it('refuses a key column the table does not have, and input missing the key column', async () => {
      seedTable();
      const noTableKey = await rejection(() => gsheet.upsert([['zip'], ['x']], { key: 'zip', worksheetTitle: TITLE }));
      expect(noTableKey.message).to.match(/Key column "zip" not found in the existing table header \[id, name, score\]/);

      const noInputKey = await rejection(() => gsheet.upsert([['name'], ['x']], { key: 'id', worksheetTitle: TITLE }));
      expect(noInputKey.message).to.match(/does not supply the key column "id"/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses duplicate or empty headers in the input', async () => {
      seedTable();
      const dup = await rejection(() =>
        gsheet.upsert([['id', 'id'], ['a', 'b']], { key: 'id', worksheetTitle: TITLE })
      );
      expect(dup.message).to.match(/Duplicate header "id" in the input/);

      const empty = await rejection(() =>
        gsheet.upsert([['id', ''], ['a', 'b']], { key: 'id', worksheetTitle: TITLE })
      );
      expect(empty.message).to.match(/Empty header in the input at position 2/);
      expect(writePosts()).to.equal(0);
    });

    it('honors an offset table range and writes at absolute coordinates', async () => {
      fake.addSpreadsheet('ss-offset', 'Offset', [{ title: 'T' }]);
      fake.setCells('ss-offset', 'T', 'A1', [['noise']]);
      fake.setCells('ss-offset', 'T', 'B2', [
        ['id', 'name'],
        ['x', 1],
      ]);

      const receipt = await gsheet.upsert(
        [
          ['id', 'name'],
          ['x', 'y'],
          ['z', 'w'],
        ],
        { key: 'id', range: 'T!B2:C9', worksheetTitle: 'T' },
        'ss-offset'
      );

      expect(receipt.worksheetTitle).to.equal('T');
      expect(receipt.updatedRanges).to.eql(['T!C3']);
      expect(receipt.addedRanges).to.eql(['T!B4:C4']);
      expect(fake.cell('ss-offset', 'T', 'C3')).to.equal('y');
      // column A is outside the table: untouched
      expect(fake.cell('ss-offset', 'T', 'A1')).to.equal('noise');
      expect(fake.cell('ss-offset', 'T', 'B4')).to.equal('z');
      expect(fake.cell('ss-offset', 'T', 'C4')).to.equal('w');
    });

    it('dry-run reports counts, planned ranges and formula conflicts without writing', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C2', [['=SUM(A2:B2)']]);
      const receipt = await gsheet.upsert(
        [['id', 'score'], ['001', 5], [99, 'new']],
        { key: 'id', worksheetTitle: TITLE, dryRun: true }
      );

      expect(receipt.dryRun).to.equal(true);
      expect(receipt.rowsUpdated).to.equal(1);
      expect(receipt.rowsAdded).to.equal(1);
      expect(receipt.updatedRanges).to.eql(['Sheet1!C2']);
      expect(receipt.addedRanges).to.eql(['Sheet1!A4:C4']);
      expect(receipt.formulasOverwritten).to.have.lengthOf(1);
      expect(receipt.formulasOverwritten?.[0]).to.contain('C2');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('=SUM(A2:B2)');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('');
      expect(writePosts()).to.equal(0);
    });

    it('refuses to overwrite a formula in a changed cell before any write', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C2', [['=SUM(A2:B2)']]);
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['001', 5]], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Cannot overwrite existing formula\(s\) without overwriteFormulas=true/);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('=SUM(A2:B2)');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('alpha');
      expect(writePosts()).to.equal(0);
    });

    it('overwrites formulas in changed cells only with overwriteFormulas', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C2', [['=SUM(A2:B2)']]);
      const receipt = await gsheet.upsert([['id', 'score'], ['001', 5]], {
        key: 'id',
        worksheetTitle: TITLE,
        overwriteFormulas: true,
      });

      expect(receipt.rowsUpdated).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('5');
      // untouched cells, formulas elsewhere included, are never part of any write
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B2')).to.equal('alpha');
    });

    it('preserves a formula when the supplied value equals the current one', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'C2', [['=SUM(A2:B2)']]);
      // in the fake a formula cell stores the formula text, so "the same value" here is the
      // identical formula string - which must count as unchanged, not as a write target
      const receipt = await gsheet.upsert([['id', 'score'], ['001', '=SUM(A2:B2)']], {
        key: 'id',
        worksheetTitle: TITLE,
      });

      expect(receipt.rowsUnchanged).to.equal(1);
      expect(receipt.plannedRanges).to.eql([]);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('=SUM(A2:B2)');
      expect(writePosts()).to.equal(0);
    });

    it('grows the grid when appended rows pass the worksheet bounds', async () => {
      fake.addSpreadsheet('ss-small', 'Small', [{ title: 'T', rowCount: 5 }]);
      fake.setCells('ss-small', 'T', 'A1', [['id'], ['a'], ['b']]);

      const receipt = await gsheet.upsert(
        [['id'], ['c'], ['d'], ['e'], ['f'], ['g']],
        { key: 'id', worksheetTitle: 'T' },
        'ss-small'
      );

      expect(receipt.rowsAdded).to.equal(5);
      expect(fake.cell('ss-small', 'T', 'A7')).to.equal('f');
      expect(fake.cell('ss-small', 'T', 'A8')).to.equal('g');
    });

    it('leaves no partial state when the write fails and never replays it internally', async () => {
      seedTable();
      fake.failRequests('values:batchUpdate', 500, 1);
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['001', 5], [42, 'delta']], { key: 'id', worksheetTitle: TITLE })
      );

      expect(error.message).to.match(/Batch update failed at batch 1\/1/);
      // nothing was written by the failed attempt
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('1');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('');

      // the retry is the caller's decision, and it applies exactly once - no duplicates
      const rerun = await gsheet.upsert([['id', 'score'], ['001', 5], [42, 'delta']], {
        key: 'id',
        worksheetTitle: TITLE,
      });
      expect(rerun.rowsUpdated).to.equal(1);
      expect(rerun.rowsAdded).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('5');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('42');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A5')).to.equal('');
    });

    it('treats an empty input as a no-op without any API call', async () => {
      seedTable();
      const requestsBefore = fake.requests.length;
      const receipt = await gsheet.upsert([], { key: 'id', worksheetTitle: TITLE });

      expect(receipt.rowsAdded).to.equal(0);
      expect(receipt.rowsUpdated).to.equal(0);
      expect(receipt.rowsUnchanged).to.equal(0);
      expect(receipt.plannedRanges).to.eql([]);
      expect(fake.requests.length).to.equal(requestsBefore);
    });

    it('treats a header-only input as a validated no-op', async () => {
      seedTable();
      const receipt = await gsheet.upsert([['id', 'name', 'score']], { key: 'id', worksheetTitle: TITLE });

      expect(receipt.existingRows).to.equal(2);
      expect(receipt.rowsAdded).to.equal(0);
      expect(receipt.rowsUpdated).to.equal(0);
      expect(receipt.rowsUnchanged).to.equal(0);
      expect(writePosts()).to.equal(0);
    });

    it('refuses a worksheet that does not exist', async () => {
      seedTable();
      const error = await rejection(() => gsheet.upsert([['id'], ['x']], { key: 'id', worksheetTitle: 'Nope' }));
      expect(error.message).to.match(/not found/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses to upsert into a worksheet without any table header', async () => {
      const error = await rejection(() => gsheet.upsert([['id'], ['x']], { key: 'id', worksheetTitle: TITLE }));
      expect(error.message).to.match(/No table header found at A1 in "Sheet1"/);
      expect(writePosts()).to.equal(0);
    });

    it('refuses a bounded range with existing rows below its end bound, matching rows included', async () => {
      seedTable();
      // the declared range stops after the first data row; the row for key 7 is never read,
      // so an append for the unseen key 42 would land on it and destroy it
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['001', 99], [42, 'delta']], {
          key: 'id',
          range: 'Sheet1!A1:C2',
          worksheetTitle: TITLE,
        })
      );

      expect(error.message).to.match(/stops at row 2/);
      expect(error.message).to.match(/A3/);
      expect(error.message).to.match(/nothing was written/);
      expect(error.code).to.equal('VALIDATION');
      // the matching row is not updated either: the refusal is all-or-nothing
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C2')).to.equal('1');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B3')).to.equal('bravo');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C3')).to.equal('2');
      expect(writePosts()).to.equal(0);
    });

    it('refuses a bounded range when a duplicate key sits farther below the bound', async () => {
      seedTable();
      fake.setCells(SPREADSHEET_ID, TITLE, 'A6', [['dup', 'echo', 9]]);
      // "dup" is unseen inside the bound, so it would append - duplicating the key hiding
      // three rows farther down; the below-bound scan must catch it before any write
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['dup', 1]], {
          key: 'id',
          range: 'Sheet1!A1:C3',
          worksheetTitle: TITLE,
        })
      );

      expect(error.message).to.match(/stops at row 3/);
      expect(error.message).to.match(/A6/);
      expect(error.code).to.equal('VALIDATION');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A6')).to.equal('dup');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'C6')).to.equal('9');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('');
      expect(writePosts()).to.equal(0);
    });

    it('refuses a bounded range below the bound even on dry run', async () => {
      seedTable();
      const error = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['001', 99], [42, 'delta']], {
          key: 'id',
          range: 'Sheet1!A1:C2',
          worksheetTitle: TITLE,
          dryRun: true,
        })
      );

      expect(error.message).to.match(/stops at row 2/);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B3')).to.equal('bravo');
      expect(writePosts()).to.equal(0);
    });

    it('allows a bounded range that reaches the last data row when nothing lies below it', async () => {
      seedTable();
      // the bound sits exactly at the last data row and the rows below it are empty, so the
      // same safety preflight passes and the append proceeds
      const receipt = await gsheet.upsert([['id', 'score'], ['42', 'delta']], {
        key: 'id',
        range: 'Sheet1!A1:C3',
        worksheetTitle: TITLE,
      });

      expect(receipt.rowsAdded).to.equal(1);
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'A4')).to.equal('42');
      expect(fake.cell(SPREADSHEET_ID, TITLE, 'B3')).to.equal('bravo');
      expect(writePosts()).to.equal(1);
    });

    it('marks local clear and upsert refusals with the structural VALIDATION code', async () => {
      seedTable();
      const unbounded = await rejection(() => gsheet.clearData({ range: 'Sheet1!A:C', worksheetTitle: TITLE }));
      expect(unbounded.code).to.equal('VALIDATION');

      const dupKey = await rejection(() =>
        gsheet.upsert([['id', 'score'], [7, 5], [7, 6]], { key: 'id', worksheetTitle: TITLE })
      );
      expect(dupKey.code).to.equal('VALIDATION');
      expect(dupKey.message).to.match(/Duplicate key 7 in the input at rows 2 and 3/);

      // a transport failure is not an input problem: it stays a plain error without the code
      fake.failRequests('values:batchUpdate', 500, 1);
      const transport = await rejection(() =>
        gsheet.upsert([['id', 'score'], ['001', 5]], { key: 'id', worksheetTitle: TITLE })
      );
      expect(transport.message).to.match(/Batch update failed at batch 1\/1/);
      expect(transport.code).to.equal(undefined);
    });
  });
});
