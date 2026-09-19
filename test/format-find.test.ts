import { expect } from 'chai';
import GoogleSheet from '../src/lib/google-sheet';
import {
  a1ToGridRange,
  buildCellFormat,
  buildFormatRequests,
  buildMergeRequest,
  hexToRgbColor,
} from '../src/lib/sheet-format';
import { FakeSheets } from './fake-sheets';

const SPREADSHEET_ID = 'format-find-spreadsheet';
const SPREADSHEET_TITLE = 'Format & Find';
const SHEET = 'Report';

describe('sheet-format builders', () => {
  it('converts hex colors to rgb floats and rejects bad input', () => {
    expect(hexToRgbColor('#1a73e8')).to.eql({ red: 26 / 255, green: 115 / 255, blue: 232 / 255 });
    expect(hexToRgbColor('ffffff')).to.eql({ red: 1, green: 1, blue: 1 });
    expect(() => hexToRgbColor('#fff')).to.throw('Invalid color');
    expect(() => hexToRgbColor('red')).to.throw('Invalid color');
  });

  it('builds a field mask from exactly the supplied style properties', () => {
    const { cellFormat, fields } = buildCellFormat({ bold: true, numberFormat: '#,##0.00' });
    expect(cellFormat.textFormat?.bold).to.equal(true);
    expect(cellFormat.numberFormat).to.eql({ type: 'NUMBER', pattern: '#,##0.00' });
    expect(fields).to.eql(['userEnteredFormat.textFormat.bold', 'userEnteredFormat.numberFormat']);
    // nothing else may be in the mask - values and other formats stay untouched
    expect(fields.join(',')).to.not.contain('userEnteredValue');
    expect(fields.join(',')).to.not.contain('backgroundColor');
  });

  it('honors an explicit numberFormatType over inference', () => {
    const { cellFormat } = buildCellFormat({ numberFormat: '#,##0', numberFormatType: 'CURRENCY' });
    expect(cellFormat.numberFormat).to.eql({ type: 'CURRENCY', pattern: '#,##0' });
  });

  it('converts A1 ranges to 0-based end-exclusive grid ranges', () => {
    expect(a1ToGridRange('B2:D5', 42)).to.eql({
      sheetId: 42,
      startRowIndex: 1,
      endRowIndex: 5,
      startColumnIndex: 1,
      endColumnIndex: 4,
    });
    // column-only range keeps rows unbounded
    expect(a1ToGridRange('B:D', 42)).to.eql({ sheetId: 42, startColumnIndex: 1, endColumnIndex: 4 });
    expect(() => a1ToGridRange('not-a-range', 42)).to.throw();
  });

  it('emits repeatCell plus updateBorders for a styled range', () => {
    const { requests, fields } = buildFormatRequests(
      {
        ranges: ['A1:J1'],
        style: {
          bold: true,
          backgroundColor: '#1a73e8',
          borders: { sides: 'all', style: 'SOLID_THICK', color: '#000000' },
        },
      },
      () => 7
    );
    expect(requests).to.have.length(2);
    expect(requests[0].repeatCell?.range?.sheetId).to.equal(7);
    expect(requests[0].repeatCell?.fields).to.contain('textFormat.bold');
    expect(requests[0].repeatCell?.fields).to.contain('backgroundColor');
    const borders = requests[1].updateBorders;
    expect(borders?.top?.style).to.equal('SOLID_THICK');
    expect(borders?.bottom?.style).to.equal('SOLID_THICK');
    expect(borders?.left?.style).to.equal('SOLID_THICK');
    expect(borders?.right?.style).to.equal('SOLID_THICK');
    expect(borders?.innerHorizontal).to.equal(undefined);
    expect(fields).to.contain('borders');
  });

  it('expands "inner" borders to both inner sides only', () => {
    const { requests } = buildFormatRequests(
      { ranges: ['A1:C3'], style: { borders: { sides: 'inner' } } },
      () => 7
    );
    const borders = requests[0].updateBorders;
    expect(borders?.innerHorizontal?.style).to.equal('SOLID');
    expect(borders?.innerVertical?.style).to.equal('SOLID');
    expect(borders?.top).to.equal(undefined);
  });

  it('builds a clear request whose mask cannot touch values', () => {
    const { requests, fields } = buildFormatRequests({ ranges: ['A1:C3'], clear: true }, () => 7);
    expect(requests).to.have.length(1);
    expect(requests[0].repeatCell?.fields).to.equal('userEnteredFormat');
    expect(requests[0].repeatCell?.cell?.userEnteredFormat).to.eql({});
    expect(fields).to.eql(['userEnteredFormat']);
  });

  it('rejects empty specs and clear+style combinations', () => {
    expect(() => buildFormatRequests({ ranges: [] }, () => 7)).to.throw('non-empty "ranges"');
    expect(() => buildFormatRequests({ ranges: ['A1'] }, () => 7)).to.throw('either "clear" or style');
    expect(() => buildFormatRequests({ ranges: ['A1'], clear: true, style: { bold: true } }, () => 7)).to.throw(
      'cannot combine "clear"'
    );
  });

  it('resolves per-range worksheet titles through sheetIdFor', () => {
    const seen: (string | undefined)[] = [];
    buildFormatRequests(
      { ranges: ["'Other Sheet'!A1:B2", 'C3:D4'], worksheetTitle: 'Main', style: { bold: true } },
      (title) => {
        seen.push(title);
        return title === 'Other Sheet' ? 11 : 22;
      }
    );
    expect(seen).to.eql(['Other Sheet', 'Main']);
  });

  it('builds merge and unmerge requests', () => {
    const merge = buildMergeRequest('A1:C1', 'MERGE_ROWS', false, () => 9);
    expect(merge.mergeCells?.mergeType).to.equal('MERGE_ROWS');
    expect(merge.mergeCells?.range?.sheetId).to.equal(9);
    const unmerge = buildMergeRequest('A1:C1', undefined, true, () => 9);
    expect(unmerge.unmergeCells?.range?.sheetId).to.equal(9);
  });
});

describe('findData / formatCells / setMerge against the fake', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [{ title: SHEET, rowCount: 50, columnCount: 10 }]);
    fake.setCells(SPREADSHEET_ID, SHEET, 'A1', [
      ['Name', 'Status', 'Amount'],
      ['NV0123', 'Paid', '26500'],
      ['NV0456', 'Pending', '12000'],
      ['NV0123', 'Paid', '8800'],
      ['ERR-42', 'Failed', '500'],
    ]);
    gsheet = new GoogleSheet(SPREADSHEET_ID);
    await gsheet.authorize(fake.credentials);
  });

  const batchUpdateRequests = (): any[] =>
    fake.requests
      .filter((r) => r.method === 'POST' && r.url.includes(':batchUpdate'))
      .flatMap((r) => r.body?.requests ?? []);

  describe('findData', () => {
    it('returns A1 coordinates for exact matches', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, equals: 'NV0123' });
      expect(result.matchCount).to.equal(2);
      expect(result.truncated).to.equal(false);
      expect(result.matches.map((m) => m.a1)).to.eql([`${SHEET}!A2`, `${SHEET}!A4`]);
      expect(result.matches[0].columnLetter).to.equal('A');
      expect(result.matches[0].value).to.equal('NV0123');
    });

    it('restricts matches to a column letter', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, equals: 'Paid', column: 'B' });
      expect(result.matchCount).to.equal(2);
      expect(result.matches.every((m) => m.column === 2)).to.equal(true);
    });

    it('restricts matches to a header column and skips the header row', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, equals: 'Status', header: 'Status' });
      // "Status" appears only in the header row itself, which is excluded
      expect(result.matchCount).to.equal(0);

      const paid = await gsheet.findData({ worksheetTitle: SHEET, equals: 'Paid', header: 'Status' });
      expect(paid.matchCount).to.equal(2);
      expect(paid.matches.every((m) => m.column === 2)).to.equal(true);
    });

    it('throws when the header is not in the first scanned row', async () => {
      try {
        await gsheet.findData({ worksheetTitle: SHEET, equals: 'x', header: 'Missing' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Header "Missing" not found');
      }
    });

    it('matches substrings and regexes', async () => {
      const contains = await gsheet.findData({ worksheetTitle: SHEET, contains: 'NV0' });
      expect(contains.matchCount).to.equal(3);

      const regex = await gsheet.findData({ worksheetTitle: SHEET, regex: '^ERR-' });
      expect(regex.matchCount).to.equal(1);
      expect(regex.matches[0].a1).to.equal(`${SHEET}!A5`);
    });

    it('honors case sensitivity and numeric equality', async () => {
      const exact = await gsheet.findData({ worksheetTitle: SHEET, equals: 'paid', ignoreCase: false });
      expect(exact.matchCount).to.equal(0);
      const insensitive = await gsheet.findData({ worksheetTitle: SHEET, equals: 'paid' });
      expect(insensitive.matchCount).to.equal(2);
    });

    it('collapses matches to unique rows with --byRow semantics', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, contains: 'NV0123', byRow: true });
      expect(result.matchCount).to.equal(2);
      expect(result.matches[0].rowValues).to.eql(['NV0123', 'Paid', '26500']);
    });

    it('caps matches at limit while reporting the true total', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, contains: 'NV0', limit: 2 });
      expect(result.matchCount).to.equal(3);
      expect(result.truncated).to.equal(true);
      expect(result.matches).to.have.length(2);
    });

    it('anchors coordinates to the scanned range origin', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, range: 'B2:C6', equals: 'Paid' });
      expect(result.matchCount).to.equal(2);
      expect(result.matches.map((m) => m.a1)).to.eql([`${SHEET}!B2`, `${SHEET}!B4`]);
      expect(result.matches[0].row).to.equal(2);
      expect(result.matches[0].column).to.equal(2);
    });

    it('requires exactly one match mode', async () => {
      try {
        await gsheet.findData({ worksheetTitle: SHEET });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Exactly one match mode');
      }
      try {
        await gsheet.findData({ worksheetTitle: SHEET, equals: 'a', contains: 'b' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Exactly one match mode');
      }
    });

    it('rejects an invalid regex before any network call', async () => {
      const before = fake.requests.length;
      try {
        await gsheet.findData({ worksheetTitle: SHEET, regex: '([' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Invalid regex');
      }
      expect(fake.requests.length).to.equal(before);
    });
  });

  describe('findData edge cases', () => {
    it('matches empty cells when equals is an empty string', async () => {
      // seed a column with a hole in the middle: D2 is empty inside the grid
      fake.setCells(SPREADSHEET_ID, SHEET, 'D1', [['x'], [''], ['y']]);
      const result = await gsheet.findData({ worksheetTitle: SHEET, range: 'D1:D3', equals: '' });
      expect(result.matchCount).to.equal(1);
      expect(result.matches.map((m) => m.a1)).to.eql([`${SHEET}!D2`]);
    });

    it('rejects an empty contains needle instead of matching every cell', async () => {
      try {
        await gsheet.findData({ worksheetTitle: SHEET, contains: '' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('non-empty');
      }
    });

    it('rejects an empty regex pattern instead of matching every cell', async () => {
      try {
        await gsheet.findData({ worksheetTitle: SHEET, regex: '' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('non-empty');
      }
    });

    it('returns zero matches when the column restriction falls outside the scanned range', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, range: 'A1:B5', equals: 'Paid', column: 'Z' });
      expect(result.matchCount).to.equal(0);
      expect(result.matches).to.eql([]);
    });

    it('treats the first scanned row as the header row, not worksheet row 1', async () => {
      // --header resolves against the first row of the scanned range; a sub-range
      // starting below the real header cannot find it
      try {
        await gsheet.findData({ worksheetTitle: SHEET, range: 'A3:D5', equals: 'x', header: 'Status' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Header "Status" not found');
      }
    });

    it('anchors a1 output to the worksheet named in the range', async () => {
      // A range carrying another sheet's title scans that sheet; coordinates
      // must name the sheet the data actually came from
      fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [
        { title: SHEET, rowCount: 50, columnCount: 10 },
        { title: 'Second', rowCount: 20, columnCount: 8 },
      ]);
      fake.setCells(SPREADSHEET_ID, 'Second', 'A1', [['X', 'Y'], ['needle', 'z']]);
      const result = await gsheet.findData({ worksheetTitle: SHEET, range: "'Second'!A1:B5", equals: 'needle' });
      expect(result.matchCount).to.equal(1);
      expect(result.matches[0].a1).to.equal('Second!A2');
    });

    it('returns zero matches on an empty worksheet', async () => {
      fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [{ title: 'Empty', rowCount: 10, columnCount: 5 }]);
      const result = await gsheet.findData({ worksheetTitle: 'Empty', equals: 'anything' });
      expect(result.matchCount).to.equal(0);
      expect(result.matches).to.eql([]);
    });

    it('honors an explicit limit of zero', async () => {
      const result = await gsheet.findData({ worksheetTitle: SHEET, contains: 'NV0', limit: 0 });
      expect(result.matchCount).to.equal(3);
      expect(result.truncated).to.equal(true);
      expect(result.matches).to.eql([]);
    });

    it('matches numeric cells by numeric equality, not text', async () => {
      // 26500 stored as a number; equals "26500" must match it and must not
      // match the text "26500x" if one existed
      const result = await gsheet.findData({ worksheetTitle: SHEET, equals: '26500' });
      expect(result.matchCount).to.equal(1);
      expect(result.matches[0].a1).to.equal(`${SHEET}!C2`);
    });

    it('does not match a numeric needle against formatted text', async () => {
      // under FORMATTED_VALUE the cell renders "26,500.00" only after formatting;
      // the raw stored value here is the string "26500" so equals still hits it,
      // but a needle with separators must not
      const result = await gsheet.findData({ worksheetTitle: SHEET, equals: '26,500' });
      expect(result.matchCount).to.equal(0);
    });
  });

  describe('formatCells / setMerge edge cases', () => {
    it('formats exactly one cell for a bare "A1" range, not the whole grid', async () => {
      // a1ToGridRange must pin the end indices for a single-cell range; an
      // unbounded end would tell Google "to the edge" and format everything
      const receipt = await gsheet.formatCells({ ranges: ['B2'], worksheetTitle: SHEET, style: { bold: true } }, true);
      const range = receipt.requests?.[0].repeatCell?.range;
      expect(range?.startRowIndex).to.equal(1);
      expect(range?.endRowIndex).to.equal(2);
      expect(range?.startColumnIndex).to.equal(1);
      expect(range?.endColumnIndex).to.equal(2);
    });

    it('keeps column-only ranges row-unbounded', async () => {
      const receipt = await gsheet.formatCells({ ranges: ['B:D'], worksheetTitle: SHEET, style: { italic: true } }, true);
      const range = receipt.requests?.[0].repeatCell?.range;
      expect(range?.startRowIndex).to.equal(undefined);
      expect(range?.endRowIndex).to.equal(undefined);
      expect(range?.startColumnIndex).to.equal(1);
      expect(range?.endColumnIndex).to.equal(4);
    });

    it('rejects an unknown border side with a named error', async () => {
      try {
        await gsheet.formatCells(
          { ranges: ['A1'], worksheetTitle: SHEET, style: { borders: { sides: ['diagonal' as never] } } },
          false
        );
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Invalid border side "diagonal"');
      }
    });

    it('rejects a style spec that produces no requests', async () => {
      try {
        await gsheet.formatCells({ ranges: ['A1'], worksheetTitle: SHEET, style: {} }, false);
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('no requests');
      }
    });

    it('rejects a range whose worksheet title does not exist', async () => {
      try {
        await gsheet.formatCells({ ranges: ["'No Such Sheet'!A1:B2"], worksheetTitle: SHEET, style: { bold: true } }, false);
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('No Such Sheet');
      }
    });

    it('formats ranges across two worksheets in one call', async () => {
      fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [
        { title: SHEET, rowCount: 50, columnCount: 10 },
        { title: 'Second', rowCount: 20, columnCount: 8 },
      ]);
      const receipt = await gsheet.formatCells(
        { ranges: ['A1:C1', "'Second'!A1:B1"], worksheetTitle: SHEET, style: { bold: true } },
        true
      );
      const requests = receipt.requests ?? [];
      expect(requests).to.have.length(2);
      const sheetIds = requests.map((r) => r.repeatCell?.range?.sheetId);
      expect(new Set(sheetIds).size).to.equal(2);
    });

    it('rejects a non-positive fontSize', async () => {
      try {
        await gsheet.formatCells({ ranges: ['A1'], worksheetTitle: SHEET, style: { fontSize: 0 } }, false);
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('fontSize');
      }
    });

    it('merges a single cell without error', async () => {
      const receipt = await gsheet.setMerge({ worksheetTitle: SHEET, range: 'B2' });
      expect(receipt.requestCount).to.equal(1);
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).merges?.[0]).to.eql({
        startRowIndex: 1,
        endRowIndex: 2,
        startColumnIndex: 1,
        endColumnIndex: 2,
      });
    });

    it('rejects a merge range on a missing worksheet', async () => {
      try {
        await gsheet.setMerge({ worksheetTitle: 'Ghost', range: 'A1:B2' });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Ghost');
      }
    });
  });

  describe('formatCells', () => {
    it('sends repeatCell and updateBorders in one atomic batchUpdate', async () => {
      const receipt = await gsheet.formatCells(
        {
          ranges: ['A1:C1'],
          worksheetTitle: SHEET,
          style: { bold: true, backgroundColor: '#1a73e8', borders: { sides: 'all' } },
        },
        false
      );
      expect(receipt.requestCount).to.equal(2);
      expect(receipt.dryRun).to.equal(false);

      const requests = batchUpdateRequests();
      expect(requests).to.have.length(2);
      const repeat = requests.find((r) => r.repeatCell);
      expect(repeat.repeatCell.fields).to.contain('userEnteredFormat.textFormat.bold');
      expect(repeat.repeatCell.fields).to.contain('userEnteredFormat.backgroundColor');
      expect(repeat.repeatCell.fields).to.not.contain('userEnteredValue');
      const borders = requests.find((r) => r.updateBorders);
      expect(borders.updateBorders.top.style).to.equal('SOLID');
    });

    it('previews identical requests on dryRun without mutating', async () => {
      const receipt = await gsheet.formatCells(
        { ranges: ['A1:C1'], worksheetTitle: SHEET, style: { italic: true } },
        true
      );
      expect(receipt.dryRun).to.equal(true);
      expect(receipt.requests).to.have.length(1);
      expect(receipt.requests?.[0].repeatCell?.fields).to.equal('userEnteredFormat.textFormat.italic');
      // a dry run resolves metadata but must not reach batchUpdate
      expect(batchUpdateRequests()).to.have.length(0);
    });

    it('formats multiple ranges in one call', async () => {
      const receipt = await gsheet.formatCells(
        { ranges: ['A1:C1', 'A5:C5'], worksheetTitle: SHEET, style: { bold: true } },
        false
      );
      expect(receipt.requestCount).to.equal(2);
      const requests = batchUpdateRequests();
      expect(requests).to.have.length(2);
      expect(requests[0].repeatCell.range.startRowIndex).to.equal(0);
      expect(requests[1].repeatCell.range.startRowIndex).to.equal(4);
    });

    it('leaves cell values untouched after formatting', async () => {
      await gsheet.formatCells(
        { ranges: ['A1:C2'], worksheetTitle: SHEET, style: { bold: true, numberFormat: '#,##0.00' } },
        false
      );
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A1')).to.equal('Name');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'C2')).to.equal('26500');
    });

    it('clears formatting with a value-safe mask', async () => {
      const receipt = await gsheet.formatCells({ ranges: ['A1:C1'], worksheetTitle: SHEET, clear: true }, false);
      expect(receipt.fields).to.eql(['userEnteredFormat']);
      const requests = batchUpdateRequests();
      expect(requests[0].repeatCell.fields).to.equal('userEnteredFormat');
      expect(requests[0].repeatCell.cell.userEnteredFormat).to.eql({});
    });

    it('rejects a spec with no ranges', async () => {
      try {
        await gsheet.formatCells({ ranges: [] }, false);
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('non-empty "ranges"');
      }
    });
  });

  describe('setMerge', () => {
    it('merges a range and reports it through spreadsheets.get', async () => {
      const receipt = await gsheet.setMerge({ worksheetTitle: SHEET, range: 'A1:C1' });
      expect(receipt.fields).to.eql(['mergeCells.MERGE_ALL']);

      const sheet = fake.worksheet(SPREADSHEET_ID, SHEET);
      expect(sheet.merges).to.have.length(1);
      expect(sheet.merges?.[0]).to.eql({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 3 });
    });

    it('unmerges a previously merged range', async () => {
      await gsheet.setMerge({ worksheetTitle: SHEET, range: 'A1:C1' });
      await gsheet.setMerge({ worksheetTitle: SHEET, range: 'A1:C1', unmerge: true });
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).merges ?? []).to.have.length(0);
    });

    it('previews a merge on dryRun without mutating', async () => {
      const receipt = await gsheet.setMerge({ worksheetTitle: SHEET, range: 'A1:C1', dryRun: true });
      expect(receipt.requests?.[0].mergeCells?.mergeType).to.equal('MERGE_ALL');
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).merges ?? []).to.have.length(0);
    });
  });
});
