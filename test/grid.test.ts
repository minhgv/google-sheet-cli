import { expect } from 'chai';
import GoogleSheet from '../src/lib/google-sheet';
import { FakeSheets } from './fake-sheets';

const SPREADSHEET_ID = 'grid-spreadsheet';
const SPREADSHEET_TITLE = 'Grid Ops';
const SHEET = 'Data';

/**
 * Structural dimension mutations against the fake. insertDimension/deleteDimension
 * actually shift the fake's cell map, so a post-op getData proves the move rather
 * than just the request shape.
 */
describe('grid dimension mutations', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [{ title: SHEET, rowCount: 10, columnCount: 5 }]);
    fake.setCells(SPREADSHEET_ID, SHEET, 'A1', [
      ['r1a', 'r1b', 'r1c'],
      ['r2a', 'r2b', 'r2c'],
      ['r3a', 'r3b', 'r3c'],
      ['r4a', 'r4b', 'r4c'],
      ['r5a', 'r5b', 'r5c'],
    ]);
    gsheet = new GoogleSheet(SPREADSHEET_ID);
    await gsheet.authorize(fake.credentials);
  });

  const batchUpdateRequests = (): any[] =>
    fake.requests
      .filter((r) => r.method === 'POST' && r.url.includes(':batchUpdate'))
      .flatMap((r) => r.body?.requests ?? []);

  describe('insert', () => {
    it('shifts rows down and grows the grid', async () => {
      const receipt = await gsheet.mutateDimension('insert', {
        worksheetTitle: SHEET,
        dimension: 'ROWS',
        start: 3,
        count: 2,
      });
      expect(receipt.requestCount).to.equal(1);

      const sheet = fake.worksheet(SPREADSHEET_ID, SHEET);
      expect(sheet.rowCount).to.equal(12);
      // rows 1-2 untouched, rows 3-5 now at 5-7
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A2')).to.equal('r2a');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A3')).to.equal('');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A5')).to.equal('r3a');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A7')).to.equal('r5a');
    });

    it('shifts columns right', async () => {
      await gsheet.mutateDimension('insert', { worksheetTitle: SHEET, dimension: 'COLUMNS', start: 2, count: 1 });
      const sheet = fake.worksheet(SPREADSHEET_ID, SHEET);
      expect(sheet.columnCount).to.equal(6);
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A1')).to.equal('r1a');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'B1')).to.equal('');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'C1')).to.equal('r1b');
    });

    it('sends inheritFromBefore through', async () => {
      await gsheet.mutateDimension('insert', {
        worksheetTitle: SHEET,
        dimension: 'ROWS',
        start: 3,
        count: 1,
        inheritFromBefore: true,
      });
      const requests = batchUpdateRequests();
      expect(requests[0].insertDimension.inheritFromBefore).to.equal(true);
    });

    it('previews on dryRun without mutating', async () => {
      const receipt = await gsheet.mutateDimension('insert', {
        worksheetTitle: SHEET,
        dimension: 'ROWS',
        start: 3,
        count: 2,
        dryRun: true,
      });
      expect(receipt.dryRun).to.equal(true);
      expect(receipt.requests?.[0].insertDimension?.range?.startIndex).to.equal(2);
      expect(receipt.requests?.[0].insertDimension?.range?.endIndex).to.equal(4);
      expect(batchUpdateRequests()).to.have.length(0);
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A3')).to.equal('r3a');
    });
  });

  describe('delete', () => {
    it('removes rows and shifts the rest up', async () => {
      await gsheet.mutateDimension('delete', { worksheetTitle: SHEET, dimension: 'ROWS', start: 2, count: 2 });
      const sheet = fake.worksheet(SPREADSHEET_ID, SHEET);
      expect(sheet.rowCount).to.equal(8);
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A1')).to.equal('r1a');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A2')).to.equal('r4a');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A3')).to.equal('r5a');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A4')).to.equal('');
    });

    it('removes columns and shifts the rest left', async () => {
      await gsheet.mutateDimension('delete', { worksheetTitle: SHEET, dimension: 'COLUMNS', start: 2, count: 1 });
      const sheet = fake.worksheet(SPREADSHEET_ID, SHEET);
      expect(sheet.columnCount).to.equal(4);
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'B1')).to.equal('r1c');
    });

    it('previews the values about to be removed on dryRun', async () => {
      const receipt = await gsheet.mutateDimension('delete', {
        worksheetTitle: SHEET,
        dimension: 'ROWS',
        start: 2,
        count: 2,
        dryRun: true,
      });
      expect(receipt.dryRun).to.equal(true);
      expect(receipt.affectedValues).to.eql([
        ['r2a', 'r2b', 'r2c'],
        ['r3a', 'r3b', 'r3c'],
      ]);
      expect(batchUpdateRequests()).to.have.length(0);
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A2')).to.equal('r2a');
    });
  });

  describe('hide', () => {
    it('marks rows hiddenByUser', async () => {
      await gsheet.mutateDimension('hide', { worksheetTitle: SHEET, dimension: 'ROWS', start: 2, count: 2 });
      const props = fake.worksheet(SPREADSHEET_ID, SHEET).dimensionProps?.rows;
      expect(props?.[2]?.hiddenByUser).to.equal(true);
      expect(props?.[3]?.hiddenByUser).to.equal(true);
      expect(props?.[1]).to.equal(undefined);
    });

    it('unhides with the unhide flag', async () => {
      await gsheet.mutateDimension('hide', { worksheetTitle: SHEET, dimension: 'ROWS', start: 2, count: 1 });
      await gsheet.mutateDimension('hide', { worksheetTitle: SHEET, dimension: 'ROWS', start: 2, count: 1, unhide: true });
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).dimensionProps?.rows?.[2]?.hiddenByUser).to.equal(false);
    });

    it('hides columns', async () => {
      await gsheet.mutateDimension('hide', { worksheetTitle: SHEET, dimension: 'COLUMNS', start: 3, count: 1 });
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).dimensionProps?.columns?.[3]?.hiddenByUser).to.equal(true);
    });
  });

  describe('resize', () => {
    it('sets an explicit pixel size', async () => {
      await gsheet.mutateDimension('resize', { worksheetTitle: SHEET, dimension: 'COLUMNS', start: 1, count: 2, pixels: 120 });
      const props = fake.worksheet(SPREADSHEET_ID, SHEET).dimensionProps?.columns;
      expect(props?.[1]?.pixelSize).to.equal(120);
      expect(props?.[2]?.pixelSize).to.equal(120);
    });

    it('sends autoResizeDimensions for auto mode', async () => {
      await gsheet.mutateDimension('resize', { worksheetTitle: SHEET, dimension: 'COLUMNS', start: 1, count: 3, autoResize: true });
      const requests = batchUpdateRequests();
      expect(requests[0].autoResizeDimensions.dimensions.startIndex).to.equal(0);
      expect(requests[0].autoResizeDimensions.dimensions.endIndex).to.equal(3);
    });

    it('rejects pixels and auto together', async () => {
      try {
        await gsheet.mutateDimension('resize', {
          worksheetTitle: SHEET,
          dimension: 'COLUMNS',
          start: 1,
          pixels: 100,
          autoResize: true,
        });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('not both');
      }
    });

    it('rejects a resize with neither pixels nor auto', async () => {
      try {
        await gsheet.mutateDimension('resize', { worksheetTitle: SHEET, dimension: 'COLUMNS', start: 1 });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('either "pixels" or "autoResize"');
      }
    });
  });

  describe('validation', () => {
    it('rejects a start below 1', async () => {
      try {
        await gsheet.mutateDimension('insert', { worksheetTitle: SHEET, dimension: 'ROWS', start: 0 });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('start >= 1');
      }
    });

    it('rejects a count below 1', async () => {
      try {
        await gsheet.mutateDimension('delete', { worksheetTitle: SHEET, dimension: 'ROWS', start: 1, count: 0 });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('count >= 1');
      }
    });

    it('rejects an invalid dimension', async () => {
      try {
        await gsheet.mutateDimension('insert', { worksheetTitle: SHEET, dimension: 'DIAGONAL' as never, start: 1 });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('ROWS" or "COLUMNS');
      }
    });

    it('rejects a missing worksheet', async () => {
      try {
        await gsheet.mutateDimension('insert', { worksheetTitle: 'Ghost', dimension: 'ROWS', start: 1 });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('Ghost');
      }
    });
  });

  describe('setFrozen', () => {
    it('freezes rows and columns', async () => {
      const receipt = await gsheet.setFrozen({ worksheetTitle: SHEET, rows: 1, columns: 2 });
      expect(receipt.fields).to.eql(['gridProperties.frozenRowCount', 'gridProperties.frozenColumnCount']);
      const sheet = fake.worksheet(SPREADSHEET_ID, SHEET);
      expect(sheet.frozenRowCount).to.equal(1);
      expect(sheet.frozenColumnCount).to.equal(2);
    });

    it('unfreezes with 0', async () => {
      await gsheet.setFrozen({ worksheetTitle: SHEET, rows: 1 });
      await gsheet.setFrozen({ worksheetTitle: SHEET, rows: 0 });
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).frozenRowCount).to.equal(0);
    });

    it('requires at least one axis', async () => {
      try {
        await gsheet.setFrozen({ worksheetTitle: SHEET });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('at least one of "rows" or "columns"');
      }
    });

    it('rejects a negative count', async () => {
      try {
        await gsheet.setFrozen({ worksheetTitle: SHEET, rows: -1 });
        expect.fail('should have thrown');
      } catch (error) {
        expect((error as Error).message).to.contain('>= 0');
      }
    });

    it('previews on dryRun without mutating', async () => {
      const receipt = await gsheet.setFrozen({ worksheetTitle: SHEET, rows: 2, dryRun: true });
      expect(receipt.requests?.[0].updateSheetProperties?.properties?.gridProperties?.frozenRowCount).to.equal(2);
      expect(fake.worksheet(SPREADSHEET_ID, SHEET).frozenRowCount).to.equal(undefined);
    });
  });
});
