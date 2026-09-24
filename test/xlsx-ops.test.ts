import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { expect } from 'chai';
import { XlsxWorkbook } from '../src/lib/xlsx';

/**
 * Offline coverage for the local-backend structural operations added by the
 * XLSX multi-tool roadmap (waves 1-2): find, inspect extensions, splice with
 * merge management, merge/unmerge, sparse cell writes, and bounded clear.
 * Every case runs against real workbooks built with XlsxWorkbook - no network.
 */
describe('XlsxWorkbook structural operations', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gsheet-xlsx-ops-'));
  });

  afterEach(async () => {
    if (tmpDir) {
      await fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    }
  });

  const buildSheet = (wb: XlsxWorkbook, name = 'Data') => {
    const ws = wb.workbook.addWorksheet(name);
    ws.getCell('A1').value = 'Name';
    ws.getCell('B1').value = 'Amount';
    ws.getCell('A2').value = 'Alice';
    ws.getCell('B2').value = 100;
    ws.getCell('A3').value = 'Bob';
    ws.getCell('B3').value = 200;
    ws.getCell('A5').value = 'Total';
    ws.getCell('B5').value = { formula: 'SUM(B2:B4)', result: 300 };
    return ws;
  };

  describe('find', () => {
    it('locates cells by equals and returns data:find-shaped matches', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.find({ worksheetTitle: 'Data', equals: 'Alice' });

      expect(result.matchCount).to.equal(1);
      expect(result.truncated).to.be.false;
      expect(result.matches[0]).to.deep.include({
        a1: 'A2',
        row: 2,
        column: 1,
        columnLetter: 'A',
        value: 'Alice',
      });
      expect(result.range).to.contain('Data!');
    });

    it('supports contains, regex, column restriction and byRow', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(wb.find({ worksheetTitle: 'Data', contains: 'lic' }).matchCount).to.equal(1);
      expect(wb.find({ worksheetTitle: 'Data', regex: '^B' }).matchCount).to.equal(1); // "Bob"
      expect(wb.find({ worksheetTitle: 'Data', contains: 'o', column: 'A' }).matchCount).to.equal(2); // Bob and Total
      const byRow = wb.find({ worksheetTitle: 'Data', contains: 'o', column: 'A', byRow: true });
      expect(byRow.matches[0].rowValues).to.deep.equal(['Bob', 200]);
    });

    it('resolves --header against the first scanned row', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.find({ worksheetTitle: 'Data', contains: '0', header: 'Amount' });
      // only column B cells: 100, 200, 0? -> 100 and 200 contain '0'
      expect(result.matches.every((m) => m.columnLetter === 'B')).to.be.true;
      expect(result.matchCount).to.equal(3); // 100, 200 and the cached 300 in B5
    });

    it('requires exactly one match mode', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      expect(() => wb.find({ worksheetTitle: 'Data' })).to.throw('Exactly one match mode');
      expect(() => wb.find({ worksheetTitle: 'Data', equals: 'a', contains: 'b' })).to.throw('Exactly one match mode');
    });

    it('honors limit while matchCount reports the true total', () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Big');
      for (let r = 1; r <= 10; r++) ws.getCell(`A${r}`).value = 'hit';

      const result = wb.find({ worksheetTitle: 'Big', equals: 'hit', limit: 3 });
      expect(result.matchCount).to.equal(10);
      expect(result.truncated).to.be.true;
      expect(result.matches).to.have.length(3);
    });
  });

  describe('inspect extensions', () => {
    it('reports merged ranges, formula counts, frozen panes and validations', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.mergeCells('A5:A6');
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.getCell('C1').dataValidation = { type: 'list', formulae: ['"a,b,c"'] };

      const inspection = wb.inspect({ includeFormulaCells: true });
      const sheet = inspection.sheets[0];

      expect(sheet.mergedRanges).to.deep.equal(['A5:A6']);
      expect(sheet.hasFormulas).to.be.true;
      expect(sheet.formulaCellCount).to.equal(1);
      expect(sheet.formulaCells).to.deep.equal(['B5']);
      expect(sheet.frozen).to.deep.equal({ rows: 1, columns: 0 });
      expect(sheet.dataValidations).to.have.length(1);
    });
  });

  describe('splice', () => {
    it('inserts rows and shifts data down', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 3, count: 2 });

      expect(result.operation).to.equal('insert');
      expect(ws.getCell('A3').value).to.be.null;
      expect(ws.getCell('A5').value).to.equal('Bob');
      expect(ws.getCell('A7').value).to.equal('Total');
      // formula cell moved with its row; the formula text is NOT rewritten
      expect(ws.getCell('B7').formula).to.equal('SUM(B2:B4)');
      expect(result.formulasAtRisk).to.be.greaterThan(0);
      expect(result.warnings.some((w) => w.includes('not rewritten'))).to.be.true;
    });

    it('deletes rows and reports removed values', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.splice('delete', { worksheetTitle: 'Data', dimension: 'ROWS', start: 2, count: 2 });

      expect(ws.getCell('A2').value).to.be.null;
      expect(ws.getCell('A3').value).to.equal('Total');
      expect(result.removedValues).to.have.length(2);
      expect(result.removedValues![0]).to.deep.equal(['Alice', 100]);
    });

    it('dryRun mutates nothing', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 3, count: 2, dryRun: true });

      expect(ws.getCell('A3').value).to.equal('Bob');
    });

    it('shifts merges below the splice point on insert', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.mergeCells('A5:B5');

      const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 3, count: 2 });

      expect(result.mergesAdjusted).to.deep.equal([{ before: 'A5:B5', after: 'A7:B7' }]);
      expect(ws.model.merges).to.deep.equal(['A7:B7']);
    });

    it('refuses when a merge straddles the insertion boundary without force', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.mergeCells('A2:B4'); // straddles start=3

      expect(() =>
        wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 3, count: 1 })
      ).to.throw('Splice blocked');
      // untouched
      expect(ws.model.merges).to.deep.equal(['A2:B4']);
    });

    it('extends a straddling merge on insert with force', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.mergeCells('A2:B4');

      const result = wb.splice('insert', {
        worksheetTitle: 'Data',
        dimension: 'ROWS',
        start: 3,
        count: 2,
        force: true,
      });

      expect(result.mergesAdjusted).to.deep.equal([{ before: 'A2:B4', after: 'A2:B6' }]);
      expect(ws.model.merges).to.deep.equal(['A2:B6']);
    });

    it('shrinks a merge intersecting a deleted span with force', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.mergeCells('A2:B4'); // delete rows 3-4 -> merge becomes A2:B2

      const result = wb.splice('delete', {
        worksheetTitle: 'Data',
        dimension: 'ROWS',
        start: 3,
        count: 2,
        force: true,
      });

      expect(result.mergesAdjusted).to.deep.equal([{ before: 'A2:B4', after: 'A2:B2' }]);
      expect(ws.model.merges).to.deep.equal(['A2:B2']);
    });

    it('drops a merge fully inside the deleted span with force', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.mergeCells('A2:B2');

      const result = wb.splice('delete', {
        worksheetTitle: 'Data',
        dimension: 'ROWS',
        start: 2,
        count: 1,
        force: true,
      });

      expect(result.mergesAdjusted).to.deep.equal([{ before: 'A2:B2' }]);
      expect(ws.model.merges).to.deep.equal([]);
    });

    it('inserts columns and shifts data right', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      wb.splice('insert', { worksheetTitle: 'Data', dimension: 'COLUMNS', start: 2, count: 1 });

      expect(ws.getCell('A2').value).to.equal('Alice');
      expect(ws.getCell('C2').value).to.equal(100);
      expect(ws.getCell('B2').value).to.be.null;
    });
  });

  describe('mergeCellsRange', () => {
    it('merges a range and unmerges it back', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const merged = wb.mergeCellsRange({ worksheetTitle: 'Data', range: 'A6:B6' });
      expect(merged.affected).to.deep.equal(['A6:B6']);
      expect(ws.model.merges).to.deep.equal(['A6:B6']);

      const unmerged = wb.mergeCellsRange({ worksheetTitle: 'Data', range: 'A6:B6', unmerge: true });
      expect(unmerged.unmerge).to.be.true;
      expect(ws.model.merges).to.deep.equal([]);
    });

    it('decomposes MERGE_ROWS into one merge per row', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.mergeCellsRange({ worksheetTitle: 'Data', range: 'A6:B7', mergeType: 'MERGE_ROWS' });

      expect(result.affected).to.deep.equal(['A6:B6', 'A7:B7']);
      expect(ws.model.merges).to.have.length(2);
    });

    it('dryRun reports affected ranges without merging', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.mergeCellsRange({ worksheetTitle: 'Data', range: 'A6:B6', dryRun: true });
      expect(result.dryRun).to.be.true;
      expect(ws.model.merges).to.deep.equal([]);
    });

    it('persists merges through save and reload', async () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.mergeCellsRange({ worksheetTitle: 'Data', range: 'A6:B6' });

      const filePath = path.join(tmpDir, 'merged.xlsx');
      await wb.save(filePath);

      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.workbook.getWorksheet('Data')!.model.merges).to.deep.equal(['A6:B6']);
      expect(reloaded.inspect().sheets[0].mergedRanges).to.deep.equal(['A6:B6']);
    });
  });

  describe('setCells', () => {
    it('writes only the listed cells and leaves others untouched', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.setCells([
        { a1: 'A4', value: 'Carol' },
        { a1: 'B4', value: 300 },
      ]);

      expect(result.applied).to.equal(2);
      expect(ws.getCell('A4').value).to.equal('Carol');
      expect(ws.getCell('B4').value).to.equal(300);
      expect(ws.getCell('B5').formula).to.equal('SUM(B2:B4)'); // untouched formula
    });

    it('refuses to overwrite a formula cell without overwriteFormulas', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(() => wb.setCells([{ a1: 'B5', value: 999 }])).to.throw('formula overwrite conflict');
      // nothing written
      expect(wb.workbook.getWorksheet('Data')!.getCell('B5').formula).to.equal('SUM(B2:B4)');
    });

    it('allows formula overwrite with the consent flag and writes new formulas', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      wb.setCells([{ a1: 'B5', value: { formula: 'SUM(B2:B6)' } }], { overwriteFormulas: true });
      expect(ws.getCell('B5').formula).to.equal('SUM(B2:B6)');
    });

    it('dryRun diffs without writing', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.setCells([{ a1: 'A2', value: 'Changed' }], { dryRun: true });
      expect(result.applied).to.equal(0);
      expect(result.changes[0].action).to.equal('update');
      expect(ws.getCell('A2').value).to.equal('Alice');
    });
  });

  describe('clearRange', () => {
    it('clears values in a bounded range and keeps styles', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.getCell('A2').font = { bold: true };

      const result = wb.clearRange({ worksheetTitle: 'Data', range: 'A2:B3' });

      expect(result.cellsCleared).to.equal(4);
      expect(ws.getCell('A2').value).to.be.null;
      expect(ws.getCell('B3').value).to.be.null;
      expect(ws.getCell('A2').font?.bold).to.be.true;
      expect(ws.getCell('A1').value).to.equal('Name'); // outside range untouched
    });

    it('refuses to clear formula cells without overwriteFormulas', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(() => wb.clearRange({ worksheetTitle: 'Data', range: 'A5:B5' })).to.throw('formula cell');
      expect(wb.workbook.getWorksheet('Data')!.getCell('B5').formula).to.equal('SUM(B2:B4)');
    });

    it('clears formulas with the consent flag', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.clearRange({ worksheetTitle: 'Data', range: 'A5:B5', overwriteFormulas: true });
      expect(result.formulasOverwritten).to.deep.equal(['B5']);
      expect(ws.getCell('B5').value).to.be.null;
    });

    it('dryRun reports counts without clearing', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.clearRange({ worksheetTitle: 'Data', range: 'A2:B3', dryRun: true });
      expect(result.cellsCleared).to.equal(4);
      expect(ws.getCell('A2').value).to.equal('Alice');
    });
  });
});
