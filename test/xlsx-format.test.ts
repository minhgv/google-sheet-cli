import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { expect } from 'chai';
import { XlsxWorkbook } from '../src/lib/xlsx';
import { XlsxFormatCellsOptions } from '../src/lib/xlsx-types';

/**
 * Offline coverage for the Wave 4 local-backend operations: format:cells
 * styling, freeze/resize/hide presentation, worksheet add/remove/rename and
 * defined-name management. Every case runs against real workbooks built with
 * XlsxWorkbook - no network.
 */
describe('XlsxWorkbook presentation and sheet management', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gsheet-xlsx-format-'));
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
    return ws;
  };

  describe('formatCells', () => {
    it('applies font, fill, alignment, wrap and number format to the range', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.formatCells({
        worksheetTitle: 'Data',
        range: 'A1:B3',
        bold: true,
        textColor: '#ffffff',
        backgroundColor: '#1a73e8',
        horizontalAlignment: 'CENTER',
        verticalAlignment: 'MIDDLE',
        wrapText: true,
        numberFormat: '#,##0.00',
      });

      expect(result.sheet).to.equal('Data');
      expect(result.range).to.equal('A1:B3');
      expect(result.cellsFormatted).to.equal(6);
      expect(result.dryRun).to.equal(false);
      for (const address of ['A1', 'B2', 'B3']) {
        const cell = ws.getCell(address);
        expect(cell.font?.bold).to.equal(true);
        expect(cell.font?.color?.argb).to.equal('FFFFFFFF');
        expect((cell.fill as { fgColor?: { argb?: string } }).fgColor?.argb).to.equal('FF1A73E8');
        expect(cell.alignment?.horizontal).to.equal('center');
        expect(cell.alignment?.vertical).to.equal('middle');
        expect(cell.alignment?.wrapText).to.equal(true);
        expect(cell.numFmt).to.equal('#,##0.00');
      }
      expect(ws.getCell('C1').font?.bold).to.be.undefined;
    });

    it('persists styles through save and reload', async () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.formatCells({ worksheetTitle: 'Data', range: 'A1:B1', bold: true, numberFormat: '0.0%' });

      const filePath = path.join(tmpDir, 'styled.xlsx');
      await wb.save(filePath);

      const reloaded = await XlsxWorkbook.load(filePath);
      const ws = reloaded.workbook.getWorksheet('Data')!;
      expect(ws.getCell('A1').font?.bold).to.equal(true);
      expect(ws.getCell('A1').numFmt).to.equal('0.0%');
      expect(ws.getCell('A2').font?.bold).to.not.equal(true);
    });

    it('maps Sheets border styles and decomposes inner sides per cell', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.formatCells({
        worksheetTitle: 'Data',
        range: 'A1:B2',
        borders: 'inner',
        borderStyle: 'DOUBLE',
        borderColor: '#ff0000',
      });

      expect(result.applied).to.deep.equal(['borders(innerHorizontal,innerVertical)']);
      // 2x2 range: inner sides land between the cells, never on the outer edge
      expect(ws.getCell('A1').border?.bottom?.style).to.equal('double');
      expect(ws.getCell('A1').border?.top).to.be.undefined;
      expect(ws.getCell('B1').border?.bottom?.style).to.equal('double');
      expect(ws.getCell('A2').border?.top?.style).to.equal('double');
      expect(ws.getCell('A2').border?.bottom).to.be.undefined;
      expect(ws.getCell('B2').border?.left?.style).to.equal('double');
      expect(ws.getCell('B2').border?.right).to.be.undefined;
      expect(ws.getCell('A1').border?.bottom?.color?.argb).to.equal('FFFF0000');

      wb.formatCells({ worksheetTitle: 'Data', range: 'C1:C1', borders: 'all', borderStyle: 'SOLID_MEDIUM' });
      const c1 = ws.getCell('C1').border;
      expect(c1?.top?.style).to.equal('medium');
      expect(c1?.bottom?.style).to.equal('medium');
      expect(c1?.left?.style).to.equal('medium');
      expect(c1?.right?.style).to.equal('medium');
    });

    it('clear resets every style group including the number format', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      wb.formatCells({ worksheetTitle: 'Data', range: 'B2', bold: true, numberFormat: '#,##0.00' });

      const result = wb.formatCells({ worksheetTitle: 'Data', range: 'B2', clear: true });

      expect(result.applied).to.deep.equal(['clear']);
      expect(ws.getCell('B2').font?.bold).to.be.undefined;
      expect(ws.getCell('B2').numFmt).to.not.equal('#,##0.00');
      expect(ws.getCell('B2').value).to.equal(100);
    });

    it('dryRun reports the planned style without touching any cell', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.formatCells({
        worksheetTitle: 'Data',
        range: 'A1:B2',
        bold: true,
        italic: true,
        backgroundColor: '#eeeeee',
        dryRun: true,
      });

      expect(result.dryRun).to.equal(true);
      expect(result.cellsFormatted).to.equal(4);
      expect(result.applied).to.deep.equal(['bold', 'italic', 'backgroundColor']);
      expect(ws.getCell('A1').font?.bold).to.be.undefined;
      expect(ws.getCell('B2').fill).to.be.undefined;
    });

    it('refuses unbounded ranges, empty specs and clear combined with style flags', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(() => wb.formatCells({ worksheetTitle: 'Data', range: 'A:B', bold: true })).to.throw('bounded on both axes');
      expect(() => wb.formatCells({ worksheetTitle: 'Data', range: 'A1:B2' })).to.throw('No formatting flags');
      expect(() => wb.formatCells({ worksheetTitle: 'Data', range: 'A1:B2', clear: true, bold: true })).to.throw(
        'cannot be combined with style flags'
      );
    });

    it('refuses invalid colors and border sides instead of writing broken styles', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(() => wb.formatCells({ worksheetTitle: 'Data', range: 'A1', textColor: 'blue' })).to.throw('#RRGGBB');
      expect(() => wb.formatCells({ worksheetTitle: 'Data', range: 'A1', borders: 'diagonal' })).to.throw('Invalid border side');
      expect(() =>
        wb.formatCells({
          worksheetTitle: 'Data',
          range: 'A1',
          borders: 'all',
          // bypass the compile-time union on purpose: this exercises the runtime guard
          borderStyle: 'HAIRY',
        } as unknown as XlsxFormatCellsOptions)
      ).to.throw('borderStyle "HAIRY" is not supported');
    });
  });

  describe('freezePanes', () => {
    it('freezes rows and columns and persists the frozen pane', async () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.freezePanes({ worksheetTitle: 'Data', rows: 1, columns: 2 });

      expect(result.frozen).to.equal(true);
      expect(result.model).to.equal("views[0] = {state:'frozen', xSplit:2, ySplit:1}");

      const filePath = path.join(tmpDir, 'frozen.xlsx');
      await wb.save(filePath);
      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.inspect().sheets[0].frozen).to.deep.equal({ rows: 1, columns: 2 });
    });

    it('keeps the other axis and unrelated view attributes when only one axis is given', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      wb.freezePanes({ worksheetTitle: 'Data', rows: 1, columns: 2 });
      const prior = (ws.views ?? [])[0];
      if (prior) prior.zoomScale = 90;

      const result = wb.freezePanes({ worksheetTitle: 'Data', rows: 2 });

      expect(result.rows).to.equal(2);
      expect(result.columns).to.equal(2); // preserved from the prior frozen view, not coerced to 0
      // views is Partial<WorksheetView>[]; the split fields only exist on the
      // frozen/split variants, so asserting them needs the structural cast
      const view = (ws.views ?? [])[0] as { xSplit?: number; ySplit?: number; topLeftCell?: string; zoomScale?: number };
      expect(view.xSplit).to.equal(2);
      expect(view.ySplit).to.equal(2);
      expect(view.topLeftCell).to.equal('C3');
      expect(view.zoomScale).to.equal(90); // merged over, not replaced
    });

    it('unfreezes the pane when both axes are 0', async () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      wb.freezePanes({ worksheetTitle: 'Data', rows: 3 });
      wb.freezePanes({ worksheetTitle: 'Data', rows: 0, columns: 0 });

      const result = wb.freezePanes({ worksheetTitle: 'Data', rows: 0, columns: 0 });
      expect(result.frozen).to.equal(false);
      expect(result.model).to.equal("views[0] = {state:'normal'}");

      // unfreezing clears the pane's leftover split and topLeftCell attributes
      const view = (ws.views ?? [])[0] as { state?: string; xSplit?: number; ySplit?: number; topLeftCell?: string };
      expect(view.state).to.equal('normal');
      expect(view.xSplit).to.equal(0);
      expect(view.ySplit).to.equal(0);
      expect(view.topLeftCell).to.be.undefined;

      const filePath = path.join(tmpDir, 'unfrozen.xlsx');
      await wb.save(filePath);
      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.inspect().sheets[0].frozen).to.be.undefined;
    });

    it('dryRun reports the model change without freezing', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.freezePanes({ worksheetTitle: 'Data', rows: 2, dryRun: true });

      expect(result.dryRun).to.equal(true);
      expect(result.frozen).to.equal(true);
      expect((ws.views ?? []).some((v) => (v as { state?: string }).state === 'frozen')).to.equal(false);
    });

    it('requires at least one axis and refuses negative splits', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(() => wb.freezePanes({ worksheetTitle: 'Data' })).to.throw('requires rows or columns');
      expect(() => wb.freezePanes({ worksheetTitle: 'Data', rows: -1 })).to.throw('non-negative integer');
    });
  });

  describe('resizeGrid', () => {
    it('converts pixel sizes into Excel width characters and height points', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 1, count: 2, pixels: 75 });
      expect(result.unit).to.equal('width-chars');
      expect(result.resized).to.deep.equal([
        { index: 1, before: undefined, after: 10 },
        { index: 2, before: undefined, after: 10 },
      ]);
      expect(ws.getColumn(1).width).to.equal(10);

      const rows = wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'ROWS', start: 2, pixels: 40 });
      expect(rows.unit).to.equal('height-points');
      expect(rows.resized[0]).to.deep.equal({ index: 2, before: undefined, after: 30 });
      expect(ws.getRow(2).height).to.equal(30);
    });

    it('auto-sizes columns to their longest text and caps the width', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      ws.getCell('A2').value = 'A'.repeat(80);

      const result = wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 1, count: 2, auto: true });

      expect(result.resized[0].after).to.equal(60); // capped
      expect(result.resized[1].after).to.equal(8); // "Amount" is 6 chars + padding
      expect(ws.getColumn(1).width).to.equal(60);
    });

    it('auto-sizes from sparse rows far below actualRowCount', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);
      const sparseText = 'a long value living in a sparse row';
      ws.getCell('A1000').value = sparseText;

      const result = wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 1, auto: true });

      // actualRowCount is 3 (rows 1-3 hold data); row 1000 must still count
      expect(result.resized[0].after).to.equal(sparseText.length + 2);
      expect(ws.getColumn(1).width).to.equal(sparseText.length + 2);
    });

    it('dryRun reports every target without resizing', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.resizeGrid({
        worksheetTitle: 'Data',
        dimension: 'COLUMNS',
        start: 1,
        count: 2,
        pixels: 75,
        dryRun: true,
      });

      expect(result.dryRun).to.equal(true);
      expect(result.resized).to.have.lengthOf(2);
      expect(ws.getColumn(1).width).to.be.undefined;
      expect(ws.getColumn(2).width).to.be.undefined;
    });

    it('requires exactly one sizing mode and refuses impossible sizes', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      expect(() => wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 1 })).to.throw('requires either');
      expect(() => wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 1, pixels: 2, auto: true })).to.throw(
        'either pixels or auto'
      );
      expect(() => wb.resizeGrid({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 1, pixels: 3 })).to.throw('too small');
    });
  });

  describe('setGridHidden', () => {
    it('hides and unhides columns, persisting the flag', async () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.setGridHidden({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 2 });
      expect(result.hidden).to.equal(true);
      expect(result.model).to.equal('column 2: hidden = true');
      expect(ws.getColumn(2).hidden).to.equal(true);

      const filePath = path.join(tmpDir, 'hidden.xlsx');
      await wb.save(filePath);
      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.workbook.getWorksheet('Data')!.getColumn(2).hidden).to.equal(true);

      reloaded.setGridHidden({ worksheetTitle: 'Data', dimension: 'COLUMNS', start: 2, unhide: true });
      expect(reloaded.workbook.getWorksheet('Data')!.getColumn(2).hidden).to.equal(false);
    });

    it('dryRun reports the model change without hiding', () => {
      const wb = XlsxWorkbook.create();
      const ws = buildSheet(wb);

      const result = wb.setGridHidden({ worksheetTitle: 'Data', dimension: 'ROWS', start: 2, count: 2, dryRun: true });

      expect(result.dryRun).to.equal(true);
      expect(result.model).to.equal('row 2-3: hidden = true');
      expect(ws.getRow(2).hidden).to.not.equal(true);
    });
  });

  describe('worksheet add/remove/rename', () => {
    it('adds a sheet and refuses duplicate titles', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.addSheet('Notes');
      expect(result).to.deep.equal({ operation: 'add', sheet: 'Notes', dryRun: false });
      expect(wb.workbook.getWorksheet('Notes')).to.not.be.undefined;

      expect(() => wb.addSheet('Notes')).to.throw('already exists');
      expect(() => wb.addSheet('  ')).to.throw('required');
      expect(() => wb.addSheet('Data', { dryRun: true })).to.throw('already exists');
    });

    it('addSheet dryRun creates nothing', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.addSheet('Planned', { dryRun: true });

      expect(result.dryRun).to.equal(true);
      expect(wb.workbook.getWorksheet('Planned')).to.be.undefined;
    });

    it('removes a sheet but refuses to remove the last visible one', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.addSheet('Second');

      const result = wb.removeSheet('Second');
      expect(result).to.deep.equal({ operation: 'remove', sheet: 'Second', dryRun: false });
      expect(wb.workbook.getWorksheet('Second')).to.be.undefined;

      expect(() => wb.removeSheet('Data')).to.throw('last visible worksheet');
      expect(() => wb.removeSheet('Nope')).to.throw('No worksheet named');
      expect(() => wb.removeSheet('')).to.throw('required');
    });

    it('drops defined names scoped to the removed sheet', async () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.addSheet('Notes');
      wb.addSheet('My Notes');
      wb.addDefinedName({ name: 'Amounts', refersTo: 'Data!$B$2:$B$3' });
      wb.addDefinedName({ name: 'NoteRef', refersTo: 'Notes!$A$1:$A$2' });
      wb.addDefinedName({ name: 'QuotedRef', refersTo: "'My Notes'!$A$1" });

      wb.removeSheet('Notes');
      wb.removeSheet('My Notes');

      const expected = [{ name: 'Amounts', ranges: ['Data!$B$2:$B$3'] }];
      expect(wb.listDefinedNames()).to.deep.equal(expected);

      const filePath = path.join(tmpDir, 'removed-sheet-names.xlsx');
      await wb.save(filePath);
      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.listDefinedNames()).to.deep.equal(expected);
    });

    it('allows removing a hidden sheet while a visible one remains', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.addSheet('Archive');
      wb.workbook.getWorksheet('Archive')!.state = 'hidden';

      const result = wb.removeSheet('Archive');

      expect(result.sheet).to.equal('Archive');
      expect(wb.workbook.getWorksheet('Archive')).to.be.undefined;
      expect(wb.workbook.getWorksheet('Data')).to.not.be.undefined;
    });

    it('renames a sheet, refuses collisions, and persists', async () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.addSheet('Second');

      const result = wb.renameSheet('Second', 'Report');
      expect(result).to.deep.equal({ operation: 'rename', from: 'Second', to: 'Report', dryRun: false });
      expect(wb.workbook.getWorksheet('Report')).to.not.be.undefined;

      expect(() => wb.renameSheet('Data', 'Report')).to.throw('already exists');
      expect(() => wb.renameSheet('Nope', 'Other')).to.throw('No worksheet named');

      const filePath = path.join(tmpDir, 'renamed.xlsx');
      await wb.save(filePath);
      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.workbook.getWorksheet('Report')).to.not.be.undefined;
    });

    it('renameSheet dryRun renames nothing', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.renameSheet('Data', 'Planned', { dryRun: true });

      expect(result.dryRun).to.equal(true);
      expect(wb.workbook.getWorksheet('Data')).to.not.be.undefined;
      expect(wb.workbook.getWorksheet('Planned')).to.be.undefined;
    });
  });

  describe('defined names', () => {
    it('adds, lists and persists a defined name', async () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const result = wb.addDefinedName({ name: 'Amounts', refersTo: 'Data!$B$2:$B$3' });
      expect(result).to.deep.equal({ name: 'Amounts', refersTo: 'Data!$B$2:$B$3', dryRun: false });

      expect(wb.listDefinedNames()).to.deep.equal([{ name: 'Amounts', ranges: ['Data!$B$2:$B$3'] }]);

      const filePath = path.join(tmpDir, 'names.xlsx');
      await wb.save(filePath);
      const reloaded = await XlsxWorkbook.load(filePath);
      expect(reloaded.inspect().definedNames).to.deep.equal([{ name: 'Amounts', ranges: ['Data!$B$2:$B$3'] }]);
      expect(reloaded.listDefinedNames()).to.deep.equal([{ name: 'Amounts', ranges: ['Data!$B$2:$B$3'] }]);
    });

    it('removes a defined name together with all of its ranges', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.addDefinedName({ name: 'Amounts', refersTo: 'Data!$B$2:$B$3' });

      const result = wb.removeDefinedName({ name: 'Amounts' });

      expect(result.name).to.equal('Amounts');
      expect(result.removedRanges).to.deep.equal(['Data!$B$2:$B$3']);
      expect(wb.listDefinedNames()).to.deep.equal([]);
      expect(() => wb.removeDefinedName({ name: 'Amounts' })).to.throw('No defined name');
      expect(() => wb.removeDefinedName({ name: '' })).to.throw('required');
    });

    it('refuses duplicate names and malformed ranges', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);
      wb.addDefinedName({ name: 'Amounts', refersTo: 'Data!$B$2:$B$3' });

      expect(() => wb.addDefinedName({ name: 'Amounts', refersTo: 'Data!$C$2:$C$3' })).to.throw('already exists');
      expect(() => wb.addDefinedName({ name: 'NoSheet', refersTo: '$B$2:$B$3' })).to.throw('sheet-qualified');
      expect(() => wb.addDefinedName({ name: 'Garbage', refersTo: 'not a range' })).to.throw('sheet-qualified');
      expect(wb.listDefinedNames()).to.have.lengthOf(1);
    });

    it('dryRun changes nothing for add and remove', () => {
      const wb = XlsxWorkbook.create();
      buildSheet(wb);

      const planned = wb.addDefinedName({ name: 'Planned', refersTo: 'Data!$A$2:$A$3', dryRun: true });
      expect(planned.dryRun).to.equal(true);
      expect(wb.listDefinedNames()).to.deep.equal([]);

      wb.addDefinedName({ name: 'Amounts', refersTo: 'Data!$B$2:$B$3' });
      const removal = wb.removeDefinedName({ name: 'Amounts', dryRun: true });
      expect(removal.dryRun).to.equal(true);
      expect(removal.removedRanges).to.deep.equal(['Data!$B$2:$B$3']);
      expect(wb.listDefinedNames()).to.have.lengthOf(1);
    });
  });
});
