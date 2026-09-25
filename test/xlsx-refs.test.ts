import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { XlsxWorkbook } from '../src/lib/xlsx';
import {
  diffDefinedNameRanges,
  rewriteFormulaRefs,
  tokenizeFormulaRefs,
  XlsxRefSpliceOp,
} from '../src/lib/xlsx-refs';

/**
 * Offline coverage for F7 (--updateRefs): the A1-ref tokenizer and shift engine
 * as pure units, the defined-name diff, the rewrite wired through
 * XlsxWorkbook.splice, and the --updateRefs command surface. No network.
 */

const shift = (
  formula: string,
  mode: 'insert' | 'delete',
  start: number,
  count: number,
  dimension: XlsxRefSpliceOp['dimension'] = 'ROWS'
): { formula: string; changed: boolean; refErrors: number } => {
  return rewriteFormulaRefs(formula, { dimension, start, count, mode });
};

describe('tokenizeFormulaRefs', () => {
  it('emits cell and range tokens with parsed coordinates and offsets', () => {
    const tokens = tokenizeFormulaRefs('SUM(A5:A20)+B1');
    expect(tokens).to.have.length(2);
    expect(tokens[0]).to.deep.include({ kind: 'range', text: 'A5:A20', index: 4 });
    expect(tokens[0].first).to.deep.equal({ col: 1, row: 5, colAbsolute: false, rowAbsolute: false });
    expect(tokens[0].second).to.deep.equal({ col: 1, row: 20, colAbsolute: false, rowAbsolute: false });
    expect(tokens[1]).to.deep.include({ kind: 'cell', text: 'B1', index: 12 });
  });

  it('keeps $ anchors in the parsed points', () => {
    const tokens = tokenizeFormulaRefs('$B$2+B$3+$C4');
    expect(tokens.map((t) => t.text)).to.deep.equal(['$B$2', 'B$3', '$C4']);
    expect(tokens[0].first.colAbsolute).to.be.true;
    expect(tokens[0].first.rowAbsolute).to.be.true;
    expect(tokens[1].first.colAbsolute).to.be.false;
    expect(tokens[1].first.rowAbsolute).to.be.true;
    expect(tokens[2].first.colAbsolute).to.be.true;
    expect(tokens[2].first.rowAbsolute).to.be.false;
  });

  it('skips references inside string literals', () => {
    const tokens = tokenizeFormulaRefs('IF(A1="A1","yes A2","no")');
    expect(tokens.map((t) => t.text)).to.deep.equal(['A1']);
  });

  it('never misreads function names like LOG10 as cell LOG10', () => {
    expect(tokenizeFormulaRefs('LOG10(A1)').map((t) => t.text)).to.deep.equal(['A1']);
    expect(tokenizeFormulaRefs('ROUND(LOG10(B2)+SUM(C3:C4),2)').map((t) => t.text)).to.deep.equal(['B2', 'C3:C4']);
  });

  it('skips numbers, booleans and defined-name-like identifiers', () => {
    expect(tokenizeFormulaRefs('A1+1.5E3+TRUE+FALSE+MyName+123').map((t) => t.text)).to.deep.equal(['A1']);
  });

  it('emits whole-column and whole-row ranges', () => {
    const tokens = tokenizeFormulaRefs('SUM(A:A)+SUM(1:3)');
    expect(tokens).to.have.length(2);
    expect(tokens[0]).to.deep.include({ kind: 'column-range', text: 'A:A' });
    expect(tokens[0].first.col).to.equal(1);
    expect(tokens[1]).to.deep.include({ kind: 'row-range', text: '1:3' });
    expect(tokens[1].first.row).to.equal(1);
    expect(tokens[1].second!.row).to.equal(3);
  });

  it('leaves sheet-qualified refs out entirely, quoted or not', () => {
    expect(tokenizeFormulaRefs('Sheet2!A1+B2').map((t) => t.text)).to.deep.equal(['B2']);
    expect(tokenizeFormulaRefs("'My Sheet'!$B$2+C3").map((t) => t.text)).to.deep.equal(['C3']);
    // a sheet-qualified range is skipped whole, including its range partner
    expect(tokenizeFormulaRefs('SUM(Sheet2!A1:B2)+C1').map((t) => t.text)).to.deep.equal(['C1']);
  });

  it('treats a word before ! as a sheet name even when it looks like a cell', () => {
    expect(tokenizeFormulaRefs('A1!B2+B3').map((t) => t.text)).to.deep.equal(['B3']);
  });

  it('skips error literals and lower-case refs keep their case', () => {
    expect(tokenizeFormulaRefs('SUM(#REF!)')).to.have.length(0);
    expect(tokenizeFormulaRefs('IF(N(A1),#N/A,1)').map((t) => t.text)).to.deep.equal(['A1']);
    expect(tokenizeFormulaRefs('sum(a1)').map((t) => t.text)).to.deep.equal(['a1']);
  });
});

describe('rewriteFormulaRefs', () => {
  it('grows a range that straddles the insertion point (AC-04 example)', () => {
    expect(shift('=SUM(A5:A20)', 'insert', 10, 3)).to.deep.equal({ formula: '=SUM(A5:A23)', changed: true, refErrors: 0 });
  });

  it('shifts refs at or after the insertion point and leaves refs above untouched', () => {
    expect(shift('SUM(A5:A20)', 'insert', 2, 3).formula).to.equal('SUM(A8:A23)');
    expect(shift('SUM(A1:A3)', 'insert', 10, 2)).to.deep.equal({ formula: 'SUM(A1:A3)', changed: false, refErrors: 0 });
    expect(shift('A5+A1', 'insert', 3, 1).formula).to.equal('A6+A1');
  });

  it('collapses refs fully inside a deleted span to #REF!', () => {
    expect(shift('SUM(B2:B4)', 'delete', 2, 3)).to.deep.equal({ formula: 'SUM(#REF!)', changed: true, refErrors: 1 });
    expect(shift('B2', 'delete', 2, 1).formula).to.equal('#REF!');
  });

  it('collapses refs pushed past the sheet edge by an insert to #REF!', () => {
    expect(shift('=SUM(A1048575:A1048576)', 'insert', 1, 2)).to.deep.equal({
      formula: '=SUM(#REF!)',
      changed: true,
      refErrors: 1,
    });
    expect(shift('A1048576', 'insert', 1, 1).formula).to.equal('#REF!');
    // the boundary itself stays legal
    expect(shift('A1048575', 'insert', 1, 1).formula).to.equal('A1048576');
  });

  it('shrinks partially overlapping refs and shifts refs below a deleted span up', () => {
    expect(shift('SUM(A5:A20)', 'delete', 10, 3).formula).to.equal('SUM(A5:A17)');
    expect(shift('SUM(A8:A11)', 'delete', 10, 3).formula).to.equal('SUM(A8:A9)');
    expect(shift('SUM(A10:A12)', 'delete', 2, 3).formula).to.equal('SUM(A7:A9)');
  });

  it('preserves $ anchors while shifting', () => {
    expect(shift('$B$2', 'insert', 1, 5).formula).to.equal('$B$7');
    expect(shift('B$2', 'insert', 1, 5).formula).to.equal('B$7');
    expect(shift('$B2', 'insert', 1, 5).formula).to.equal('$B7');
    expect(shift('$A$5:$A$20', 'delete', 10, 3).formula).to.equal('$A$5:$A$17');
  });

  it('rewrites only the real refs, never literals, names or numbers', () => {
    expect(shift('IF(A1="A1","y","n")', 'insert', 1, 1).formula).to.equal('IF(A2="A1","y","n")');
    expect(shift('LOG10(A1)*1.5E3+TRUE', 'insert', 1, 1).formula).to.equal('LOG10(A2)*1.5E3+TRUE');
  });

  it('leaves cross-sheet refs untouched in phase 1', () => {
    expect(shift('SUM(Sheet2!B2:B4)+C1', 'insert', 1, 1).formula).to.equal('SUM(Sheet2!B2:B4)+C2');
    expect(shift("SUM('My Sheet'!$B$2)", 'delete', 1, 10).formula).to.equal("SUM('My Sheet'!$B$2)");
  });

  it('shifts whole-column refs only on COLUMNS ops and whole-row refs only on ROWS ops', () => {
    expect(shift('SUM(A:A)', 'insert', 5, 2).changed).to.be.false;
    expect(shift('SUM(A:A)', 'insert', 1, 2, 'COLUMNS').formula).to.equal('SUM(C:C)');
    expect(shift('SUM(1:3)', 'insert', 2, 2).formula).to.equal('SUM(1:5)');
    expect(shift('SUM(1:3)', 'insert', 2, 2, 'COLUMNS').changed).to.be.false;
  });

  it('keeps existing #REF! errors as-is without counting them', () => {
    expect(shift('SUM(#REF!)', 'delete', 1, 5)).to.deep.equal({ formula: 'SUM(#REF!)', changed: false, refErrors: 0 });
  });

  it('returns formulas without refs unchanged', () => {
    expect(shift('', 'insert', 1, 1)).to.deep.equal({ formula: '', changed: false, refErrors: 0 });
    expect(shift('"just a string A1"', 'insert', 1, 1).changed).to.be.false;
  });
});

describe('diffDefinedNameRanges', () => {
  const names = (entries: Record<string, string[]>): Map<string, string[]> => new Map(Object.entries(entries));

  it('counts moved ranges as rewritten', () => {
    expect(diffDefinedNameRanges(names({ N: ['Data!$B$2:$B$4'] }), names({ N: ['Data!$B$4:$B$6'] }))).to.deep.equal({
      rewritten: 1,
      broken: 0,
    });
  });

  it('counts ranges removed with the deleted span as broken', () => {
    expect(diffDefinedNameRanges(names({ N: ['Data!$B$3'] }), names({}))).to.deep.equal({ rewritten: 0, broken: 1 });
    expect(diffDefinedNameRanges(names({ N: ['Data!$A$1', 'Data!$B$2'] }), names({ N: ['Data!$A$1'] }))).to.deep.equal({
      rewritten: 0,
      broken: 1,
    });
  });

  it('reports nothing when names are untouched', () => {
    const same = names({ N: ['Data!$A$1:$A$5'] });
    expect(diffDefinedNameRanges(same, same)).to.deep.equal({ rewritten: 0, broken: 0 });
  });
});

describe('XlsxWorkbook.splice with updateRefs', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gsheet-xlsx-refs-'));
  });

  afterEach(async () => {
    if (tmpDir) {
      await fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    }
  });

  const buildSheet = (wb: XlsxWorkbook, name = 'Data') => {
    const ws = wb.workbook.addWorksheet(name);
    ws.getCell('A1').value = { formula: 'SUM(A5:A20)', result: 0 };
    ws.getCell('B5').value = { formula: 'SUM(B2:B4)', result: 0 };
    return ws;
  };

  it('rewrites same-sheet refs on insert and reports the counters (AC-04)', () => {
    const wb = XlsxWorkbook.create();
    const ws = buildSheet(wb);

    const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 10, count: 3, updateRefs: true });

    expect(ws.getCell('A1').formula).to.equal('SUM(A5:A23)');
    // B5 and its references both sit above the insertion point: untouched
    expect(ws.getCell('B5').formula).to.equal('SUM(B2:B4)');
    expect(result.refsRewritten).to.equal(1);
    expect(result.refsBroken).to.equal(0);
    expect(result.warnings.some((w) => w.includes('not rewritten'))).to.be.false;
  });

  it('leaves formulas stale and warns without the flag', () => {
    const wb = XlsxWorkbook.create();
    const ws = buildSheet(wb);

    const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 10, count: 3 });

    expect(ws.getCell('A1').formula).to.equal('SUM(A5:A20)');
    expect(result.refsRewritten).to.equal(0);
    expect(result.refsBroken).to.equal(0);
    expect(result.warnings.some((w) => w.includes('not rewritten'))).to.be.true;
  });

  it('collapses refs inside the deleted span to #REF! on delete', () => {
    const wb = XlsxWorkbook.create();
    const ws = buildSheet(wb);

    const result = wb.splice('delete', { worksheetTitle: 'Data', dimension: 'ROWS', start: 2, count: 3, updateRefs: true });

    // B5 moved up to B2; its whole reference range B2:B4 was deleted
    expect(ws.getCell('B2').formula).to.equal('SUM(#REF!)');
    expect(result.refsBroken).to.equal(1);
  });

  it('single-shifts shared-formula slaves after a splice instead of double-shifting them', async () => {
    const wb = XlsxWorkbook.create();
    const ws = wb.workbook.addWorksheet('Data');
    ws.getCell('B1').value = { formula: 'SUM(A5:A10)', result: 0 };
    // B2/B3 share the master formula with a +1/+2 row offset: pre-splice they
    // resolve to SUM(A6:A11) and SUM(A7:A12)
    ws.getCell('B2').value = { sharedFormula: 'B1', result: 0 };
    ws.getCell('B3').value = { sharedFormula: 'B1', result: 0 };

    wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 2, count: 1, updateRefs: true });

    const filePath = path.join(tmpDir, 'shared-formula.xlsx');
    await wb.save(filePath);
    const reloaded = await XlsxWorkbook.load(filePath);
    const sheet = reloaded.workbook.getWorksheet('Data')!;
    expect(sheet.getCell('B1').formula).to.equal('SUM(A6:A11)');
    // the slaves themselves moved down with the insert (old B2->B3, B3->B4);
    // reading them live after the master was rewritten (the old behavior)
    // would shift their refs twice
    expect(sheet.getCell('B3').formula).to.equal('SUM(A7:A12)');
    expect(sheet.getCell('B4').formula).to.equal('SUM(A8:A13)');
  });

  it('rewrites refs on COLUMNS ops', () => {
    const wb = XlsxWorkbook.create();
    const ws = wb.workbook.addWorksheet('Data');
    ws.getCell('A1').value = { formula: 'SUM(B2:D2)', result: 0 };

    wb.splice('insert', { worksheetTitle: 'Data', dimension: 'COLUMNS', start: 3, count: 1, updateRefs: true });

    expect(ws.getCell('A1').formula).to.equal('SUM(B2:E2)');
  });

  it('shifts defined names along with the insert', () => {
    const wb = XlsxWorkbook.create();
    const ws = buildSheet(wb);
    wb.workbook.definedNames.add('Data!$B$2:$B$4', 'DataBlock');

    const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 1, count: 2, updateRefs: true });

    const definedNamesModel = wb.workbook.definedNames as unknown as { model: { name: string; ranges: string[] }[] };
    const block = definedNamesModel.model.find((entry) => entry.name === 'DataBlock');
    expect(block!.ranges).to.deep.equal(['Data!$B$4:$B$6']);
    // 2 formula cells + 1 defined-name range
    expect(result.refsRewritten).to.equal(3);
    expect(result.refsBroken).to.equal(0);
  });

  it('drops defined names that fall fully inside the deleted span', () => {
    const wb = XlsxWorkbook.create();
    const ws = buildSheet(wb);
    wb.workbook.definedNames.add('Data!$B$3', 'Gone');

    const result = wb.splice('delete', { worksheetTitle: 'Data', dimension: 'ROWS', start: 2, count: 3, updateRefs: true });

    const definedNamesModel = wb.workbook.definedNames as unknown as { model: { name: string; ranges: string[] }[] };
    expect(definedNamesModel.model.find((entry) => entry.name === 'Gone')).to.be.undefined;
    expect(result.refsBroken).to.equal(2); // 1 cell ref + 1 defined-name range
  });

  it('leaves formulas on other sheets and cross-sheet refs untouched', () => {
    const wb = XlsxWorkbook.create();
    buildSheet(wb);
    const other = wb.workbook.addWorksheet('Other');
    other.getCell('B1').value = { formula: 'Data!B2+1', result: 0 };

    const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 1, count: 3, updateRefs: true });

    expect(other.getCell('B1').formula).to.equal('Data!B2+1');
    expect(result.refsRewritten).to.equal(2); // only the two Data sheet cells
  });

  it('rewrites nothing on a dry run even with the flag', () => {
    const wb = XlsxWorkbook.create();
    const ws = buildSheet(wb);

    const result = wb.splice('insert', { worksheetTitle: 'Data', dimension: 'ROWS', start: 10, count: 3, updateRefs: true, dryRun: true });

    expect(ws.getCell('A1').formula).to.equal('SUM(A5:A20)');
    expect(result.refsRewritten).to.equal(0);
    expect(result.refsBroken).to.equal(0);
  });
});

describe('grid insert/delete --updateRefs command surface', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gsheet-xlsx-refs-cmd-'));
  });

  afterEach(async () => {
    if (tmpDir) {
      await fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    }
  });

  const buildFixtureWorkbook = async (dir: string): Promise<string> => {
    const wb = XlsxWorkbook.create();
    const ws = wb.workbook.addWorksheet('Data');
    ws.getCell('A1').value = { formula: 'SUM(A5:A20)', result: 0 };
    const filePath = path.join(dir, 'refs-fixture.xlsx');
    await wb.save(filePath);
    return filePath;
  };

  it('grid:insert --workbook --updateRefs rewrites formulas in the saved output', async () => {
    const workbookPath = await buildFixtureWorkbook(tmpDir);
    const outputPath = path.join(tmpDir, 'out.xlsx');
    const { error, result } = await runCommand([
      'grid:insert',
      `--workbook=${workbookPath}`,
      '-t',
      'Data',
      '--dimension=ROWS',
      '--start=10',
      '--count=3',
      '--updateRefs',
      `--output=${outputPath}`,
      '--rawOutput',
    ]);
    if (error) throw error;

    const receipt = result as Record<string, unknown>;
    expect(receipt.refsRewritten).to.equal(1);
    expect(receipt.refsBroken).to.equal(0);

    const reloaded = await XlsxWorkbook.load(outputPath);
    const ws = reloaded.workbook.getWorksheet('Data');
    expect(ws!.getCell('A1').formula).to.equal('SUM(A5:A23)');
  });

  it('grid:insert refuses --updateRefs together with --spreadsheetId', async () => {
    const { error } = await runCommand([
      'grid:insert',
      '--updateRefs',
      '--spreadsheetId=SRC',
      '--worksheetTitle=T',
      '--dimension=ROWS',
      '--start=1',
    ]);
    expect(error).to.exist;
    expect(String(error?.message)).to.contain('cannot be combined');
  });

  it('grid:delete refuses --updateRefs together with --spreadsheetId', async () => {
    const { error } = await runCommand([
      'grid:delete',
      '--updateRefs',
      '--spreadsheetId=SRC',
      '--worksheetTitle=T',
      '--dimension=ROWS',
      '--start=1',
    ]);
    expect(error).to.exist;
    expect(String(error?.message)).to.contain('cannot be combined');
  });
});
