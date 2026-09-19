import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { expect } from 'chai';
import JSZip from 'jszip';
import {
  XlsxWorkbook,
  colLetterToIndex,
  indexToColLetter,
  parseA1Range,
  parseCellAddress,
  cellAddress,
  formatA1Range,
  ReportDocument,
} from '../src/lib/xlsx';

describe('XlsxWorkbook & XLSX Helpers', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gsheet-xlsx-test-'));
  });

  afterEach(async () => {
    if (tmpDir) {
      await fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    }
  });

  // -------------------------------------------------------------------------
  // 1. A1 Range Parser & Column Conversions
  // -------------------------------------------------------------------------
  describe('A1 Range Parsing & Column Indexing', () => {
    it('converts column letters to 1-based indices and back', () => {
      expect(colLetterToIndex('A')).to.equal(1);
      expect(colLetterToIndex('Z')).to.equal(26);
      expect(colLetterToIndex('AA')).to.equal(27);
      expect(colLetterToIndex('AZ')).to.equal(52);
      expect(colLetterToIndex('BA')).to.equal(53);
      expect(colLetterToIndex('XFD')).to.equal(16384);

      expect(indexToColLetter(1)).to.equal('A');
      expect(indexToColLetter(26)).to.equal('Z');
      expect(indexToColLetter(27)).to.equal('AA');
      expect(indexToColLetter(52)).to.equal('AZ');
      expect(indexToColLetter(53)).to.equal('BA');
      expect(indexToColLetter(16384)).to.equal('XFD');
    });

    it('throws on out-of-bounds column indices or malformed strings', () => {
      expect(() => colLetterToIndex('')).to.throw(/Invalid column string/);
      expect(() => colLetterToIndex('123')).to.throw(/Invalid column letters/);
      expect(() => colLetterToIndex('XFE')).to.throw(/exceeds maximum/);
      expect(() => indexToColLetter(0)).to.throw(/out of bounds/);
      expect(() => indexToColLetter(16385)).to.throw(/out of bounds/);
    });

    it('parses single cells and cell addresses', () => {
      expect(cellAddress(1, 1)).to.equal('A1');
      expect(cellAddress(10, 27)).to.equal('AA10');

      const addr = parseCellAddress('B5');
      expect(addr.row).to.equal(5);
      expect(addr.col).to.equal(2);

      const parsedSingle = parseA1Range('C15');
      expect(parsedSingle.isSingleCell).to.be.true;
      expect(parsedSingle.startCol).to.equal(3);
      expect(parsedSingle.startRow).to.equal(15);
      expect(parsedSingle.endCol).to.equal(3);
      expect(parsedSingle.endRow).to.equal(15);
    });

    it('parses bounded ranges and normalizes inverted bounds', () => {
      const parsed = parseA1Range('A1:C10');
      expect(parsed.startCol).to.equal(1);
      expect(parsed.startRow).to.equal(1);
      expect(parsed.endCol).to.equal(3);
      expect(parsed.endRow).to.equal(10);
      expect(parsed.isSingleCell).to.be.false;

      // Inverted: bottom-right to top-left
      const inverted = parseA1Range('C10:A1');
      expect(inverted.startCol).to.equal(1);
      expect(inverted.startRow).to.equal(1);
      expect(inverted.endCol).to.equal(3);
      expect(inverted.endRow).to.equal(10);
    });

    it('parses sheet-qualified ranges and escaped single quotes', () => {
      const standard = parseA1Range('Sheet1!A1:B5');
      expect(standard.sheetName).to.equal('Sheet1');
      expect(standard.startCol).to.equal(1);
      expect(standard.endRow).to.equal(5);

      const quoted = parseA1Range("'Monthly Summary'!B2:D20");
      expect(quoted.sheetName).to.equal('Monthly Summary');
      expect(quoted.startCol).to.equal(2);
      expect(quoted.endRow).to.equal(20);

      const escaped = parseA1Range("'O''Reilly''s Dept'!A1:C3");
      expect(escaped.sheetName).to.equal("O'Reilly's Dept");
      expect(escaped.startCol).to.equal(1);
      expect(escaped.endCol).to.equal(3);
    });

    it('parses full column and full row ranges', () => {
      const fullCol = parseA1Range('A:C', { maxRows: 500 });
      expect(fullCol.isFullCol).to.be.true;
      expect(fullCol.startCol).to.equal(1);
      expect(fullCol.endCol).to.equal(3);
      expect(fullCol.endRow).to.equal(500);

      const fullRow = parseA1Range('2:10', { maxCols: 100 });
      expect(fullRow.isFullRow).to.be.true;
      expect(fullRow.startRow).to.equal(2);
      expect(fullRow.endRow).to.equal(10);
      expect(fullRow.endCol).to.equal(100);
    });

    it('formats ranges to valid A1 strings', () => {
      expect(
        formatA1Range({ sheetName: 'Sheet1', startCol: 1, startRow: 1, endCol: 3, endRow: 10 })
      ).to.equal('Sheet1!A1:C10');

      expect(
        formatA1Range({ sheetName: 'My Sheet', startCol: 2, startRow: 2, endCol: 2, endRow: 2 })
      ).to.equal("'My Sheet'!B2");
    });

    it('throws on malformed range syntax', () => {
      expect(() => parseA1Range('')).to.throw(/Empty or invalid/);
      expect(() => parseA1Range('!A1')).to.throw(/empty sheet name/);
      expect(() => parseA1Range("'Unclosed!A1")).to.throw(/Malformed sheet-qualified/);
      expect(() => parseA1Range('A0')).to.throw(/Invalid A1 range syntax/);
      expect(() => parseA1Range('A1:B0')).to.throw(/Invalid A1 range syntax/);
    });
  });

  // -------------------------------------------------------------------------
  // 2. Workbook Creation, Inspection & Capability Detection
  // -------------------------------------------------------------------------
  describe('Creation & Inspection', () => {
    it('creates a new workbook and inspects capabilities', () => {
      const wb = XlsxWorkbook.create({ creator: 'TestRunner' });
      expect(wb.workbook.creator).to.equal('TestRunner');

      const ws = wb.workbook.addWorksheet('Data');
      ws.getCell('A1').value = 'Header';
      ws.getCell('A2').value = 123;
      ws.getCell('A3').value = { formula: 'SUM(A1:A2)', result: 123 };

      const inspection = wb.inspect();
      expect(inspection.sheetCount).to.equal(1);
      expect(inspection.sheets[0].name).to.equal('Data');
      expect(inspection.sheets[0].hasFormulas).to.be.true;
      expect(inspection.capabilities.canRead).to.be.true;
      expect(inspection.capabilities.canWriteSafely).to.be.true;
      expect(inspection.capabilities.canRecalculateFormulas).to.be.false;
      expect(inspection.hasUnsupportedFeatures).to.be.false;
    });
  });

  // -------------------------------------------------------------------------
  // 3. Range Reading, Formula Cache Freshness & Modes
  // -------------------------------------------------------------------------
  describe('Read Modes & Formula Cache Freshness', () => {
    it('reads cells in unformatted, formatted, and formula modes', async () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Report');

      ws.getCell('A1').value = 'Revenue';
      ws.getCell('B1').value = 1500000;
      ws.getCell('B1').numFmt = '$#,##0.00';

      // Cell with formula + cached result
      ws.getCell('C1').value = { formula: 'B1*1.1', result: 1650000 };

      // Cell with formula but NO cached result (stale/unknown)
      ws.getCell('D1').value = { formula: 'SUM(A1:C1)', result: undefined };

      // Literal string starting with '='
      ws.getCell('E1').value = '=LiteralEqualSign';

      // 1. Unformatted mode
      const unformatted = wb.read('Report!A1:E1', { mode: 'unformatted' });
      expect(unformatted.values[0][0]).to.equal('Revenue');
      expect(unformatted.values[0][1]).to.equal(1500000);
      expect(unformatted.values[0][2]).to.equal(1650000); // Cached result
      expect(unformatted.values[0][3]).to.be.null; // No cached result
      expect(unformatted.values[0][4]).to.equal('=LiteralEqualSign');

      // 2. Formatted mode
      const formatted = wb.read('Report!A1:E1', { mode: 'formatted' });
      expect(formatted.values[0][0]).to.equal('Revenue');
      expect(formatted.values[0][4]).to.equal('=LiteralEqualSign');

      // 3. Formula mode
      const formulaMode = wb.read('Report!A1:E1', { mode: 'formula' });
      expect(formulaMode.formulaCount).to.equal(2);
      expect(formulaMode.freshnessSummary.cachedCount).to.equal(1);
      expect(formulaMode.freshnessSummary.unknownCount).to.equal(1);
      expect(formulaMode.freshnessSummary.notFormulaCount).to.equal(3);

      const c1Detail = formulaMode.cellDetails[0][2];
      expect(c1Detail.formula).to.equal('B1*1.1');
      expect(c1Detail.freshness).to.equal('cached');
      expect(c1Detail.cachedResult).to.equal(1650000);

      const d1Detail = formulaMode.cellDetails[0][3];
      expect(d1Detail.formula).to.equal('SUM(A1:C1)');
      expect(d1Detail.freshness).to.equal('unknown');
      // XlsxCellReadInfo.cachedResult is optional and nullable: ExcelJS stores
      // result:undefined as null, so an absent cache legitimately surfaces as
      // null or undefined. Only a materialized value would contradict the
      // unknown freshness; do not pin one nullish sentinel over the other.
      const d1Cache: unknown = d1Detail.cachedResult;
      expect(d1Cache === null || d1Cache === undefined, 'absent formula cache must stay nullish').to.be.true;

      const e1Detail = formulaMode.cellDetails[0][4];
      expect(e1Detail.formula).to.be.undefined;
      expect(e1Detail.freshness).to.equal('not-formula');
      expect(e1Detail.value).to.equal('=LiteralEqualSign');
    });
    it('reads translated relative formulas for shared-formula slave cells without returning master address', () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('SharedSheet');
      ws.getCell('A1').value = 10;
      ws.getCell('A2').value = 20;

      // Master cell B1
      ws.getCell('B1').value = { formula: 'A1*2', result: 20 };
      // Slave cell B2 sharing formula from master B1
      ws.getCell('B2').value = { sharedFormula: 'B1', result: 40 };

      const readRes = wb.read('SharedSheet!B1:B2', { mode: 'formula' });
      expect(readRes.cellDetails[0][0].formula).to.equal('A1*2');
      expect(readRes.cellDetails[1][0].formula).to.equal('A2*2'); // Relative-adjusted!
      expect(readRes.cellDetails[1][0].formula).to.not.equal('B1'); // Never master address!
      expect(readRes.cellDetails[1][0].cachedResult).to.equal(40);
    });
  });

  // -------------------------------------------------------------------------
  // 4. Report Document Apply, Styles & Managed Extent Clearing
  // -------------------------------------------------------------------------
  describe('Report Document Apply & Managed Range Clearing', () => {
    it('applies report document with start cell, presentation and number formats', () => {
      const wb = XlsxWorkbook.create();
      const doc: ReportDocument = {
        provenance: {
          templateId: 'finance-summary-v1',
          templateVersion: 1,
          sourceHash: 'hash123',
        },
        sheets: [
          {
            name: 'Financials',
            startCell: 'B2',
            freezeRows: 1,
            columnWidths: [20, 15, 15],
            numberFormats: [{ column: 2, format: '$#,##0' }],
            rows: [
              ['Category', 'Amount', 'Tax'],
              ['Consulting', 10000, { formula: 'C3*0.1', result: 1000 }],
              ['Software', 25000, { formula: 'C4*0.1', result: 2500 }],
            ],
          },
        ],
      };

      const result = wb.apply(doc);
      expect(result.success).to.be.true;
      expect(result.appliedChanges).to.be.greaterThan(0);

      const ws = wb.workbook.getWorksheet('Financials');
      expect(ws).to.exist;
      expect(ws!.getCell('B2').value).to.equal('Category');
      expect(ws!.getCell('C3').value).to.equal(10000);
      expect(ws!.getCell('C3').numFmt).to.equal('$#,##0');

      // Formula cell check
      const taxCell = ws!.getCell('D3');
      expect(taxCell.formula).to.equal('C3*0.1');
      expect(taxCell.result).to.equal(1000);

      // Freeze view check
      const frozenView = ws!.views[0];
      if (frozenView.state !== 'frozen') {
        throw new Error(`Expected frozen worksheet view, got state=${String(frozenView.state)}`);
      }
      expect(frozenView.ySplit).to.equal(1);

      // Column widths
      expect(ws!.getColumn(2).width).to.equal(20);
    });

    it('preserves existing template formatting unless overridden', () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Styled');

      // Pre-style a cell in the template
      const headerCell = ws.getCell('A1');
      headerCell.value = 'Old Header';
      headerCell.font = { bold: true, color: { argb: 'FFFF0000' } };

      const doc: ReportDocument = {
        provenance: { templateId: 't1', templateVersion: 1, sourceHash: 'h1' },
        sheets: [
          {
            name: 'Styled',
            startCell: 'A1',
            rows: [['New Header']],
          },
        ],
      };

      wb.apply(doc, { preserveTemplateStyles: true });

      const updatedCell = ws.getCell('A1');
      expect(updatedCell.value).to.equal('New Header');
      expect(updatedCell.font?.bold).to.be.true;
      expect(updatedCell.font?.color?.argb).to.equal('FFFF0000');
    });

    it('clears previous managed extent on shorter rerun', () => {
      const wb = XlsxWorkbook.create();

      // First run: 4 data rows (A1:B4)
      const doc1: ReportDocument = {
        provenance: { templateId: 't1', templateVersion: 1, sourceHash: 'h1' },
        sheets: [
          {
            name: 'Data',
            clearManagedRange: true,
            rows: [
              ['Item 1', 10],
              ['Item 2', 20],
              ['Item 3', 30],
              ['Item 4', 40],
            ],
          },
        ],
      };
      wb.apply(doc1);

      const ws = wb.workbook.getWorksheet('Data');
      expect(ws!.getCell('A4').value).to.equal('Item 4');
      expect(ws!.getCell('B4').value).to.equal(40);

      // Second run: only 2 data rows (A1:B2) with clearManagedRange
      const doc2: ReportDocument = {
        provenance: { templateId: 't1', templateVersion: 1, sourceHash: 'h2' },
        sheets: [
          {
            name: 'Data',
            clearManagedRange: true,
            rows: [
              ['Item 1', 100],
              ['Item 2', 200],
            ],
          },
        ],
      };
      wb.apply(doc2);

      expect(ws!.getCell('A1').value).to.equal('Item 1');
      expect(ws!.getCell('B1').value).to.equal(100);
      expect(ws!.getCell('A2').value).to.equal('Item 2');
      expect(ws!.getCell('B2').value).to.equal(200);

      // Stale rows (A3:B4) must be cleared (null)
      expect(ws!.getCell('A3').value).to.be.null;
      expect(ws!.getCell('B3').value).to.be.null;
      expect(ws!.getCell('A4').value).to.be.null;
      expect(ws!.getCell('B4').value).to.be.null;
    });
  });

  // -------------------------------------------------------------------------
  // 5. Formula Overwrite Protection & Diff Preview
  // -------------------------------------------------------------------------
  describe('Formula Overwrite Safety & Preview', () => {
    it('refuses to overwrite existing formula by default', () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Calc');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = { formula: 'A1*2', result: 200 };

      const doc: ReportDocument = {
        provenance: { templateId: 't1', templateVersion: 1, sourceHash: 'h1' },
        sheets: [
          {
            name: 'Calc',
            startCell: 'A1',
            rows: [
              [500],
              [1000], // Attempts to overwrite formula cell A2 with scalar 1000
            ],
          },
        ],
      };

      // Preview should show conflict
      const diff = wb.preview(doc);
      expect(diff.totalConflicts).to.equal(1);
      expect(diff.conflicts[0].address).to.equal('A2');
      expect(diff.conflicts[0].conflictReason).to.include('already contains formula');

      // Apply should throw without overwriteFormulas: true
      expect(() => wb.apply(doc)).to.throw(/formula overwrite conflict/);

      // Verify cell was NOT mutated
      expect(ws.getCell('A2').formula).to.equal('A1*2'); // Still formula
      // With overwriteFormulas: true, it should succeed
      const okResult = wb.apply(doc, { overwriteFormulas: true });
      expect(okResult.success).to.be.true;
      expect(ws.getCell('A2').value).to.equal(1000);
    });

    it('dry run does not mutate workbook', () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Test');
      ws.getCell('A1').value = 'Original';

      const doc: ReportDocument = {
        provenance: { templateId: 't1', templateVersion: 1, sourceHash: 'h1' },
        sheets: [
          {
            name: 'Test',
            rows: [['Mutated']],
          },
        ],
      };

      const res = wb.apply(doc, { dryRun: true });
      expect(res.success).to.be.true;
      expect(res.appliedChanges).to.equal(0);
      expect(ws.getCell('A1').value).to.equal('Original');
    });
    it('allows identical normalized formula re-application without requiring overwriteFormulas', () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Calc');
      ws.getCell('A1').value = 100;
      ws.getCell('A2').value = { formula: 'A1*2', result: 200 };

      const doc: ReportDocument = {
        provenance: { templateId: 't1', templateVersion: 1, sourceHash: 'h1' },
        sheets: [
          {
            name: 'Calc',
            startCell: 'A1',
            rows: [
              [100],
              [{ formula: 'A1*2', result: 200 }], // Identical formula
            ],
          },
        ],
      };

      const diff = wb.preview(doc);
      expect(diff.totalConflicts).to.equal(0);
      expect(diff.conflicts).to.have.lengthOf(0);

      // Apply succeeds without overwriteFormulas: true
      const result = wb.apply(doc);
      expect(result.success).to.be.true;
      expect(ws.getCell('A2').formula).to.equal('A1*2');
    });
  });

  // -------------------------------------------------------------------------
  // 6. Safe Atomic Save, Backup, and Conflict Detection
  // -------------------------------------------------------------------------
  describe('Safe Atomic File Saving & Conflict Detection', () => {
    it('saves workbook atomically to disk and computes SHA-256', async () => {
      const wb = XlsxWorkbook.create();
      const ws = wb.workbook.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Persisted Data';

      const savePath = path.join(tmpDir, 'output.xlsx');
      const saveRes = await wb.save(savePath);

      expect(saveRes.savedPath).to.equal(path.resolve(savePath));
      expect(saveRes.bytesWritten).to.be.greaterThan(0);
      expect(saveRes.sha256Hash).to.have.length(64);
      expect(fs.existsSync(savePath)).to.be.true;

      // Reload saved file
      const loaded = await XlsxWorkbook.load(savePath);
      const readRes = loaded.read('Sheet1!A1');
      expect(readRes.values[0][0]).to.equal('Persisted Data');
    });

    it('requires overwrite: true or inPlace: true when target file already exists', async () => {
      const wb = XlsxWorkbook.create();
      wb.workbook.addWorksheet('Sheet1').getCell('A1').value = 'First';

      const filePath = path.join(tmpDir, 'existing.xlsx');
      await wb.save(filePath);

      const wb2 = XlsxWorkbook.create();
      wb2.workbook.addWorksheet('Sheet1').getCell('A1').value = 'Second';

      // Save without overwrite flag should throw
      let threw = false;
      try {
        await wb2.save(filePath);
      } catch (err: unknown) {
        threw = true;
        expect((err as Error).message).to.include('already exists');
      }
      expect(threw).to.be.true;

      // With overwrite: true, it replaces the file
      const overwriteRes = await wb2.save(filePath, { overwrite: true });
      expect(overwriteRes.isOverwritten).to.be.true;
    });

    it('creates a backup copy when inPlace is true', async () => {
      const initialPath = path.join(tmpDir, 'original.xlsx');
      const wb1 = XlsxWorkbook.create();
      wb1.workbook.addWorksheet('Sheet1').getCell('A1').value = 'Original Content';
      await wb1.save(initialPath);

      // Load original
      const loaded = await XlsxWorkbook.load(initialPath);
      loaded.workbook.getWorksheet('Sheet1')!.getCell('A1').value = 'Modified Content';

      // Save inPlace
      const saveRes = await loaded.save(undefined, { inPlace: true });
      expect(saveRes.backupCreated).to.exist;
      expect(fs.existsSync(saveRes.backupCreated!)).to.be.true;

      // Verify original backup has old content
      const backupWb = await XlsxWorkbook.load(saveRes.backupCreated!);
      expect(backupWb.read('Sheet1!A1').values[0][0]).to.equal('Original Content');
    });

    it('detects disk conflict when file changes before in-place save', async () => {
      const targetFile = path.join(tmpDir, 'conflict-target.xlsx');
      const wb = XlsxWorkbook.create();
      wb.workbook.addWorksheet('Sheet1').getCell('A1').value = 'V1';
      await wb.save(targetFile);

      // Session A loads target
      const sessionA = await XlsxWorkbook.load(targetFile);

      // External modification happens on disk
      await fs.promises.writeFile(targetFile, Buffer.from('externally modified content'));

      // Session A tries to save inPlace -> should detect fingerprint mismatch and throw conflict error
      let conflictDetected = false;
      try {
        await sessionA.save(undefined, { inPlace: true });
      } catch (err: unknown) {
        conflictDetected = true;
        expect((err as Error).message).to.include('Conflict detected');
      }
      expect(conflictDetected).to.be.true;
    });
  });

  // -------------------------------------------------------------------------
  // 7. JSZip Preflight Inspection & Unsupported Feature Guards
  // -------------------------------------------------------------------------
  describe('JSZip Preflight & Unsupported Feature Safeguards', () => {
    it('detects macros (VBA) in XLSX archive and refuses unsafe save without override', async () => {
      const vbaZip = new JSZip();
      vbaZip.file('[Content_Types].xml', '<Types></Types>');
      vbaZip.file('xl/workbook.xml', '<workbook><sheets><sheet name="MacroSheet" sheetId="1"/></sheets></workbook>');
      vbaZip.file('xl/vbaProject.bin', Buffer.from('synthetic-vba-binary'));
      const buffer = await vbaZip.generateAsync({ type: 'nodebuffer' });

      const macroPath = path.join(tmpDir, 'macro-test.xlsm');
      await fs.promises.writeFile(macroPath, buffer);

      const loaded = await XlsxWorkbook.load(macroPath);
      const inspection = loaded.inspect();

      expect(inspection.hasUnsupportedFeatures).to.be.true;
      expect(inspection.capabilities.canWriteSafely).to.be.false;
      expect(inspection.unsupportedFeatures[0].type).to.equal('macro');

      // Attempting to save without override should throw
      const outPath = path.join(tmpDir, 'macro-out.xlsx');
      let rejected = false;
      try {
        await loaded.save(outPath);
      } catch (err: unknown) {
        rejected = true;
        expect((err as Error).message).to.include('Cannot safely save workbook containing unsupported features');
      }
      expect(rejected).to.be.true;

      // Saving with allowUnsupportedFeatures: true succeeds
      const forcedSave = await loaded.save(outPath, { allowUnsupportedFeatures: true });
      expect(forcedSave.savedPath).to.exist;
    });

    it('detects charts and pivot tables during preflight', async () => {
      const chartZip = new JSZip();
      chartZip.file('xl/workbook.xml', '<workbook><sheets><sheet name="ChartSheet" sheetId="1"/></sheets></workbook>');
      chartZip.file('xl/charts/chart1.xml', '<c:chartSpace></c:chartSpace>');
      chartZip.file('xl/pivotTables/pivotTable1.xml', '<pivotTableDefinition></pivotTableDefinition>');
      chartZip.file('xl/externalLinks/externalLink1.xml', '<externalLink></externalLink>');

      const buffer = await chartZip.generateAsync({ type: 'nodebuffer' });
      const chartPath = path.join(tmpDir, 'features-test.xlsx');
      await fs.promises.writeFile(chartPath, buffer);

      const loaded = await XlsxWorkbook.load(chartPath);
      const inspection = loaded.inspect();

      expect(inspection.hasUnsupportedFeatures).to.be.true;
      const types = inspection.unsupportedFeatures.map((f) => f.type);
      expect(types).to.include('chart');
      expect(types).to.include('pivotTable');
      expect(types).to.include('externalLink');
    });
  });
});
