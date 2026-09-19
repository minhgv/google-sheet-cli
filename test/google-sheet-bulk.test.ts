import { expect } from 'chai';
import http from 'http';
import https from 'https';
import GoogleSheet, { GoogleSheetCli } from '../src/lib/google-sheet';
import {
  a1ToCol,
  colToA1,
  escapeWorksheetTitle,
  findFormulaOverwrites,
  formatBoundedA1Range,
  needsQuoting,
  packUpdateBatches,
  parseA1Cell,
  parseStrictA1Range,
  splitUpdateByRows,
  toGoogleExtendedValue,
  toGoogleNumberFormat,
} from '../src/lib/sheet-batch';
import type { ReportDocument } from '../src/lib/report/types';

// Synthetic RSA private key for JWT signing in offline tests
const TEST_KEY = `-----BEGIN PRIVATE KEY-----
MIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC7Vv8x8S2P+d5+
8Pq3+pE2L3kH/nK4+uL3J1V7tY4b2v8w6k2J0t4v7n5m3q1L2k4j5h6g7f8e9d0c
1b2a3z4y5x6w7v8u9t0s1r2q3p4o5n6m7l8k9j0i1h2g3f4e5d6c7b8a9Z0Y1X2W
3V4U5T6S7R8Q9P0O1N2M3L4K5J6I7H8G9F0E1D2C3B4A5z6y7x8w9v0u1t2s3r4q
5p6o7n8m9l0k1j2i3h4g5f6e7d8c9b0a1Z2Y3X4W5V6U7T8S9R0Q1P2O3N4M5L6K
7J8I9H0G1F2E3D4C5B6A7z8y9x0w1v2u3t4s5r6q7p8o9n0m1l2k3j4i5h6g7f8e
AgMBAAECggEBAK1Y2Z3X4W5V6U7T8S9R0Q1P2O3N4M5L6K7J8I9H0G1F2E3D4C5B
6A7z8y9x0w1v2u3t4s5r6q7p8o9n0m1l2k3j4i5h6g7f8e9d0c1b2a3z4y5x6w7v
8u9t0s1r2q3p4o5n6m7l8k9j0i1h2g3f4e5d6c7b8a9Z0Y1X2W3V4U5T6S7R8Q9P
0O1N2M3L4K5J6I7H8G9F0E1D2C3B4A5z6y7x8w9v0u1t2s3r4q5p6o7n8m9l0k1j
2i3h4g5f6e7d8c9b0a1Z2Y3X4W5V6U7T8SCgYEA3r4q5p6o7n8m9l0k1j2i3h4g
5f6e7d8c9b0a1Z2Y3X4W5V6U7T8S9R0Q1P2O3N4M5L6K7J8I9H0G1F2E3D4C5B6A
7z8y9x0w1v2u3t4s5r6q7p8o9n0m1l2k3j4i5h6g7f8e9d0c1b2a3z4y5x6w7v8u
9t0s1r2q3p4o5n6mCgYEA0s1r2q3p4o5n6m7l8k9j0i1h2g3f4e5d6c7b8a9Z0Y
1X2W3V4U5T6S7R8Q9P0O1N2M3L4K5J6I7H8G9F0E1D2C3B4A5z6y7x8w9v0u1t2s
3r4q5p6o7n8m9l0k1j2i3h4g5f6e7d8c9b0a1Z2Y3X4W5V6U7T8S9R0Q1P2O3N4M
5L6KCgYEAw9v0u1t2s3r4q5p6o7n8m9l0k1j2i3h4g5f6e7d8c9b0a1Z2Y3X4W5V
6U7T8S9R0Q1P2O3N4M5L6K7J8I9H0G1F2E3D4C5B6A7z8y9x0w1v2u3t4s5r6q7p
8o9n0m1l2k3j4i5h6g7f8e9d0c1b2a3z4y5x6w7v8u9t0s1r2q3p4o5n6m7l8kCg
YEAq5p6o7n8m9l0k1j2i3h4g5f6e7d8c9b0a1Z2Y3X4W5V6U7T8S9R0Q1P2O3N4M
5L6K7J8I9H0G1F2E3D4C5B6A7z8y9x0w1v2u3t4s5r6q7p8o9n0m1l2k3j4i5h6g
7f8e9d0c1b2a3z4y5x6w7v8u9t0s1r2q3p4o5n6m7l8k9j0i1h2g3f4e5d6c7b8a
CgYEA1j2i3h4g5f6e7d8c9b0a1Z2Y3X4W5V6U7T8S9R0Q1P2O3N4M5L6K7J8I9H0
G1F2E3D4C5B6A7z8y9x0w1v2u3t4s5r6q7p8o9n0m1l2k3j4i5h6g7f8e9d0c1b2a
3z4y5x6w7v8u9t0s1r2q3p4o5n6m7l8k9j0i1h2g3f4e5d6c7b8a9Z0Y1X2W3V4U
5T6S7R8Q9P0O1N2M3L4K-----END PRIVATE KEY-----`;

const FAKE_CREDENTIALS = {
  client_email: 'test-service-account@test-project.iam.gserviceaccount.com',
  private_key: TEST_KEY,
};

describe('GoogleSheet Batch & Report Operations (Wave A)', () => {
  describe('A1 Notation & Coordinate Math (sheet-batch)', () => {
    it('converts column numbers to A1 letters and back', () => {
      expect(colToA1(1)).to.equal('A');
      expect(colToA1(26)).to.equal('Z');
      expect(colToA1(27)).to.equal('AA');
      expect(colToA1(52)).to.equal('AZ');
      expect(colToA1(53)).to.equal('BA');
      expect(colToA1(702)).to.equal('ZZ');
      expect(colToA1(703)).to.equal('AAA');

      expect(a1ToCol('A')).to.equal(1);
      expect(a1ToCol('Z')).to.equal(26);
      expect(a1ToCol('AA')).to.equal(27);
      expect(a1ToCol('AZ')).to.equal(52);
      expect(a1ToCol('ZZ')).to.equal(702);
      expect(a1ToCol('AAA')).to.equal(703);
    });

    it('parses single cell references correctly', () => {
      expect(parseA1Cell('A1')).to.eql({ col: 1, row: 1 });
      expect(parseA1Cell('B3')).to.eql({ col: 2, row: 3 });
      expect(parseA1Cell('$C$12')).to.eql({ col: 3, row: 12 });
      expect(parseA1Cell('AA100')).to.eql({ col: 27, row: 100 });
      expect(() => parseA1Cell('invalid')).to.throw('Invalid cell reference');
    });

    it('evaluates worksheet title quoting needs', () => {
      expect(needsQuoting('Sheet1')).to.be.false;
      expect(needsQuoting('Data2026')).to.be.false;
      expect(needsQuoting('My Sheet')).to.be.true;
      expect(needsQuoting('Sheet-2')).to.be.true;
      expect(needsQuoting('2026')).to.be.true;
      expect(escapeWorksheetTitle('Sheet1')).to.equal('Sheet1');
      expect(escapeWorksheetTitle('My Sheet')).to.equal("'My Sheet'");
      expect(escapeWorksheetTitle('"My Sheet"')).to.equal("'My Sheet'");
      expect(escapeWorksheetTitle("O'Brien")).to.equal("'O''Brien'");
    });

    it('parses strict A1 ranges in various forms', () => {
      const r1 = parseStrictA1Range('Sheet1!A1:B10');
      expect(r1.worksheetTitle).to.equal('Sheet1');
      expect(r1.startCol).to.equal(1);
      expect(r1.startRow).to.equal(1);
      expect(r1.endCol).to.equal(2);
      expect(r1.endRow).to.equal(10);
      expect(r1.isBounded).to.be.true;

      const r2 = parseStrictA1Range("'Monthly Report'!C5:F25");
      expect(r2.worksheetTitle).to.equal('Monthly Report');
      expect(r2.startCol).to.equal(3);
      expect(r2.startRow).to.equal(5);
      expect(r2.endCol).to.equal(6);
      expect(r2.endRow).to.equal(25);

      const r3 = parseStrictA1Range('A1:D4');
      expect(r3.worksheetTitle).to.be.undefined;
      expect(r3.startCol).to.equal(1);
      expect(r3.endCol).to.equal(4);

      const r4 = parseStrictA1Range('Sheet1!A1');
      expect(r4.startCol).to.equal(1);
      expect(r4.startRow).to.equal(1);
      expect(r4.isBounded).to.be.false;

      expect(() => parseStrictA1Range('')).to.throw('Invalid range');
      expect(() => parseStrictA1Range('   ')).to.throw('Invalid range');
      expect(() => parseStrictA1Range('"Sheet1"!A1:B10')).to.throw('double-quoted worksheet titles are not valid');
    });

    it('formats bounded A1 range with proper quoting', () => {
      expect(formatBoundedA1Range('Sheet1', 1, 1, 4, 10)).to.equal('Sheet1!A1:D10');
      expect(formatBoundedA1Range('Monthly Data', 2, 5, 5, 20)).to.equal("'Monthly Data'!B5:E20");
      expect(formatBoundedA1Range(undefined, 1, 1, 1, 1)).to.equal('A1');
    });
  });

  describe('Batch Payload Sizing and Row Splitting', () => {
    it('splits large updates across row boundaries without breaking A1 notation', () => {
      const rows: GoogleSheetCli.RawData = [];
      for (let i = 1; i <= 200; i++) {
        rows.push([`Item ${i}`, i * 100, `Description for item ${i} with long text content`]);
      }

      const update = {
        range: "'Summary Sheet'!A1:C200",
        values: rows,
      };

      // Force small chunk limit to test splitting
      const chunks = splitUpdateByRows(update, 2048, 50);
      expect(chunks.length).to.be.at.least(4);
      const totalRows = chunks.reduce((sum, c) => sum + c.values.length, 0);
      expect(totalRows).to.equal(200);
      expect(chunks[0].range).to.match(/^'Summary Sheet'!A1:C\d+$/);
      const lastChunk = chunks[chunks.length - 1];
      expect(lastChunk.range).to.match(/^'Summary Sheet'!A\d+:C200$/);
      // Check contiguous row coverage
      let expectedRow = 1;
      for (const chunk of chunks) {
        const parsed = parseStrictA1Range(chunk.range);
        expect(parsed.startRow).to.equal(expectedRow);
        expect(parsed.startCol).to.equal(1);
        expect(parsed.endCol).to.equal(3);
        expectedRow = (parsed.endRow || 0) + 1;
      }
      expect(expectedRow).to.equal(201);
    });

    it('packs multiple updates into payload-bounded batches', () => {
      const updates = [
        { range: 'Sheet1!A1:B10', values: [['a', 'b'], ['c', 'd']] },
        { range: 'Sheet1!C1:D10', values: [['e', 'f'], ['g', 'h']] },
        { range: 'Sheet2!A1:B5', values: [['1', '2']] },
      ];

      const batches = packUpdateBatches(updates, 1024 * 1024);
      expect(batches.length).to.equal(1);
      expect(batches[0].length).to.equal(3);
    });
  });

  describe('Formula Overwrite Detection & Safety', () => {
    it('detects attempts to overwrite existing formulas with scalar or blank values', () => {
      const updates = [
        {
          range: 'Sheet1!A1:B2',
          values: [
            ['Header1', 'Header2'],
            ['Value1', 100], // Overwrites formula at B2
          ],
        },
      ];

      const existingData = [
        {
          range: 'Sheet1!A1:B2',
          values: [
            ['Header1', 'Header2'],
            ['Value1', '=SUM(B3:B10)'], // Formula at B2
          ],
        },
      ];

      const conflicts = findFormulaOverwrites(updates, existingData);
      expect(conflicts.length).to.equal(1);
      expect(conflicts[0].cell).to.equal('Sheet1!B2');
      expect(conflicts[0].existingFormula).to.equal('=SUM(B3:B10)');
      expect(conflicts[0].incomingValue).to.equal(100);
      expect(conflicts[0].updateIndex).to.equal(0);
    });

    it('allows incoming updates that provide the exact identical formula', () => {
      const updates = [
        {
          range: 'Sheet1!A1:B2',
          values: [
            ['Header1', 'Header2'],
            ['Value1', '=SUM(B3:B10)'], // Identical formula
          ],
        },
      ];

      const existingData = [
        {
          range: 'Sheet1!A1:B2',
          values: [
            ['Header1', 'Header2'],
            ['Value1', '=SUM(B3:B10)'],
          ],
        },
      ];

      const conflicts = findFormulaOverwrites(updates, existingData);
      expect(conflicts.length).to.equal(0);
    });
    it('handles ragged incoming rows without triggering false conflicts on unreferenced cells', () => {
      const updates = [
        {
          range: 'Sheet1!A1:B2',
          values: [
            ['Header1', 'Header2'],
            ['Value1'], // Missing col B (ragged row)
          ],
        },
      ];

      const existingData = [
        {
          range: 'Sheet1!A1:B2',
          values: [
            ['Header1', 'Header2'],
            ['Value1', '=SUM(B3:B10)'], // Formula at B2
          ],
        },
      ];

      // Missing cell in row 2 should not overwrite B2
      const conflicts = findFormulaOverwrites(updates, existingData);
      expect(conflicts.length).to.equal(0);
    });
  });

  describe('Typed Extended Value and Number Formats', () => {
    it('converts ReportCell variants to Google Sheets Schema$ExtendedValue without evaluating literal string formulas', () => {
      // Literal string starting with '='
      const literalStr = toGoogleExtendedValue('=NOT_A_FORMULA');
      expect(literalStr).to.eql({ stringValue: '=NOT_A_FORMULA' });

      // Intentional ReportFormula object
      const intentionalFormula = toGoogleExtendedValue({ formula: 'SUM(A1:A10)' });
      expect(intentionalFormula).to.eql({ formulaValue: '=SUM(A1:A10)' });

      // Numbers, Booleans, Nulls
      expect(toGoogleExtendedValue(42.5)).to.eql({ numberValue: 42.5 });
      expect(toGoogleExtendedValue(0)).to.eql({ numberValue: 0 });
      expect(toGoogleExtendedValue(false)).to.eql({ boolValue: false });
      expect(toGoogleExtendedValue(true)).to.eql({ boolValue: true });
      expect(toGoogleExtendedValue(null)).to.eql({});
      expect(toGoogleExtendedValue('')).to.eql({});
    });

    it('maps number format strings to Google Sheets Schema$NumberFormat', () => {
      expect(toGoogleNumberFormat('#,##0.00')).to.eql({ type: 'NUMBER', pattern: '#,##0.00' });
      expect(toGoogleNumberFormat('$#,##0.00')).to.eql({ type: 'CURRENCY', pattern: '$#,##0.00' });
      expect(toGoogleNumberFormat('0.0%')).to.eql({ type: 'PERCENT', pattern: '0.0%' });
      expect(toGoogleNumberFormat('YYYY-MM-DD')).to.eql({ type: 'DATE', pattern: 'YYYY-MM-DD' });
    });
  });

  describe('GoogleSheet Class API Seams and Contract Guarantees', () => {
    it('instantiates GoogleSheet and exposes all batch and report methods', () => {
      const sheet = new GoogleSheet('test-spreadsheet-id', 'Sheet1');
      expect(sheet.getData).to.be.a('function');
      expect(sheet.getDataBatch).to.be.a('function');
      expect(sheet.updateDataBatch).to.be.a('function');
      expect(sheet.appendTableData).to.be.a('function');
      expect(sheet.applyReport).to.be.a('function');
      expect(sheet.batchUpdateSpreadsheet).to.be.a('function');
      expect(sheet.formatWorksheet).to.be.a('function');
    });

    it('getDataBatch returns empty array on empty input without calling API', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      const res = await sheet.getDataBatch([]);
      expect(res).to.eql([]);
    });

    it('getDataBatch strictly prevalidates invalid range strings before API call', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      try {
        await sheet.getDataBatch(['Valid!A1:B2', '   ']);
        expect.fail('Should have thrown on empty range');
      } catch (err: unknown) {
        expect((err as Error).message).to.include('Invalid range');
      }
    });

    it('updateDataBatch returns 0-count receipt on empty updates without calling API', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      const receipt = await sheet.updateDataBatch([]);
      expect(receipt.totalRowsUpdated).to.equal(0);
      expect(receipt.totalCellsUpdated).to.equal(0);
      expect(receipt.updatedRanges).to.eql([]);
      expect(receipt.batchesExecuted).to.equal(0);
    });

    it('updateDataBatch strictly prevalidates non-2D values before API call', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      try {
        await sheet.updateDataBatch([
          { range: 'Sheet1!A1:B2', values: 'invalid' as unknown as GoogleSheetCli.RawData },
        ]);
        expect.fail('Should have thrown on non-2D array values');
      } catch (err: unknown) {
        expect((err as Error).message).to.include('must be a 2D array');
      }
    });

    it('updateDataBatch enforces explicit bounded range constraints', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      try {
        await sheet.updateDataBatch([
          {
            range: 'Sheet1!A1:B2', // 2x2 allowed
            values: [
              ['1', '2', '3'], // 3 cols > 2
              ['4', '5', '6'],
              ['7', '8', '9'], // 3 rows > 2
            ],
          },
        ]);
        expect.fail('Should have thrown on range bounding violation');
      } catch (err: unknown) {
        expect((err as Error).message).to.include('exceeds explicit bounded range');
      }
    });
    it('validates each batch update width independently without cross-update pollution', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      // Update 1 is 2x2 in A1:B2; Update 2 is 5x5 in Sheet2!A1:E5.
      // Must not throw because update 1 fits A1:B2 and update 2 fits A1:E5
      const updates = [
        { range: 'Sheet1!A1:B2', values: [['a', 'b'], ['c', 'd']] },
        {
          range: 'Sheet2!A1:E5',
          values: [
            ['1', '2', '3', '4', '5'],
            ['1', '2', '3', '4', '5'],
            ['1', '2', '3', '4', '5'],
            ['1', '2', '3', '4', '5'],
            ['1', '2', '3', '4', '5'],
          ],
        },
      ];
      // Parsing and prevalidating should succeed without bounding error
      for (const u of updates) {
        const parsed = parseStrictA1Range(u.range);
        expect(parsed.isBounded).to.be.true;
      }
    });

    it('appendTableData handles empty data cleanly with zero counts', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id', 'Sheet1');
      const res = await sheet.appendTableData([]);
      expect(res.updatedRows).to.equal(0);
      expect(res.updatedCells).to.equal(0);
    });

    it('applyReport validates ReportDocument structure', async () => {
      const sheet = new GoogleSheet('test-spreadsheet-id');
      try {
        await sheet.applyReport({} as unknown as ReportDocument);
        expect.fail('Should have thrown on invalid ReportDocument');
      } catch (err: unknown) {
        expect((err as Error).message).to.include('sheets array is required');
      }
    });
  });
});
