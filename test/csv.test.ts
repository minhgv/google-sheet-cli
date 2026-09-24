import { expect } from 'chai';
import { parse as parseCsvSync } from 'csv-parse/sync';
import {
  CSV_DEFAULTS,
  CSV_INJECTION_POLICIES,
  CsvCell,
  CsvNewline,
  CsvOptions,
  csvEscapeFormulaText,
  serializeCsv,
} from '../src/lib/csv';

/**
 * The deterministic CSV serializer and its injection policy (T-07, AC-06/AC-07).
 *
 * These pins are the byte-level contract `data:export-csv` (T-08) is built on: same input,
 * same bytes; RFC 4180 quoting and CRLF records; blanks and nulls as empty fields; typed
 * numbers and booleans rendered by String(value) with no rounding; dangerous strings
 * prefixed under `safe` and byte-faithful under `preserve`; ragged or unrepresentable
 * input refused, never padded. A round-trip through csv-parse (the repo's own reader)
 * proves a real consumer reads the bytes back.
 */
describe('csv serializer', () => {
  const CRLF = '\r\n';

  describe('defaults and locked vocabulary', () => {
    it('defaults to comma, CRLF, no BOM and the safe policy', () => {
      expect(CSV_DEFAULTS).to.deep.equal({ delimiter: ',', newline: '\r\n', bom: false, injectionPolicy: 'safe' });
    });

    it('locks the injection policy vocabulary', () => {
      expect(CSV_INJECTION_POLICIES).to.eql(['safe', 'preserve']);
    });

    it('resolves the default policy into the receipt', () => {
      expect(serializeCsv([['a']]).receipt.policy).to.equal('safe');
      expect(serializeCsv([['a']], { injectionPolicy: 'preserve' }).receipt.policy).to.equal('preserve');
    });

    it('never invents a header row', () => {
      const { csv } = serializeCsv([
        ['h1', 'h2'],
        ['v1', 'v2'],
      ]);
      expect(csv).to.equal(`h1,h2${CRLF}v1,v2${CRLF}`);
    });
  });

  describe('determinism', () => {
    const rows: CsvCell[][] = [
      ['id', 'note'],
      ['=a,b', 'line1\nline2'],
      [-12.5, null],
      ['héllo 🌍', '"quoted"'],
    ];

    it('produces byte-identical output for the same input', () => {
      const first = serializeCsv(rows);
      const second = serializeCsv(rows);
      expect(first.csv).to.equal(second.csv);
      expect(Buffer.from(first.csv, 'utf8').equals(Buffer.from(second.csv, 'utf8'))).to.be.true;
      expect(first.receipt).to.deep.equal(second.receipt);
    });

    it('serializes the machine-readable receipt identically too', () => {
      expect(JSON.stringify(serializeCsv(rows).receipt)).to.equal(JSON.stringify(serializeCsv(rows).receipt));
    });

    it('round-trips through csv-parse, the repo’s own CSV reader', () => {
      // preserve: byte fidelity - a plain consumer reads back exactly what went in
      const preserved = serializeCsv(rows, { injectionPolicy: 'preserve' });
      expect(parseCsvSync(preserved.csv) as unknown).to.eql([
        ['id', 'note'],
        ['=a,b', 'line1\nline2'],
        ['-12.5', ''],
        ['héllo 🌍', '"quoted"'],
      ]);

      // safe: the neutralization is what a consumer actually reads back
      const safe = serializeCsv(rows);
      expect(parseCsvSync(safe.csv) as unknown).to.eql([
        ['id', 'note'],
        ["'=a,b", 'line1\nline2'],
        ['-12.5', ''],
        ['héllo 🌍', '"quoted"'],
      ]);
      expect(safe.receipt.transformedCells).to.equal(1);
    });
  });

  describe('quoting and escaping (RFC 4180)', () => {
    it('doubles internal quotes and wraps the field', () => {
      expect(serializeCsv([['has "quote"']]).csv).to.equal(`"has ""quote"""${CRLF}`);
    });

    it('quotes fields containing the delimiter', () => {
      expect(serializeCsv([['a,b', 'plain']]).csv).to.equal(`"a,b",plain${CRLF}`);
    });

    it('quotes multiline fields without translating their newlines', () => {
      expect(serializeCsv([['line1\nline2', 'x']]).csv).to.equal(`"line1\nline2",x${CRLF}`);
      expect(serializeCsv([['a\rb', 'x']]).csv).to.equal(`"a\rb",x${CRLF}`);
    });

    it('leaves plain fields unquoted', () => {
      expect(serializeCsv([['plain text 123', 'x-y_z']]).csv).to.equal(`plain text 123,x-y_z${CRLF}`);
    });

    it('keeps unicode byte-faithful', () => {
      const { csv } = serializeCsv([['héllo 🌍', '日本語']]);
      expect(csv).to.equal(`héllo 🌍,日本語${CRLF}`);
      expect(Buffer.from(csv, 'utf8').toString('utf8')).to.equal(csv);
    });
  });

  describe('blanks, nulls and typed values', () => {
    it('renders empty strings and nulls as empty fields', () => {
      expect(serializeCsv([['', null, 'x']]).csv).to.equal(`,,x${CRLF}`);
      expect(serializeCsv([[null]]).csv).to.equal(CRLF);
    });

    it('renders numbers via String(value) with no rounding', () => {
      expect(serializeCsv([[0, -1, 0.30000000000000004, 1e21]]).csv).to.equal(`0,-1,0.30000000000000004,1e+21${CRLF}`);
    });

    it('renders booleans via String(value), never TRUE/FALSE', () => {
      expect(serializeCsv([[true, false]]).csv).to.equal(`true,false${CRLF}`);
    });

    it('never quotes numbers or booleans', () => {
      const { csv } = serializeCsv([[-1, true]]);
      expect(csv).to.equal(`-1,true${CRLF}`);
      expect(csv.includes('"')).to.be.false;
    });
  });

  describe('injection policy safe (default)', () => {
    it('prefixes each dangerous leading character with an apostrophe', () => {
      const cases: [string, string][] = [
        ['=SUM(A1)', `'=SUM(A1)${CRLF}`],
        ['+48123', `'+48123${CRLF}`],
        ['-note', `'-note${CRLF}`],
        ['@cmd', `'@cmd${CRLF}`],
        ['\tindented', `'\tindented${CRLF}`],
        ['\rcarriage', `"'\rcarriage"${CRLF}`], // quoted too: still contains CR after the prefix
      ];
      for (const [input, expected] of cases) {
        const { csv, receipt } = serializeCsv([[input]]);
        expect(csv, JSON.stringify(input)).to.equal(expected);
        expect(receipt.transformedCells, JSON.stringify(input)).to.equal(1);
      }
    });

    it('applies the prefix before the quoting decision', () => {
      expect(serializeCsv([['=a,b']]).csv).to.equal(`"'=a,b"${CRLF}`);
    });

    it('keeps typed negative numbers numeric - the policy applies to string cells only', () => {
      const { csv, receipt } = serializeCsv([[-1, '-1', '+2', 3 - 4]]);
      expect(csv).to.equal(`-1,'-1,'+2,-1${CRLF}`);
      expect(receipt.transformedCells).to.equal(2);
    });

    it('prefixes formula strings - formula text in CSV is the dangerous case', () => {
      expect(serializeCsv([['=SUM(A1:B2)']]).csv).to.equal(`'=SUM(A1:B2)${CRLF}`);
      expect(serializeCsv([['=SUM(1,2)']]).csv).to.equal(`"'=SUM(1,2)"${CRLF}`);
    });

    it('leaves safe strings untouched and counts nothing', () => {
      const { csv, receipt } = serializeCsv([['plain', 'a=b mid-string', 'x\t', '']]);
      expect(csv).to.equal(`plain,a=b mid-string,x\t,${CRLF}`);
      expect(receipt.transformedCells).to.equal(0);
      expect(receipt.warnings).to.eql([]);
    });
  });

  describe('injection policy preserve', () => {
    it('emits dangerous cells byte-faithfully and warns with a count', () => {
      const { csv, receipt } = serializeCsv([['=SUM(A1)']], { injectionPolicy: 'preserve' });
      expect(csv).to.equal(`=SUM(A1)${CRLF}`);
      expect(receipt.transformedCells).to.equal(0);
      expect(receipt.warnings).to.have.lengthOf(1);
      expect(receipt.warnings[0]).to.include('preserve');
      expect(receipt.warnings[0]).to.include('1 string cell(s)');
    });

    it('counts every dangerous cell without leaking cell contents into the receipt', () => {
      const { csv, receipt } = serializeCsv([['=a', 'b', '@c']], { injectionPolicy: 'preserve' });
      expect(csv).to.equal(`=a,b,@c${CRLF}`);
      expect(receipt.warnings.join(' ')).to.include('2 string cell(s)');
      expect(receipt.warnings.join(' ')).to.not.include('=a');
      expect(receipt.warnings.join(' ')).to.not.include('@c');
    });

    it('stays silent when nothing dangerous passed through', () => {
      const { receipt } = serializeCsv([['plain', 1, null]], { injectionPolicy: 'preserve' });
      expect(receipt.warnings).to.eql([]);
      expect(receipt.transformedCells).to.equal(0);
    });
  });

  describe('receipt', () => {
    it('records row and column counts', () => {
      const { receipt } = serializeCsv([
        [1, 2],
        [3, 4],
        [5, 6],
      ]);
      expect(receipt.rows).to.equal(3);
      expect(receipt.columns).to.equal(2);
    });

    it('reports zero rows and columns for empty input', () => {
      const { csv, receipt } = serializeCsv([]);
      expect(csv).to.equal('');
      expect(receipt.rows).to.equal(0);
      expect(receipt.columns).to.equal(0);
    });

    it('reports a zero-column single empty row', () => {
      const { csv, receipt } = serializeCsv([[]]);
      expect(csv).to.equal(CRLF);
      expect(receipt.rows).to.equal(1);
      expect(receipt.columns).to.equal(0);
    });

    it('returns an immutable result - callers build on it, never into it', () => {
      const result = serializeCsv([['=x']]);
      expect(Object.isFrozen(result)).to.be.true;
      expect(Object.isFrozen(result.receipt)).to.be.true;
      expect(Object.isFrozen(result.receipt.warnings)).to.be.true;
    });
  });

  describe('newlines and encoding', () => {
    it('terminates every record with CRLF by default, including the last', () => {
      const { csv } = serializeCsv([
        ['a', 'b'],
        ['c', 'd'],
      ]);
      expect(csv).to.equal(`a,b${CRLF}c,d${CRLF}`);
      expect((csv.match(/\r\n/g) || []).length).to.equal(2);
    });

    it('honors newline "\\n" for LF-only pipelines', () => {
      const { csv } = serializeCsv(
        [
          ['a', 'b'],
          ['c', 'd'],
        ],
        { newline: '\n' }
      );
      expect(csv).to.equal('a,b\nc,d\n');
      expect(csv.includes('\r')).to.be.false;
    });

    it('emits no BOM unless asked', () => {
      expect(serializeCsv([['a']]).csv.startsWith('\uFEFF')).to.be.false;
    });

    it('prepends a UTF-8 BOM when bom is set, leaving the body byte-identical', () => {
      const plain = serializeCsv([['a', 'b']]).csv;
      const { csv } = serializeCsv([['a', 'b']], { bom: true });
      expect(csv.startsWith('\uFEFF')).to.be.true;
      expect(csv.slice(1)).to.equal(plain);
    });

    it('emits a lone BOM for empty input', () => {
      expect(serializeCsv([], { bom: true }).csv).to.equal('\uFEFF');
    });
  });

  describe('rectangular enforcement', () => {
    it('refuses ragged rows instead of padding', () => {
      expect(() => serializeCsv([[1, 2], [3]])).to.throw(
        TypeError,
        'Ragged CSV input: row 1 has 1 cell(s), expected 2'
      );
    });

    it('refuses rows wider than the first row', () => {
      expect(() => serializeCsv([['a'], ['b', 'c']])).to.throw(
        TypeError,
        'Ragged CSV input: row 1 has 2 cell(s), expected 1'
      );
    });

    it('throws Error instances per repo convention', () => {
      try {
        serializeCsv([[1], [2, 3]]);
        expect.fail('should have thrown');
      } catch (err) {
        expect(err).to.be.an.instanceof(Error);
      }
    });
  });

  describe('deterministic refusals', () => {
    const expectRefusal = (fn: () => unknown, fragment: string): void => {
      expect(fn).to.throw(TypeError);
      expect(fn).to.throw(fragment);
    };

    it('refuses delimiters that are not one safe character', () => {
      expectRefusal(() => serializeCsv([['a']], { delimiter: '' }), 'must be exactly one character');
      expectRefusal(() => serializeCsv([['a']], { delimiter: ';;' }), 'must be exactly one character');
      expectRefusal(() => serializeCsv([['a']], { delimiter: '"' }), 'cannot be a quote, CR or LF');
      expectRefusal(() => serializeCsv([['a']], { delimiter: '\n' }), 'cannot be a quote, CR or LF');
    });

    it('refuses unknown record separators', () => {
      expectRefusal(() => serializeCsv([['a']], { newline: 'x' as unknown as CsvNewline }), 'newline');
      expectRefusal(() => serializeCsv([['a']], { newline: '' as unknown as CsvNewline }), 'newline');
    });

    it('refuses policies outside the locked vocabulary', () => {
      expectRefusal(
        () => serializeCsv([['a']], { injectionPolicy: 'Safe' as unknown as CsvOptions['injectionPolicy'] }),
        'injection policy'
      );
    });

    it('refuses non-boolean bom flags', () => {
      expectRefusal(() => serializeCsv([['a']], { bom: 'yes' as unknown as boolean }), 'bom flag');
    });

    it('refuses non-finite numbers instead of writing ambiguous text', () => {
      expectRefusal(() => serializeCsv([[Number.NaN]]), 'Non-finite number at CSV row 0 column 0');
      expectRefusal(() => serializeCsv([[1, Number.POSITIVE_INFINITY]]), 'Non-finite number at CSV row 0 column 1');
    });

    it('refuses unsupported cell types, pointing undefined at null', () => {
      expectRefusal(() => serializeCsv([[undefined as unknown as CsvCell]]), 'undefined (use null for an empty field)');
      expectRefusal(() => serializeCsv([[{ v: 1 } as unknown as CsvCell]]), 'Unsupported CSV cell at row 0 column 0');
    });

    it('refuses non-array rows and rows', () => {
      expectRefusal(() => serializeCsv('nope' as unknown as CsvCell[][]), 'rows');
      expectRefusal(() => serializeCsv([[1], 'x' as unknown as CsvCell[]]), 'row 1');
      expectRefusal(() => serializeCsv([['a']], null as unknown as CsvOptions), 'options');
    });
  });

  describe('csvEscapeFormulaText', () => {
    it('prefixes dangerous formula text for callers rendering outside serializeCsv', () => {
      expect(csvEscapeFormulaText('=SUM(A1)')).to.equal("'=SUM(A1)");
      expect(csvEscapeFormulaText('-5')).to.equal("'-5");
      expect(csvEscapeFormulaText('@x')).to.equal("'@x");
    });

    it('returns plain text unchanged', () => {
      expect(csvEscapeFormulaText('plain')).to.equal('plain');
      expect(csvEscapeFormulaText('')).to.equal('');
      expect(csvEscapeFormulaText('a=b later')).to.equal('a=b later');
    });

    it('agrees with what serializeCsv does to the same string as data', () => {
      const text = '=SUM(A1:B2)';
      expect(serializeCsv([[csvEscapeFormulaText(text)]]).csv).to.equal(serializeCsv([[text]]).csv);
    });
  });
});
