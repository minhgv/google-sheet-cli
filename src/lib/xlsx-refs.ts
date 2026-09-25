/**
 * Formula-reference rewriting for local XLSX splices (F7 / --updateRefs).
 *
 * A regex-class A1-ref tokenizer plus a shift engine over formula strings. This
 * is deliberately not a formula parser: only reference tokens are rewritten,
 * formulas are never evaluated. Phase-1 scope (see docs/plans
 * 2026-09-23-xlsx-multitool-roadmap.md, F7):
 *
 * - rewritten: same-sheet cell refs (A5, $B$2), ranges (A5:A20), whole-column
 *   (A:A) and whole-row (1:3) refs, with $ anchors preserved.
 * - left untouched: string literals ("..."), sheet-qualified refs ('Sheet 2'!A1
 *   and Sheet2!A1 - phase 2), defined-name-like identifiers, function names
 *   (LOG10), numbers, booleans and error literals (#REF!).
 * - refs fully inside a deleted span collapse to #REF! (Excel semantics);
 *   partially overlapping ranges shrink; refs at/after an insertion point shift
 *   by the insert count, and refs pushed past the sheet edge by an insert
 *   collapse to #REF! too (Excel refuses such inserts outright).
 *
 * Pure functions, no ExcelJS dependency.
 */

import { XlsxDimension } from './xlsx-types';

const MAX_EXCEL_COL = 16384; // XFD
const MAX_EXCEL_ROW = 1048576;

/** The splice operation formula references are rewritten against. */
export interface XlsxRefSpliceOp {
  dimension: XlsxDimension;
  /** 1-based first row/column affected */
  start: number;
  /** how many rows/columns the operation covers */
  count: number;
  mode: 'insert' | 'delete';
}

export type XlsxRefTokenKind = 'cell' | 'range' | 'column-range' | 'row-range';

/** One endpoint of a reference token. Coordinates are 1-based; the axis a token kind does not carry is 0. */
export interface XlsxRefPoint {
  col: number;
  row: number;
  colAbsolute: boolean;
  rowAbsolute: boolean;
}

/** A reference token found by the tokenizer. Only refs are emitted; everything else is skipped. */
export interface XlsxRefToken {
  kind: XlsxRefTokenKind;
  /** raw text of the token as written */
  text: string;
  /** 0-based character offset of the token inside the scanned body (leading "=" excluded) */
  index: number;
  first: XlsxRefPoint;
  /** range end; undefined for single cells */
  second?: XlsxRefPoint;
}

const CELL_REF_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$/;
const COL_REF_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_REF_RE = /^(\$?)([1-9][0-9]*)$/;
const WORD_RE = /[A-Za-z0-9_.$]+/y; // sticky: candidate refs, names, numbers and their mixtures
const ERROR_RE = /#[A-Za-z0-9./_]*[!?]?/y; // #REF!, #DIV/0!, #N/A, #NAME? ...

const colLettersToIndex = (letters: string): number => {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) & 0x1f);
  }
  return n;
};

const indexToColLetters = (idx: number): string => {
  let s = '';
  let n = idx;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

const parseCellWord = (word: string): XlsxRefPoint | null => {
  const m = CELL_REF_RE.exec(word);
  if (!m) return null;
  const col = colLettersToIndex(m[2].toUpperCase());
  if (col < 1 || col > MAX_EXCEL_COL) return null;
  const row = parseInt(m[4], 10);
  if (row > MAX_EXCEL_ROW) return null;
  return { col, row, colAbsolute: m[1] === '$', rowAbsolute: m[3] === '$' };
};

const parseColWord = (word: string): XlsxRefPoint | null => {
  const m = COL_REF_RE.exec(word);
  if (!m) return null;
  const col = colLettersToIndex(m[2].toUpperCase());
  if (col < 1 || col > MAX_EXCEL_COL) return null;
  return { col, row: 0, colAbsolute: m[1] === '$', rowAbsolute: false };
};

const parseRowWord = (word: string): XlsxRefPoint | null => {
  const m = ROW_REF_RE.exec(word);
  if (!m) return null;
  const row = parseInt(m[2], 10);
  if (row > MAX_EXCEL_ROW) return null;
  return { col: 0, row, colAbsolute: false, rowAbsolute: m[1] === '$' };
};

/** Skips a quoted literal ("..." string or '...' sheet name), honoring doubled-quote escapes. */
const skipQuoted = (s: string, start: number, quote: string): number => {
  let i = start + 1;
  while (i < s.length) {
    if (s[i] === quote) {
      if (s[i + 1] === quote) {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i++;
  }
  return s.length;
};

/**
 * Shifts an ordered [lo, hi] span on the operation axis.
 * Returns null when the span collapses to #REF!: either fully consumed by a
 * delete, or - in insert mode - shifted past the dimension's sheet edge. The
 * latter mirrors Excel refusing the insert outright; serializing an
 * out-of-bounds ref like A1048577 would corrupt the file silently.
 */
const shiftSpan = (lo: number, hi: number, op: XlsxRefSpliceOp): { lo: number; hi: number } | null => {
  const s = op.start;
  const n = op.count;
  const spanEnd = s + n - 1;
  if (op.mode === 'insert') {
    const limit = op.dimension === 'ROWS' ? MAX_EXCEL_ROW : MAX_EXCEL_COL;
    const nlo = lo >= s ? lo + n : lo;
    const nhi = hi >= s ? hi + n : hi;
    if (nhi > limit) return null;
    return { lo: nlo, hi: nhi };
  }
  const nlo = lo < s ? lo : lo > spanEnd ? lo - n : s;
  const nhi = hi < s ? hi : hi > spanEnd ? hi - n : s - 1;
  if (nlo > nhi) return null;
  return { lo: nlo, hi: nhi };
};

/**
 * Renders one endpoint: `$?letters` and/or `$?row` per the point's anchors.
 * Column letters keep the casing of the reference as written ("a1" stays lowercase);
 * style selects which parts an endpoint kind carries.
 */
const renderRefPoint = (p: XlsxRefPoint, raw: string, col: number, row: number, style: 'cell' | 'column' | 'row'): string => {
  let out = '';
  if (style !== 'row') {
    const m = /[A-Za-z]+/.exec(raw);
    const letters = indexToColLetters(col);
    out += `${p.colAbsolute ? '$' : ''}${m && /^[a-z]+$/.test(m[0]) ? letters.toLowerCase() : letters}`;
  }
  if (style !== 'column') {
    out += `${p.rowAbsolute ? '$' : ''}${row}`;
  }
  return out;
};

/** Renders one token against the splice op; null means the ref collapsed to #REF!. */
const renderToken = (t: XlsxRefToken, op: XlsxRefSpliceOp): string | null => {
  const byRows = op.dimension === 'ROWS';
  const raws = t.text.split(':');
  const firstRaw = raws[0] ?? '';
  const secondRaw = raws[1] ?? firstRaw;

  if (t.kind === 'column-range' || t.kind === 'row-range') {
    if (!t.second) return t.text;
    const onAxis = t.kind === 'column-range' ? !byRows : byRows;
    if (!onAxis) return t.text; // whole-column refs ignore ROWS ops and vice versa
    const lo = Math.min(t.first.col || t.first.row, t.second.col || t.second.row);
    const hi = Math.max(t.first.col || t.first.row, t.second.col || t.second.row);
    const shifted = shiftSpan(lo, hi, op);
    if (!shifted) return null;
    const first = t.kind === 'column-range'
      ? renderRefPoint(t.first, firstRaw, shifted.lo, 0, 'column')
      : renderRefPoint(t.first, firstRaw, 0, shifted.lo, 'row');
    const second = t.kind === 'column-range'
      ? renderRefPoint(t.second, secondRaw, shifted.hi, 0, 'column')
      : renderRefPoint(t.second, secondRaw, 0, shifted.hi, 'row');
    return `${first}:${second}`;
  }

  // cell or cell range: rewrite the row on ROWS ops, the column on COLUMNS ops
  const endpoints = [t.first, ...(t.second ? [t.second] : [])].map((p, idx) => ({
    p,
    raw: idx === 0 ? firstRaw : secondRaw,
    v: byRows ? p.row : p.col,
  }));
  endpoints.sort((a, b) => a.v - b.v);
  const shifted = shiftSpan(endpoints[0].v, endpoints[endpoints.length - 1].v, op);
  if (!shifted) return null;

  const rendered = endpoints.map(({ p, raw }, idx) => {
    const value = idx === 0 ? shifted.lo : shifted.hi;
    return byRows ? renderRefPoint(p, raw, p.col, value, 'cell') : renderRefPoint(p, raw, value, p.row, 'cell');
  });
  return t.second ? `${rendered[0]}:${rendered[1]}` : rendered[0];
};

/**
 * Scans a formula body and emits only reference tokens. String literals,
 * sheet-qualified references ('Sheet 2'!A1, Sheet2!A1:B2 - the whole qualified
 * ref including a range partner is skipped), identifiers, function names,
 * numbers, booleans and error literals produce no tokens.
 */
export function tokenizeFormulaRefs(formula: string): XlsxRefToken[] {
  const tokens: XlsxRefToken[] = [];
  const len = formula.length;
  let i = 0;
  // set after a Sheet! prefix: the next ref belongs to another sheet and is skipped whole
  let skipQualifiedRef = false;

  const wordAt = (pos: number): { word: string; end: number } | null => {
    WORD_RE.lastIndex = pos;
    const m = WORD_RE.exec(formula);
    return m ? { word: m[0], end: m.index + m[0].length } : null;
  };

  while (i < len) {
    const c = formula[i];

    if (c === '"') {
      i = skipQuoted(formula, i, '"');
      continue;
    }
    if (c === "'") {
      const end = skipQuoted(formula, i, "'");
      if (formula[end] === '!') {
        i = end + 1;
        skipQualifiedRef = true;
      } else {
        i = end;
      }
      continue;
    }
    if (c === '#') {
      ERROR_RE.lastIndex = i;
      const m = ERROR_RE.exec(formula);
      i = m ? m.index + m[0].length : i + 1;
      continue;
    }
    if (!/[A-Za-z0-9_.$]/.test(c)) {
      i++;
      continue;
    }

    const w = wordAt(i);
    if (!w) {
      i++;
      continue;
    }

    if (skipQualifiedRef) {
      // leave the sheet-qualified ref untouched, including a ":ref" range partner
      skipQualifiedRef = false;
      let end = w.end;
      if (formula[end] === ':') {
        const partner = wordAt(end + 1);
        if (partner && formula[partner.end] !== '!') end = partner.end;
      }
      i = end;
      continue;
    }

    // an unquoted sheet name prefix (Sheet2!A1) - even when the sheet name itself
    // looks like a cell reference (A1!B2 is a valid cross-sheet form)
    if (formula[w.end] === '!') {
      i = w.end + 1;
      skipQualifiedRef = true;
      continue;
    }

    // a word followed by "(" is a function call, never a reference - this is what
    // keeps LOG10( from parsing as cell LOG10 (column LOG, row 10)
    if (formula[w.end] === '(') {
      i = w.end;
      continue;
    }

    const cell = parseCellWord(w.word);
    const col = parseColWord(w.word);
    const row = parseRowWord(w.word);

    if (formula[w.end] === ':') {
      const partnerW = wordAt(w.end + 1);
      if (partnerW && formula[partnerW.end] !== '!') {
        const partnerCell = parseCellWord(partnerW.word);
        if (cell && partnerCell) {
          tokens.push({ kind: 'range', text: formula.slice(i, partnerW.end), index: i, first: cell, second: partnerCell });
          i = partnerW.end;
          continue;
        }
        const partnerCol = parseColWord(partnerW.word);
        if (col && partnerCol) {
          tokens.push({ kind: 'column-range', text: formula.slice(i, partnerW.end), index: i, first: col, second: partnerCol });
          i = partnerW.end;
          continue;
        }
        const partnerRow = parseRowWord(partnerW.word);
        if (row && partnerRow) {
          tokens.push({ kind: 'row-range', text: formula.slice(i, partnerW.end), index: i, first: row, second: partnerRow });
          i = partnerW.end;
          continue;
        }
      }
    }

    if (cell) {
      tokens.push({ kind: 'cell', text: w.word, index: i, first: cell });
    }
    // bare column/row words without a colon partner are numbers or names, not refs
    i = w.end;
  }

  return tokens;
}

export interface XlsxFormulaRefsResult {
  /** the rewritten formula (leading "=" preserved when present) */
  formula: string;
  /** true when the rewrite changed the formula text */
  changed: boolean;
  /** how many references collapsed to #REF! */
  refErrors: number;
}

/**
 * Rewrites the same-sheet A1 references in one formula string against a
 * splice operation. Sheet-qualified refs and everything that is not a
 * reference are left untouched; $ anchors and letter casing are preserved;
 * refs fully inside a deleted span become #REF!.
 */
export function rewriteFormulaRefs(formula: string, op: XlsxRefSpliceOp): XlsxFormulaRefsResult {
  if (!formula || typeof formula !== 'string') {
    return { formula, changed: false, refErrors: 0 };
  }
  const hasLeadingEquals = formula.startsWith('=');
  const body = hasLeadingEquals ? formula.slice(1) : formula;
  const tokens = tokenizeFormulaRefs(body);
  if (tokens.length === 0) {
    return { formula, changed: false, refErrors: 0 };
  }

  let refErrors = 0;
  const pieces: string[] = [];
  let cursor = 0;
  for (const t of tokens) {
    pieces.push(body.slice(cursor, t.index));
    const rendered = renderToken(t, op);
    if (rendered === null) refErrors++;
    pieces.push(rendered ?? '#REF!');
    cursor = t.index + t.text.length;
  }
  pieces.push(body.slice(cursor));

  const out = (hasLeadingEquals ? '=' : '') + pieces.join('');
  return { formula: out, changed: out !== formula, refErrors };
}

export interface XlsxDefinedNameDiff {
  /** defined-name ranges that were moved, grown or shrunk by the splice */
  rewritten: number;
  /** defined-name ranges that were removed because they fell inside the deleted span */
  broken: number;
}

/**
 * Diffs defined-name range strings before/after a splice. Names are keyed by
 * name; a range string that changed form counts as rewritten, and a shrinkage
 * of a name's range count counts as broken (the removed range fell inside the
 * deleted span). Names on other sheets keep identical strings and therefore
 * never contribute.
 */
export function diffDefinedNameRanges(
  before: Map<string, string[]>,
  after: Map<string, string[]>
): XlsxDefinedNameDiff {
  let rewritten = 0;
  let broken = 0;
  for (const [name, beforeRanges] of before) {
    const afterRanges = after.get(name);
    if (!afterRanges || afterRanges.length === 0) {
      broken += beforeRanges.length;
      continue;
    }
    const beforeSet = new Set(beforeRanges);
    for (const r of afterRanges) {
      if (!beforeSet.has(r)) rewritten++;
    }
    broken += Math.max(0, beforeRanges.length - afterRanges.length);
  }
  return { rewritten, broken };
}
