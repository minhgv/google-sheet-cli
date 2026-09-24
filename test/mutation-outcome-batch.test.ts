import { expect } from 'chai';
import { PassThrough } from 'stream';
import https from 'https';
import GoogleSheet, { GoogleSheetCli, buildBatchUpdateMutationOutcome } from '../src/lib/google-sheet';
import { GSheetError } from '../src/lib/cli-errors';
import { serializeMutationOutcome } from '../src/lib/mutation-outcome';
import { FakeSheets } from './fake-sheets';

/**
 * T-05: updateDataBatch reports typed partial mutation outcomes on top of the locked
 * vocabulary in src/lib/mutation-outcome.ts. Message text stays byte-identical to the plain
 * `Error`s previously thrown (the regression pins); the report rides on
 * `error.mutationOutcome` with the original transport error as `cause`.
 *
 * Faults are injected through the seams FakeSheets already exposes (`failRequests`) plus a
 * local `https.request` wrapper for the ordinal and lost-response cases fake-sheets.ts cannot
 * express - fake-sheets.ts itself is not edited. The `cause` message is whatever the client
 * stack surfaces for the injected body, so the wrapper-text assertions build their expected
 * string from `error.cause`: byte-identical relative to the actual cause, without pinning
 * gaxios' own error text.
 */

const SPREADSHEET_ID = 'fake-spreadsheet-id';
const SPREADSHEET_TITLE = 'Mutation outcome spreadsheet';
const SHEET = 'Data';

/**
 * The client stack expects the handful of ClientRequest members the fake also stubs out; a
 * PassThrough widened with exactly those is the smallest thing gaxios/node-fetch accepts.
 */
type FaultRequest = PassThrough & {
  abort: () => undefined;
  setTimeout: (...args: unknown[]) => undefined;
  setNoDelay: () => undefined;
  setSocketKeepAlive: () => undefined;
  flushHeaders: () => undefined;
};

/** The response half of a fabricated answer: status line plus the headers node-fetch reads. */
type FaultResponse = PassThrough & {
  statusCode: number;
  statusMessage: string;
  headers: Record<string, string>;
  rawHeaders: string[];
};

// Node types the core https module's `request` as readonly overloads; the fault wrappers need
// the mutable runtime binding, so the type is widened once here.
const httpsModule = https as unknown as { request: (...args: unknown[]) => unknown };

const faultRequest = (): FaultRequest =>
  Object.assign(new PassThrough(), {
    abort: () => undefined,
    setTimeout: () => undefined,
    setNoDelay: () => undefined,
    setSocketKeepAlive: () => undefined,
    flushHeaders: () => undefined,
  });

const faultResponse = (status: number, body: unknown): FaultResponse => {
  const payload = Buffer.from(JSON.stringify(body), 'utf8');
  const contentType = 'application/json; charset=UTF-8';
  return Object.assign(new PassThrough(), {
    statusCode: status,
    statusMessage: 'Bad Request',
    headers: { 'content-type': contentType, 'content-length': String(payload.length) },
    rawHeaders: ['content-type', contentType, 'content-length', String(payload.length)],
  });
};

/** Best-effort URL string from either https.request signature. */
const requestUrl = (args: unknown[]): string => {
  const first = args[0];
  if (typeof first === 'string') return first;
  if (first instanceof URL) return first.href;
  if (first !== null && typeof first === 'object' && 'path' in first) {
    const path = first.path;
    return `https:/${typeof path === 'string' ? path : ''}`;
  }
  return '';
};

const causeMessage = (err: unknown): string => {
  const cause = (err as { cause?: unknown }).cause;
  return cause instanceof Error ? cause.message : String(cause);
};

/**
 * Run a call that is expected to reject and hand back whatever it threw
 *
 * @param {() => Promise<unknown>} fn
 * @returns {Promise<unknown>}
 */
const rejection = async (fn: () => Promise<unknown>): Promise<unknown> => {
  try {
    return await fn();
  } catch (error) {
    return error;
  }
  throw new Error('expected the call to reject, but it resolved');
};

/**
 * Wrap the fake's installed https.request: requests whose URL matches die with a raw
 * transport error (no HTTP response at all), everything else flows through the fake. Returns
 * an undo function.
 *
 * @param {(url: string) => boolean} match
 * @returns {() => void}
 */
const dropMatchingRequests = (match: (url: string) => boolean): (() => void) => {
  const original = httpsModule.request;
  const patched = function (this: unknown, ...args: unknown[]) {
    if (match(requestUrl(args))) {
      const req = faultRequest();
      req.on('data', () => undefined);
      req.on('end', () => undefined);
      setImmediate(() => req.emit('error', Object.assign(new Error('socket hang up'), { code: 'ECONNRESET' })));
      return req;
    }
    return original.apply(this, args);
  };
  httpsModule.request = patched;
  return () => {
    httpsModule.request = original;
  };
};

/**
 * Wrap the fake's installed https.request: the nth request whose URL contains `pathPart` is
 * answered with a fabricated HTTP error response; every other request flows through. Returns
 * an undo function.
 *
 * @param {number} n - 1-based ordinal of the matching request to reject
 * @param {string} pathPart
 * @param {number} status
 * @returns {() => void}
 */
const rejectNthRequest = (n: number, pathPart: string, status: number): (() => void) => {
  const original = httpsModule.request;
  let seen = 0;
  const patched = function (this: unknown, ...args: unknown[]) {
    const callback = [...args].find((argument): argument is (...cbArgs: unknown[]) => void => typeof argument === 'function');
    if (requestUrl(args).includes(pathPart)) {
      seen += 1;
      if (seen === n) {
        const req = faultRequest();
        req.on('data', () => undefined);
        req.on('end', () => {
          const body = { error: { code: status, message: `Injected rejection #${n} matching "${pathPart}"`, status: 'FAILED_PRECONDITION' } };
          const res = faultResponse(status, body);
          if (callback) callback(res);
          req.emit('response', res);
          res.end(Buffer.from(JSON.stringify(body), 'utf8'));
        });
        return req;
      }
    }
    return original.apply(this, args);
  };
  httpsModule.request = patched;
  return () => {
    httpsModule.request = original;
  };
};

describe('updateDataBatch typed mutation outcomes (T-05)', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  /**
   * Recreate the fake spreadsheet and a client authorized against it
   *
   * @param {number} rowCount
   * @returns {Promise<void>}
   */
  const setupGrid = async (rowCount: number): Promise<void> => {
    fake.reset();
    fake.addSpreadsheet(SPREADSHEET_ID, SPREADSHEET_TITLE, [{ title: SHEET, rowCount, columnCount: 5 }]);
    gsheet = new GoogleSheet(SPREADSHEET_ID);
    await gsheet.authorize(fake.credentials);
  };

  before(() => fake.install());
  after(() => fake.uninstall());
  beforeEach(async () => setupGrid(100));

  it('reports a refused first chunk rejected, the tail not-attempted, and replay as safe', async () => {
    fake.failRequests('values:batchUpdate', 400, 1);
    const error = (await rejection(() =>
      gsheet.updateDataBatch([{ range: `${SHEET}!A1:B2`, values: [['a', 'b'], ['c', 'd']] }])
    )) as GSheetError;

    expect(error).to.be.instanceOf(GSheetError);
    expect(error.code).to.equal('REQUEST_INVALID');
    expect(error.retryable).to.equal(false);
    expect(error.message).to.equal(
      `Batch update failed at batch 1/1 [${SHEET}!A1:B2]: ${causeMessage(error)}. Successfully completed ranges: none`
    );
    expect((error as { cause?: unknown }).cause).to.be.an('error');

    const outcome = error.mutationOutcome!;
    expect(outcome.operation).to.equal('updateDataBatch');
    expect(outcome.phase).to.equal('values-write');
    expect(outcome.requests).to.have.lengthOf(1);
    expect(outcome.requests[0].kind).to.equal('values.batchUpdate');
    expect(outcome.requests[0].state).to.equal('rejected');
    expect(outcome.requests[0].httpStatus).to.equal(400);
    expect(outcome.requests[0].a1Ranges).to.eql([`${SHEET}!A1:B2`]);
    expect(outcome.retryGuidance).to.equal('safe-replay');
    expect(outcome.retryGuidanceReason).to.be.a('string').that.is.not.empty;
    expect(outcome.gridGrowth).to.deep.equal({ attempted: false, completed: false });

    // the response proved non-application: nothing was written
    expect(fake.cell(SPREADSHEET_ID, SHEET, 'A1')).to.equal('');
  });

  it('keeps the acknowledged chunk acknowledged when a later chunk is refused', async () => {
    const undo = rejectNthRequest(2, '/values:batchUpdate', 400);
    try {
      const error = (await rejection(() =>
        gsheet.updateDataBatch(
          [
            { range: `${SHEET}!A1:B1`, values: [['a1', 'b1']] },
            { range: `${SHEET}!A2:B2`, values: [['a2', 'b2']] },
          ],
          { chunkByteSize: 1 }
        )
      )) as GSheetError;

      expect(error.code).to.equal('REQUEST_INVALID');
      expect(error.message).to.equal(
        `Batch update failed at batch 2/2 [${SHEET}!A2:B2]: ${causeMessage(error)}. Successfully completed ranges: ${SHEET}!A1:B1`
      );

      const outcome = error.mutationOutcome!;
      expect(outcome.requests.map((request) => request.state)).to.eql(['acknowledged', 'rejected']);
      expect(outcome.requests[0].a1Ranges).to.eql([`${SHEET}!A1:B1`]);
      expect(outcome.requests[1].httpStatus).to.equal(400);
      expect(outcome.retryGuidance).to.equal('verify-then-replay');
      expect(outcome.retryGuidanceReason).to.match(/[Vv]erify/);

      // the serialized projection counts both states for envelope consumers
      const serialized = serializeMutationOutcome(outcome);
      expect(serialized.summary).to.deep.equal({ total: 2, acknowledged: 1, rejected: 1, unknown: 0, notAttempted: 0 });

      // the fake really holds chunk one and nothing of chunk two
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A1')).to.equal('a1');
      expect(fake.cell(SPREADSHEET_ID, SHEET, 'A2')).to.equal('');
    } finally {
      undo();
    }
  });

  it('records grid growth completed when only the value write is refused', async () => {
    await setupGrid(3);
    fake.failRequests('values:batchUpdate', 400, 1);
    const error = (await rejection(() =>
      gsheet.updateDataBatch([{ range: `${SHEET}!A4:A5`, values: [['x'], ['y']] }])
    )) as GSheetError;

    // the growth really applied in the backend before the write was refused
    expect(fake.worksheet(SPREADSHEET_ID, SHEET).rowCount).to.be.at.least(5);

    expect(error.message).to.match(new RegExp(`^Batch update failed at batch 1/1 \\[${SHEET}!A4:A5\\]`));
    const outcome = error.mutationOutcome!;
    expect(outcome.phase).to.equal('values-write');
    expect(outcome.gridGrowth).to.deep.equal({
      attempted: true,
      completed: true,
      details: '1 appendDimension request(s) applied',
    });
    expect(outcome.requests[0].state).to.equal('rejected');
    expect(outcome.requests[0].httpStatus).to.equal(400);
    expect(outcome.retryGuidance).to.equal('safe-replay');
  });

  it('keeps every value chunk not-attempted when the growth request itself fails', async () => {
    await setupGrid(3);
    fake.failRequests(`${SPREADSHEET_ID}:batchUpdate`, 500, 1);
    const error = (await rejection(() =>
      gsheet.updateDataBatch([{ range: `${SHEET}!A4:A5`, values: [['x'], ['y']] }])
    )) as GSheetError;

    expect(error).to.be.instanceOf(GSheetError);
    expect(error.code).to.equal('UPSTREAM');
    expect(error.retryable).to.equal(true);
    // this path adds no wrapper prose: the raw upstream message is the whole message
    expect(error.message).to.equal(causeMessage(error));

    const outcome = error.mutationOutcome!;
    expect(outcome.phase).to.equal('grid-growth');
    expect(outcome.gridGrowth).to.deep.equal({
      attempted: true,
      completed: false,
      details: '1 appendDimension request(s) not confirmed',
    });
    expect(outcome.requests.map((request) => request.state)).to.eql(['not-attempted']);
    expect(outcome.retryGuidance).to.equal('verify-then-replay');

    // the grid did not grow
    expect(fake.worksheet(SPREADSHEET_ID, SHEET).rowCount).to.equal(3);
  });

  it('answers a 5xx-after-send as unknown and verify-then-replay for this idempotent overwrite', async () => {
    fake.failRequests('values:batchUpdate', 500, 1);
    const error = (await rejection(() =>
      gsheet.updateDataBatch([{ range: `${SHEET}!A1:B1`, values: [['v']] }])
    )) as GSheetError;

    expect(error.code).to.equal('UPSTREAM');
    expect(error.retryable).to.equal(true);
    const outcome = error.mutationOutcome!;
    expect(outcome.requests[0].state).to.equal('unknown');
    expect(outcome.requests[0].httpStatus).to.equal(500);
    expect(outcome.retryGuidance).to.equal('verify-then-replay');
    expect(outcome.retryGuidanceReason).to.match(/[Vv]erify/);
  });

  it('answers a lost response as unknown with no status and the original error as cause', async () => {
    const undo = dropMatchingRequests((url) => url.includes('/values:batchUpdate'));
    try {
      const error = (await rejection(() =>
        gsheet.updateDataBatch([{ range: `${SHEET}!A1:B1`, values: [['v']] }])
      )) as GSheetError;

      const outcome = error.mutationOutcome!;
      expect(outcome.requests[0].state).to.equal('unknown');
      expect(outcome.requests[0].httpStatus).to.equal(undefined);
      expect(outcome.retryGuidance).to.equal('verify-then-replay');
      expect((error as { cause?: unknown }).cause).to.be.an('error');
    } finally {
      undo();
    }
  });

  it('reports validate-phase outcomes for preflight refusals with unchanged messages', async () => {
    const non2d = (await rejection(() =>
      gsheet.updateDataBatch([
        // deliberately invalid input to drive the preflight refusal
        { range: `${SHEET}!A1:B2`, values: 'nope' as unknown as GoogleSheetCli.RawData },
      ])
    )) as GSheetError;
    expect(non2d).to.be.instanceOf(GSheetError);
    expect(non2d.message).to.equal(`Update values for range "${SHEET}!A1:B2" must be a 2D array`);
    expect(non2d.code).to.equal('VALIDATION');
    expect(non2d.mutationOutcome!.phase).to.equal('validate');
    expect(non2d.mutationOutcome!.requests).to.eql([]);
    expect(non2d.mutationOutcome!.retryGuidance).to.equal('safe-replay');

    const bounded = (await rejection(() =>
      gsheet.updateDataBatch([{ range: `${SHEET}!A1:B2`, values: [['1', '2'], ['3', '4'], ['5', '6']] }])
    )) as GSheetError;
    expect(bounded.message).to.include('exceeds explicit bounded range');
    expect(bounded.code).to.equal('VALIDATION');
    expect(bounded.mutationOutcome!.phase).to.equal('validate');
    expect(bounded.mutationOutcome!.requests).to.eql([]);

    // formula-overwrite refusal: a preflight rejection before anything was dispatched
    fake.setCells(SPREADSHEET_ID, SHEET, 'A1', [['=SUM(B1:B2)']]);
    const conflict = (await rejection(() =>
      gsheet.updateDataBatch([{ range: `${SHEET}!A1:B1`, values: [['x', 'y']] }])
    )) as GSheetError;
    expect(conflict.message).to.match(
      /Cannot overwrite existing formula\(s\) without overwriteFormulas=true\. Conflicts detected: .+"=SUM\(B1:B2\)"/
    );
    expect(conflict.code).to.equal('VALIDATION');
    expect(conflict.mutationOutcome!.phase).to.equal('validate');
    expect(conflict.mutationOutcome!.requests).to.eql([]);
    expect(conflict.mutationOutcome!.retryGuidance).to.equal('safe-replay');
    expect(fake.cell(SPREADSHEET_ID, SHEET, 'A1')).to.equal('=SUM(B1:B2)');
  });

  describe('buildBatchUpdateMutationOutcome (direct)', () => {
    const chunks = [
      [{ range: 'S!A1' }, { range: 'S!B1' }],
      [{ range: 'S!A2' }],
      [{ range: 'S!A3' }],
    ];

    it('maps the acknowledged prefix, the evidence-based failing chunk and the not-attempted tail', () => {
      const report = buildBatchUpdateMutationOutcome({
        batches: chunks,
        batchesExecuted: 1,
        cause: Object.assign(new Error('boom'), { response: { status: 429 } }),
      });
      expect(report.requests.map((request) => request.state)).to.eql(['acknowledged', 'rejected', 'not-attempted']);
      // one physical chunk keeps every logical range it covers
      expect(report.requests[0].a1Ranges).to.eql(['S!A1', 'S!B1']);
      expect(report.requests[0].kind).to.equal('values.batchUpdate');
      expect(report.requests[1].httpStatus).to.equal(429);
      expect(report.requests[1].causeSummary).to.match(/rejected before application/);
      expect(report.requests[2].causeSummary).to.match(/not dispatched/);
      // chunk one was acknowledged, so the derivation cannot call replay blind-safe: part of
      // the input is already applied and a replay must verify first
      expect(report.retryGuidance).to.equal('verify-then-replay');
    });

    it('locks never-blind-replay for a lost response on a non-idempotent operation', () => {
      const report = buildBatchUpdateMutationOutcome({
        batches: chunks,
        batchesExecuted: 1,
        cause: new Error('socket hang up'),
        nonIdempotent: true,
      });
      expect(report.requests[1].state).to.equal('unknown');
      expect(report.requests[1].httpStatus).to.equal(undefined);
      expect(report.retryGuidance).to.equal('never-blind-replay');
      expect(report.retryGuidanceReason).to.match(/duplicate/);
    });

    it('reports an unconfirmed growth conservatively and every chunk not-attempted in the grid-growth phase', () => {
      const report = buildBatchUpdateMutationOutcome({
        batches: chunks,
        batchesExecuted: 0,
        cause: Object.assign(new Error('backend error'), { status: 503 }),
        gridGrowth: { attempted: true, completed: false, details: '1 appendDimension request(s) not confirmed' },
        phase: 'grid-growth',
        failedChunkDispatched: false,
      });
      expect(report.phase).to.equal('grid-growth');
      expect(report.requests.map((request) => request.state)).to.eql(['not-attempted', 'not-attempted', 'not-attempted']);
      expect(report.requests[0].httpStatus).to.equal(undefined);
      expect(report.gridGrowth).to.deep.equal({
        attempted: true,
        completed: false,
        details: '1 appendDimension request(s) not confirmed',
      });
      // unconfirmed growth is more conservative than the request-only safe-replay derivation
      expect(report.retryGuidance).to.equal('verify-then-replay');
      expect(report.retryGuidanceReason).to.match(/grid/);
    });

    it('defaults the operation label and the values-write phase', () => {
      const report = buildBatchUpdateMutationOutcome({
        batches: [[{ range: 'S!A1' }]],
        batchesExecuted: 0,
        cause: Object.assign(new Error('boom'), { status: 500 }),
      });
      expect(report.operation).to.equal('updateDataBatch');
      expect(report.phase).to.equal('values-write');
      expect(report.requests[0].state).to.equal('unknown');
    });
  });
});
