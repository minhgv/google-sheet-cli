import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { generateKeyPairSync } from 'crypto';

import {
  ErrorIssue,
  GSheetError,
  GSheetErrorCode,
  isRedactionEnabled,
  redactDiagnostic,
  setRedactionEnabled,
  toErrorEnvelope,
} from '../src/lib/cli-errors';
import {
  MUTATION_OUTCOME_STATES,
  MUTATION_PHASES,
  MutationOutcomeBuilder,
  MutationOutcomeReport,
  type MutationOutcomeState,
  type MutationPhase,
  type RetryGuidanceClassification,
  serializeMutationOutcome,
  unknownOutcomeGuidance,
} from '../src/lib/mutation-outcome';
import * as factory from '../src/lib/factory';
import GoogleSheet from '../src/lib/google-sheet';
import type { GoogleSheetCli } from '../src/lib/google-sheet';

/**
 * The mutation-outcome vocabulary and the redacted-diagnostics mode (plan
 * 2026-09-21-agent-gateway-hardening, T-01 + T-04).
 *
 * Pure cases pin the locked vocabulary, the builder's normalization/derivation rules and the
 * envelope projection. Command cases run through `runCommand` with the factory swapped for a
 * throwing stub, so the `--redacted` flag is proven end to end: parsed in `init()`, applied by
 * `toErrorEnvelope`, visible in the emitted envelope on stderr.
 */

const SPREADSHEET_ID = 'offline-spreadsheet-id';
const WORKSHEET_TITLE = 'offline-worksheet';
const CLIENT_EMAIL = 'offline@example.iam.gserviceaccount.com';

const PRIVATE_KEY = generateKeyPairSync('rsa', {
  modulusLength: 2048,
  privateKeyEncoding: { format: 'pem', type: 'pkcs8' },
  publicKeyEncoding: { format: 'pem', type: 'spki' },
}).privateKey;

const MANAGED_ENV = [
  'GSHEET_CLIENT_EMAIL',
  'GSHEET_PRIVATE_KEY',
  'GSHEET_CREDENTIALS_FILE',
  'GSHEET_REDACTED',
  'SPREADSHEET_ID',
  'WORKSHEET_TITLE',
];

const AUTH_FLAGS = [`--spreadsheetId=${SPREADSHEET_ID}`, `--worksheetTitle=${WORKSHEET_TITLE}`, `--clientEmail=${CLIENT_EMAIL}`];

const FAKE_PEM = '-----BEGIN PRIVATE KEY-----\nCANARY-KEY-BODY\n-----END PRIVATE KEY-----';

interface EnvelopeIssue {
  row?: number;
  column?: number;
  a1?: string;
  field?: string;
  code?: string;
  message: string;
  value?: string | number | boolean | null;
}

interface EnvelopeRequest {
  requestIndex: number;
  kind: string;
  a1Ranges: string[];
  state: string;
  httpStatus?: number;
  causeSummary?: string;
}

interface Envelope {
  error: {
    code: string;
    message: string;
    retryable: boolean;
    retryAfterMs?: number;
    issues?: EnvelopeIssue[];
    mutation?: {
      operation: string;
      phase: string;
      summary: { total: number; acknowledged: number; rejected: number; unknown: number; notAttempted: number };
      requests: EnvelopeRequest[];
      gridGrowth: { attempted: boolean; completed: boolean; details?: string };
      retryGuidance: string;
      retryGuidanceReason: string;
    };
  };
}

const parseEnvelope = (stderr: string): Envelope => JSON.parse(stderr) as Envelope;

/** One failing chunk; the PEM body must never survive into any envelope or diagnostic. */
const sampleReport = (): MutationOutcomeReport =>
  new MutationOutcomeBuilder('data:update')
    .phase('values-write')
    .gridGrowth(true, true, 'grew rows to 120')
    .request({ requestIndex: 0, kind: 'grid.grow', state: 'acknowledged', httpStatus: 200, causeSummary: 'grid grew first' })
    .request({ requestIndex: 1, kind: 'values.update', a1Ranges: 'Sheet1!A1:B2', state: 'acknowledged', httpStatus: 200 })
    .request({
      requestIndex: 2,
      kind: 'values.update',
      a1Ranges: ['Sheet1!C1:C4'],
      state: 'unknown',
      httpStatus: 500,
      causeSummary: `connection died after dispatch ${FAKE_PEM}`,
    })
    .request({ requestIndex: 3, kind: 'values.update', a1Ranges: ['Sheet1!D1:D4'], state: 'not-attempted' })
    .build();

const sampleIssue = (): ErrorIssue => ({
  row: 2,
  column: 3,
  a1: 'Sheet1!C2',
  code: 'CELL_INVALID',
  message: 'cell failed validation',
  value: 'CANARY-VALUE',
});

const sampleError = (): GSheetError =>
  new GSheetError(GSheetErrorCode.UPSTREAM, 'update failed partway through', {
    mutationOutcome: sampleReport(),
    details: { issues: [sampleIssue()] },
    cause: { private_key: 'CANARY-KEY-BODY', rows: [['CANARY-CELL']] },
  });

/** Stub that authorizes fine and fails the first metadata call with the injected error. */
class ThrowingStubGoogleSheet {
  injected: unknown = sampleError();

  async authorize(): Promise<void> {}

  async authorizeOAuth(): Promise<void> {
    throw new Error('authorizeOAuth is not expected in these cases');
  }

  async getWorksheetMetadata(): Promise<GoogleSheetCli.WorksheetMetadata> {
    throw this.injected;
  }
}

const mutableFactory = factory as { createGoogleSheet: () => GoogleSheet };
const realCreateGoogleSheet = factory.createGoogleSheet;

describe('mutation outcome contract and redacted diagnostics', () => {
  let savedEnv: Record<string, string | undefined>;

  beforeEach(() => {
    savedEnv = Object.fromEntries(MANAGED_ENV.map((name) => [name, process.env[name]]));
    for (const name of MANAGED_ENV) delete process.env[name];
    process.env.GSHEET_PRIVATE_KEY = PRIVATE_KEY;
    process.env.NODE_ENV = 'test';
    mutableFactory.createGoogleSheet = () => new ThrowingStubGoogleSheet() as unknown as GoogleSheet;
  });

  afterEach(() => {
    mutableFactory.createGoogleSheet = realCreateGoogleSheet;
    for (const [name, value] of Object.entries(savedEnv)) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
    setRedactionEnabled(false);
  });

  describe('mutation outcome vocabulary', () => {
    it('locks the outcome states and phases', () => {
      expect([...MUTATION_OUTCOME_STATES]).to.deep.equal(['acknowledged', 'rejected', 'unknown', 'not-attempted']);
      expect([...MUTATION_PHASES]).to.deep.equal(['validate', 'authorize', 'grid-growth', 'values-write', 'finalize', 'unknown']);
    });

    it('normalizes a single A1 range and a list, trimming and dropping blanks', () => {
      const fromString = new MutationOutcomeBuilder('data:update').request({ requestIndex: 0, kind: 'values.update', a1Ranges: 'Sheet1!A1:B2', state: 'acknowledged' }).build();
      expect([...fromString.requests[0].a1Ranges]).to.deep.equal(['Sheet1!A1:B2']);
      const fromList = new MutationOutcomeBuilder('data:update')
        .request({ requestIndex: 0, kind: 'values.update', a1Ranges: ['Sheet1!A1', '  ', '', null, undefined, 'Sheet1!B2'], state: 'rejected' })
        .build();
      expect([...fromList.requests[0].a1Ranges]).to.deep.equal(['Sheet1!A1', 'Sheet1!B2']);
    });

    it('freezes the built report so the vocabulary cannot drift after the fact', () => {
      const report = sampleReport();
      expect(Object.isFrozen(report)).to.be.true;
      expect(Object.isFrozen(report.requests)).to.be.true;
      expect(Object.isFrozen(report.requests[0])).to.be.true;
      expect(Object.isFrozen(report.requests[1].a1Ranges)).to.be.true;
      expect(Object.isFrozen(report.gridGrowth)).to.be.true;
    });

    it('rejects values outside the locked vocabulary at the call site', () => {
      expect(() => new MutationOutcomeBuilder('  ').build()).to.throw(TypeError);
      expect(() => new MutationOutcomeBuilder('op').request({ requestIndex: -1, kind: 'values.update', state: 'acknowledged' })).to.throw(TypeError);
      expect(() => new MutationOutcomeBuilder('op').request({ requestIndex: 0, kind: '  ', state: 'acknowledged' })).to.throw(TypeError);
      expect(() => new MutationOutcomeBuilder('op').request({ requestIndex: 0, kind: 'values.update', state: 'applied' as MutationOutcomeState })).to.throw(TypeError);
      expect(() => new MutationOutcomeBuilder('op').request({ requestIndex: 0, kind: 'values.update', state: 'acknowledged', httpStatus: 42 })).to.throw(TypeError);
      expect(() => new MutationOutcomeBuilder('op').phase('teardown' as MutationPhase)).to.throw(TypeError);
      expect(() => new MutationOutcomeBuilder('op').retryGuidance('nuke-it' as RetryGuidanceClassification, 'because')).to.throw(TypeError);
    });

    it('derives conservative guidance: nothing dispatched or nothing applied replays safely', () => {
      const empty = new MutationOutcomeBuilder('data:update').build();
      expect(empty.retryGuidance).to.equal('safe-replay');
      const allRejected = new MutationOutcomeBuilder('data:update')
        .request({ requestIndex: 0, kind: 'values.update', state: 'rejected' })
        .request({ requestIndex: 1, kind: 'values.update', state: 'not-attempted' })
        .build();
      expect(allRejected.retryGuidance).to.equal('safe-replay');
      expect(allRejected.retryGuidanceReason).to.contain('cannot duplicate');
    });

    it('derives verify-then-replay when any request is unknown or acknowledged', () => {
      const withUnknown = new MutationOutcomeBuilder('data:append-table')
        .request({ requestIndex: 0, kind: 'values.append', state: 'acknowledged' })
        .request({ requestIndex: 1, kind: 'values.append', state: 'unknown' })
        .build();
      expect(withUnknown.retryGuidance).to.equal('verify-then-replay');
      const withAcknowledged = new MutationOutcomeBuilder('data:update')
        .request({ requestIndex: 0, kind: 'values.update', state: 'acknowledged' })
        .request({ requestIndex: 1, kind: 'values.update', state: 'rejected' })
        .build();
      expect(withAcknowledged.retryGuidance).to.equal('verify-then-replay');
      expect(withAcknowledged.retryGuidanceReason).to.contain('Verify');
    });

    it('lets the caller override guidance - non-idempotent unknown outcomes must say never-blind-replay', () => {
      expect(unknownOutcomeGuidance(false)).to.equal('verify-then-replay');
      expect(unknownOutcomeGuidance(true)).to.equal('never-blind-replay');
      const appended = new MutationOutcomeBuilder('data:append-table')
        .request({ requestIndex: 0, kind: 'values.append', a1Ranges: 'Sheet1!A11:A12', state: 'unknown' })
        .retryGuidance('never-blind-replay', 'an append that actually landed would be duplicated by a blind replay')
        .build();
      expect(appended.retryGuidance).to.equal('never-blind-replay');
      expect(appended.retryGuidanceReason).to.contain('duplicated');
    });

    it('records grid growth and phase separately from the cell requests', () => {
      const report = sampleReport();
      expect(report.phase).to.equal('values-write');
      expect(report.gridGrowth).to.deep.equal({ attempted: true, completed: true, details: 'grew rows to 120' });
      expect(report.requests[0].kind).to.equal('grid.grow');
      expect(report.requests[0].a1Ranges).to.deep.equal([]);
    });

    it('projects the report for the envelope: prose sanitized unredacted, dropped when redacted', () => {
      const clean = (message: string): string => message.replace(/CANARY/g, 'SCRUBBED');
      const unredacted = serializeMutationOutcome(sampleReport(), { sanitize: clean });
      expect(unredacted.summary).to.deep.equal({ total: 4, acknowledged: 2, rejected: 0, unknown: 1, notAttempted: 1 });
      expect(unredacted.requests[2].causeSummary).to.contain('SCRUBBED');
      expect(unredacted.requests[2].causeSummary).to.not.contain('CANARY');
      expect(unredacted.requests[2].httpStatus).to.equal(500);
      expect(unredacted.gridGrowth.details).to.equal('grew rows to 120');
      expect(unredacted.retryGuidanceReason).to.be.a('string').with.length.above(0);

      const redacted = serializeMutationOutcome(sampleReport(), { redacted: true, sanitize: clean });
      for (const request of redacted.requests) expect(request).to.not.have.property('causeSummary');
      expect(redacted.gridGrowth).to.not.have.property('details');
      expect(redacted.requests[2]).to.deep.equal({ requestIndex: 2, kind: 'values.update', a1Ranges: ['Sheet1!C1:C4'], state: 'unknown', httpStatus: 500 });
      expect(redacted.retryGuidance).to.equal('verify-then-replay');
      expect(redacted.retryGuidanceReason).to.have.length.above(0);
      expect(redacted.summary).to.deep.equal(unredacted.summary);
    });
  });

  describe('GSheetError envelope with a mutation outcome', () => {
    it('carries the sanitized mutation block and keeps legacy fields intact', () => {
      const envelope = toErrorEnvelope(sampleError());
      expect(envelope.error.code).to.equal('UPSTREAM');
      expect(envelope.error.retryable).to.equal(true);
      expect(envelope.error.mutation).to.exist;
      expect(envelope.error.mutation?.operation).to.equal('data:update');
      expect(envelope.error.mutation?.phase).to.equal('values-write');
      expect(envelope.error.mutation?.summary).to.deep.equal({ total: 4, acknowledged: 2, rejected: 0, unknown: 1, notAttempted: 1 });
      expect(envelope.error.mutation?.retryGuidance).to.equal('verify-then-replay');
      expect(envelope.error.mutation?.requests[2].causeSummary).to.contain('[REDACTED_KEY]');
      expect(envelope.error.mutation?.requests[2].state).to.equal('unknown');
    });

    it('never leaks the cause or cell contents into the envelope', () => {
      // the issue's own bounded `value` is allowed through unredacted (pinned above); what must
      // never pass is the cause tree - its key material and its nested data rows
      const serialized = JSON.stringify(toErrorEnvelope(sampleError()));
      expect(serialized).to.not.contain('CANARY-KEY-BODY');
      expect(serialized).to.not.contain('CANARY-CELL');
      expect(serialized).to.not.contain('BEGIN PRIVATE KEY');
      expect(serialized).to.not.contain('"private_key"');
    });

    it('serializes legacy errors without the new fields exactly as before', () => {
      const legacy = toErrorEnvelope(new GSheetError(GSheetErrorCode.NOT_FOUND, 'Worksheet "X" not found'));
      expect(legacy).to.deep.equal({ error: { code: 'NOT_FOUND', message: 'Worksheet "X" not found', retryable: false } });
      expect(legacy.error).to.not.have.property('mutation');

      const throttled = toErrorEnvelope(new GSheetError(GSheetErrorCode.RATE_LIMITED, 'slow down', { retryAfterMs: 500 }));
      expect(throttled).to.deep.equal({ error: { code: 'RATE_LIMITED', message: 'slow down', retryable: true, retryAfterMs: 500 } });
      expect(throttled.error).to.not.have.property('mutation');
    });

    it('serializes non-GSheetError values exactly as before', () => {
      expect(toErrorEnvelope(new Error('boom'))).to.deep.equal({ error: { code: 'INTERNAL', message: 'boom', retryable: false } });
      expect(toErrorEnvelope('nope')).to.deep.equal({ error: { code: 'INTERNAL', message: 'Unexpected error', retryable: false } });
    });

    it('keeps issue values when redaction is off - current behavior preserved', () => {
      const envelope = toErrorEnvelope(new GSheetError(GSheetErrorCode.DATA_INVALID, 'validation failed', { details: { issues: [sampleIssue()] } }));
      expect(envelope.error.issues).to.deep.equal([
        { row: 2, column: 3, a1: 'Sheet1!C2', code: 'CELL_INVALID', message: 'cell failed validation', value: 'CANARY-VALUE' },
      ]);
    });
  });

  describe('redaction mode', () => {
    it('starts off', () => {
      expect(isRedactionEnabled()).to.be.false;
    });

    it('drops issue values but keeps coordinates, codes and messages when on', () => {
      setRedactionEnabled(true);
      const envelope = toErrorEnvelope(new GSheetError(GSheetErrorCode.DATA_INVALID, 'validation failed', { details: { issues: [sampleIssue()] } }));
      expect(envelope.error.issues).to.deep.equal([{ row: 2, column: 3, a1: 'Sheet1!C2', code: 'CELL_INVALID', message: 'cell failed validation' }]);
    });

    it('strips mutation prose but keeps states, ranges, counts and status when on', () => {
      setRedactionEnabled(true);
      const mutation = toErrorEnvelope(sampleError()).error.mutation;
      expect(mutation).to.exist;
      for (const request of mutation?.requests ?? []) expect(request).to.not.have.property('causeSummary');
      expect(mutation?.gridGrowth).to.not.have.property('details');
      expect(mutation?.requests[2]).to.deep.equal({ requestIndex: 2, kind: 'values.update', a1Ranges: ['Sheet1!C1:C4'], state: 'unknown', httpStatus: 500 });
      expect(mutation?.summary.total).to.equal(4);
      expect(mutation?.retryGuidanceReason).to.have.length.above(0);
    });

    it('returns to byte-identical legacy output when toggled back off', () => {
      setRedactionEnabled(true);
      setRedactionEnabled(false);
      expect(toErrorEnvelope(new GSheetError(GSheetErrorCode.DATA_INVALID, 'validation failed', { details: { issues: [sampleIssue()] } }))).to.deep.equal({
        error: { code: 'DATA_INVALID', message: 'validation failed', retryable: false, issues: [{ row: 2, column: 3, a1: 'Sheet1!C2', code: 'CELL_INVALID', message: 'cell failed validation', value: 'CANARY-VALUE' }] },
      });
    });
  });

  describe('redactDiagnostic', () => {
    it('replaces scalar content and credentials with a placeholder, keeping structure', () => {
      const redacted = redactDiagnostic({
        a1: 'Sheet1!A1',
        count: 3,
        value: 'CANARY-VALUE',
        formula: '=SUM(A1:A9)',
        private_key: 'CANARY-KEY-BODY',
        'client-secret': 'CANARY-SECRET',
      });
      expect(redacted).to.deep.equal({
        a1: 'Sheet1!A1',
        count: 3,
        value: '[REDACTED]',
        formula: '[REDACTED]',
        private_key: '[REDACTED]',
        'client-secret': '[REDACTED]',
      });
    });

    it('keeps array counts while dropping array content, and recurses into nested objects', () => {
      const redacted = redactDiagnostic({
        values: [
          ['a', 'b'],
          ['c', 'd'],
        ],
        data: { formula: '=B6*2', rows: ['x', 'y', 'z'], a1: 'Sheet1!B6', note: 'kept prose' },
      });
      expect(redacted).to.deep.equal({
        values: [null, null],
        data: { formula: '[REDACTED]', rows: [null, null, null], a1: 'Sheet1!B6', note: 'kept prose' },
      });
    });

    it('scrubs secret patterns out of surviving strings', () => {
      const redacted = redactDiagnostic({ message: `dispatch failed with ${FAKE_PEM} and Bearer ya29.supersecret` });
      expect(JSON.stringify(redacted)).to.not.contain('CANARY-KEY-BODY');
      expect(JSON.stringify(redacted)).to.not.contain('ya29.supersecret');
      expect(JSON.stringify(redacted)).to.contain('[REDACTED_KEY]');
    });

    it('never mutates the input and cuts cyclic references', () => {
      const input: Record<string, unknown> = { value: 'CANARY-VALUE', nested: { formula: '=X' } };
      const snapshot = JSON.stringify(input);
      redactDiagnostic(input);
      expect(JSON.stringify(input)).to.equal(snapshot);

      const cyclic: Record<string, unknown> = { a1: 'Sheet1!A1', values: ['secret'] };
      cyclic.self = cyclic;
      expect(() => redactDiagnostic(cyclic)).to.not.throw();
      expect(redactDiagnostic(cyclic)).to.deep.equal({ a1: 'Sheet1!A1', values: [null], self: null });
    });
  });

  describe('--redacted flag wiring', () => {
    it('declares the shared flag with the GSHEET_REDACTED env alias', async () => {
      const { error, stdout } = await runCommand(['data:schema', '--help']);
      if (error) throw error;
      expect(stdout).to.contain('--redacted');
      expect(stdout).to.contain('GSHEET_REDACTED');
    });

    it('emits an envelope with issue values and mutation prose without the flag', async () => {
      const { error, stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      expect(error).to.exist;
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('UPSTREAM');
      expect(envelope.error.issues?.[0].value).to.equal('CANARY-VALUE');
      expect(envelope.error.mutation?.requests[2].causeSummary).to.contain('[REDACTED_KEY]');
      expect(envelope.error.mutation?.gridGrowth.details).to.equal('grew rows to 120');
    });

    it('emits a redacted envelope when --redacted is passed', async () => {
      const { error, stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json', '--redacted']);
      expect(error).to.exist;
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('UPSTREAM');
      expect(envelope.error.issues?.[0]).to.not.have.property('value');
      expect(envelope.error.issues?.[0].a1).to.equal('Sheet1!C2');
      for (const request of envelope.error.mutation?.requests ?? []) expect(request).to.not.have.property('causeSummary');
      expect(envelope.error.mutation?.gridGrowth).to.not.have.property('details');
      expect(envelope.error.mutation?.requests[2].state).to.equal('unknown');
      expect(JSON.stringify(envelope)).to.not.contain('CANARY');
    });

    it('honors the GSHEET_REDACTED environment variable without the flag', async () => {
      process.env.GSHEET_REDACTED = 'true';
      const { error, stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      expect(error).to.exist;
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.issues?.[0]).to.not.have.property('value');
      expect(envelope.error.mutation?.gridGrowth).to.not.have.property('details');
    });

    it('redacts dry-run style diagnostics through redactDiagnostic', () => {
      const diagnostic = { operation: 'data:update', dryRun: true, a1Ranges: ['Sheet1!A1:B2'], incoming: [['CANARY-CELL']] };
      expect(redactDiagnostic(diagnostic)).to.deep.equal({ operation: 'data:update', dryRun: true, a1Ranges: ['Sheet1!A1:B2'], incoming: [null] });
    });
  });

  it('leaves redaction off after the suite for later test files', () => {
    expect(isRedactionEnabled()).to.be.false;
  });
});
