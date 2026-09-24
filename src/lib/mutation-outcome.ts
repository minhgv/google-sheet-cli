/**
 * Typed mutation-outcome contract for multi-request write operations.
 *
 * A batch mutation (chunked value writes plus grid growth) can partially apply: some physical
 * requests land, some are rejected, and for some the transport dies before a usable response,
 * so the operation as a whole has no single success/failure bit. This module defines the locked
 * vocabulary later waves consume to report that truthfully.
 *
 * Locked vocabulary (do not rename; downstream consumers match on these strings):
 * - outcome states: `acknowledged` | `rejected` | `unknown` | `not-attempted`
 * - retry guidance: `safe-replay` | `verify-then-replay` | `never-blind-replay`
 *
 * Semantics:
 * - `acknowledged`  - a response proves the request was applied (or the API contract guarantees
 *                     application once accepted). Never relabeled afterwards.
 * - `rejected`      - a response proves the request was NOT applied (definitive rejection).
 * - `unknown`       - the request was dispatched but no usable response came back; it may or
 *                     may not have been applied. Conservative uncertainty is correct, not a bug.
 * - `not-attempted` - the request was never dispatched (skipped after an earlier failure).
 *
 * Outcomes are tracked per PHYSICAL request/chunk while `a1Ranges` retains the LOGICAL range(s)
 * from the caller's input, so a consumer can always answer "which of my asked-for cells were
 * written?" without re-deriving chunk boundaries. Grid-growth effects (row/column inserts) are
 * tracked separately from cell updates: a grid can grow while the value write fails.
 *
 * Retry guidance is conservative by construction: an `unknown` outcome never becomes an
 * unconditional replay instruction. `never-blind-replay` must be chosen explicitly for
 * non-idempotent operations (append / create / copy) - see `unknownOutcomeGuidance`.
 *
 * Pure module: types, a builder, a serializer, no I/O and no dependency on the rest of lib.
 */

/** Outcome state of one physical request/chunk. */
export type MutationOutcomeState = 'acknowledged' | 'rejected' | 'unknown' | 'not-attempted';

/** All outcome states, locked order for tests and docs. */
export const MUTATION_OUTCOME_STATES = ['acknowledged', 'rejected', 'unknown', 'not-attempted'] as const;

/** Conservative replay recommendation for the operation as a whole. */
export type RetryGuidanceClassification = 'safe-replay' | 'verify-then-replay' | 'never-blind-replay';

/** Operation phase a failure occurred in. Extending this union later is additive. */
export type MutationPhase = 'validate' | 'authorize' | 'grid-growth' | 'values-write' | 'finalize' | 'unknown';

/** All phases, locked order for tests and docs. */
export const MUTATION_PHASES = ['validate', 'authorize', 'grid-growth', 'values-write', 'finalize', 'unknown'] as const;

/** Outcome of one physical request/chunk. `a1Ranges` are the caller's logical ranges, verbatim. */
export interface MutationRequestOutcome {
  /** 0-based dispatch position of the physical request/chunk. */
  readonly requestIndex: number;
  /** What kind of physical request this was (e.g. 'values.update', 'values.append', 'grid.grow', 'worksheet.add'). */
  readonly kind: string;
  /** Logical A1 range(s) from the operation input this request covers; empty for grid-level requests. */
  readonly a1Ranges: readonly string[];
  /** Outcome state of this physical request. */
  readonly state: MutationOutcomeState;
  /** HTTP status when one was observed for this request. */
  readonly httpStatus?: number;
  /** Short summary of what happened. Prose, sanitized before serialization; dropped when redacted. */
  readonly causeSummary?: string;
}

/** Grid-growth effect, tracked separately from cell updates. */
export interface MutationGridGrowth {
  /** The operation attempted to grow the grid (add rows/columns). */
  readonly attempted: boolean;
  /** A response proved the growth was applied. */
  readonly completed: boolean;
  /** Short prose detail (e.g. "grew rows to 120"). Sanitized; dropped when redacted. */
  readonly details?: string;
}

/** Counts per outcome state - the "counts" the envelope contract keeps under redaction. */
export interface MutationRequestSummary {
  total: number;
  acknowledged: number;
  rejected: number;
  unknown: number;
  notAttempted: number;
}

/** Full report for one failed (or completed-with-caveats) mutation operation. */
export interface MutationOutcomeReport {
  /** Operation name, e.g. 'data:update' or 'updateDataBatch'. */
  readonly operation: string;
  /** Phase where the failure occurred. */
  readonly phase: MutationPhase;
  /** One entry per physical request/chunk, in dispatch order. */
  readonly requests: readonly MutationRequestOutcome[];
  /** Grid-growth effect, tracked separately from the value writes. */
  readonly gridGrowth: MutationGridGrowth;
  /** Conservative replay recommendation for the whole operation. */
  readonly retryGuidance: RetryGuidanceClassification;
  /** Why this guidance applies; what a caller must verify or avoid. */
  readonly retryGuidanceReason: string;
}

/** Envelope-safe projection of one request: structure only, prose included only when unredacted. */
export interface SerializedMutationRequest {
  requestIndex: number;
  kind: string;
  a1Ranges: string[];
  state: MutationOutcomeState;
  httpStatus?: number;
  /** Present only when redaction is off; sanitized. */
  causeSummary?: string;
}

/** Envelope-safe projection of the report. Coordinates, states and counts survive redaction; prose does not. */
export interface SerializedMutationOutcome {
  operation: string;
  phase: MutationPhase;
  /** Per-state request counts, so a consumer reads coverage without walking `requests`. */
  summary: MutationRequestSummary;
  requests: SerializedMutationRequest[];
  gridGrowth: {
    attempted: boolean;
    completed: boolean;
    /** Present only when redaction is off; sanitized. */
    details?: string;
  };
  retryGuidance: RetryGuidanceClassification;
  retryGuidanceReason: string;
}

/** Input for one `MutationOutcomeBuilder.request()` call; `a1Ranges` accepts a single range or a list. */
export interface MutationRequestInput {
  requestIndex: number;
  kind: string;
  a1Ranges?: string | readonly (string | undefined | null)[];
  state: MutationOutcomeState;
  httpStatus?: number;
  causeSummary?: string;
}

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);

/** Freeze a report deeply: the built contract is immutable, later waves build new ones. */
const deepFreeze = <T>(value: T): T => {
  if (Array.isArray(value)) {
    value.forEach(deepFreeze);
  } else if (isPlainObject(value)) {
    Object.values(value).forEach(deepFreeze);
  }
  return Object.isFrozen(value) ? value : Object.freeze(value);
};

const throwInvalid = (field: string, value: unknown): never => {
  throw new TypeError(`Invalid mutation outcome ${field}: ${String(value)}`);
};

const normalizeRanges = (a1Ranges: MutationRequestInput['a1Ranges']): string[] => {
  if (a1Ranges === undefined) return [];
  const list = typeof a1Ranges === 'string' ? [a1Ranges] : a1Ranges;
  return list
    .filter((range): range is string => typeof range === 'string')
    .map((range) => range.trim())
    .filter((range) => range.length > 0);
};

const isValidHttpStatus = (status: number): boolean =>
  // valid HTTP status codes are integers 100-599; same range asHttpStatus enforces in cli-errors
  Number.isInteger(status) && status >= 100 && status <= 599;

/**
 * Guidance for an `unknown` request outcome. Idempotent operations (reads, deterministic
 * overwrites) may be replayed after verification; non-idempotent ones (append / create /
 * copy) must never be blindly replayed because a hidden application would duplicate the write.
 * Wave B (updateDataBatch) classifies its operation once and calls this.
 */
export function unknownOutcomeGuidance(nonIdempotent: boolean): RetryGuidanceClassification {
  return nonIdempotent ? 'never-blind-replay' : 'verify-then-replay';
}

const deriveGuidance = (requests: readonly MutationRequestOutcome[]): { retryGuidance: RetryGuidanceClassification; retryGuidanceReason: string } => {
  const byState = (state: MutationOutcomeState) => requests.some((request) => request.state === state);
  if (requests.length === 0) {
    return {
      retryGuidance: 'safe-replay',
      retryGuidanceReason: 'No request was dispatched, so nothing can have been applied.',
    };
  }
  if (byState('unknown')) {
    return {
      retryGuidance: 'verify-then-replay',
      retryGuidanceReason:
        'At least one request was dispatched without a usable response; it may or may not have been applied. Verify remote state before any replay.',
    };
  }
  if (byState('acknowledged')) {
    return {
      retryGuidance: 'verify-then-replay',
      retryGuidanceReason:
        'Some requests were acknowledged before the failure; part of the input may already be applied. Verify which ranges are written before replaying.',
    };
  }
  return {
    retryGuidance: 'safe-replay',
    retryGuidanceReason:
      'Every dispatched request was rejected before application or never attempted, so replaying cannot duplicate a write.',
  };
};

/**
 * Assembles a `MutationOutcomeReport` as an operation progresses. The builder validates the
 * locked vocabulary eagerly (a bad state or phase throws at the call site, not in the
 * envelope) and derives conservative retry guidance when the caller does not set one
 * explicitly - non-idempotent operations should always set `never-blind-replay` explicitly
 * for `unknown` outcomes via `retryGuidance()` / `unknownOutcomeGuidance()`.
 */
export class MutationOutcomeBuilder {
  private readonly requests: MutationRequestOutcome[] = [];
  private grid: MutationGridGrowth = { attempted: false, completed: false };
  private ph: MutationPhase = 'unknown';
  private explicitGuidance?: { retryGuidance: RetryGuidanceClassification; retryGuidanceReason: string };

  constructor(public readonly operation: string) {
    if (typeof operation !== 'string' || operation.trim().length === 0) {
      throwInvalid('operation name', operation);
    }
  }

  /** Record the outcome of one physical request/chunk. */
  request(input: MutationRequestInput): this {
    if (!Number.isInteger(input.requestIndex) || input.requestIndex < 0) {
      throwInvalid('requestIndex', input.requestIndex);
    }
    if (typeof input.kind !== 'string' || input.kind.trim().length === 0) {
      throwInvalid('request kind', input.kind);
    }
    if (!MUTATION_OUTCOME_STATES.includes(input.state)) {
      throwInvalid('state', input.state);
    }
    const outcome: MutationRequestOutcome = {
      requestIndex: input.requestIndex,
      kind: input.kind,
      a1Ranges: normalizeRanges(input.a1Ranges),
      state: input.state,
      ...(input.httpStatus !== undefined
        ? isValidHttpStatus(input.httpStatus)
          ? { httpStatus: input.httpStatus }
          : throwInvalid('httpStatus', input.httpStatus)
        : {}),
      ...(input.causeSummary !== undefined ? { causeSummary: String(input.causeSummary) } : {}),
    };
    this.requests.push(outcome);
    return this;
  }

  /** Record the grid-growth effect, separately from cell updates. */
  gridGrowth(attempted: boolean, completed: boolean, details?: string): this {
    this.grid = {
      attempted: Boolean(attempted),
      completed: Boolean(completed),
      ...(details !== undefined ? { details: String(details) } : {}),
    };
    return this;
  }

  /** Set the phase the failure occurred in. */
  phase(phase: MutationPhase): this {
    if (!MUTATION_PHASES.includes(phase)) {
      throwInvalid('phase', phase);
    }
    this.ph = phase;
    return this;
  }

  /**
   * Override the derived guidance. Required for non-idempotent operations with `unknown`
   * outcomes (`never-blind-replay`), since the derivation cannot know operation semantics.
   */
  retryGuidance(classification: RetryGuidanceClassification, reason: string): this {
    if (
      !['safe-replay', 'verify-then-replay', 'never-blind-replay'].includes(classification)
    ) {
      throwInvalid('retryGuidance classification', classification);
    }
    if (typeof reason !== 'string' || reason.trim().length === 0) {
      throwInvalid('retryGuidance reason', reason);
    }
    this.explicitGuidance = { retryGuidance: classification, retryGuidanceReason: reason };
    return this;
  }

  /** Produce the frozen report; guidance is derived conservatively when not set explicitly. */
  build(): MutationOutcomeReport {
    return deepFreeze({
      operation: this.operation,
      phase: this.ph,
      requests: deepFreeze(this.requests.map((request) => deepFreeze({ ...request, a1Ranges: [...request.a1Ranges] }))),
      gridGrowth: deepFreeze({ ...this.grid }),
      ...(this.explicitGuidance ?? deriveGuidance(this.requests)),
    });
  }
}

export interface SerializeMutationOutcomeOptions {
  /** When true, prose (cause summaries, grid-growth details) is dropped; states/ranges/counts/status stay. */
  redacted?: boolean;
  /** Prose sanitizer, e.g. `sanitizeMessage` from cli-errors. Defaults to identity. */
  sanitize?: (message: string) => string;
}

const summarize = (requests: readonly MutationRequestOutcome[]): MutationRequestSummary => {
  const summary: MutationRequestSummary = { total: requests.length, acknowledged: 0, rejected: 0, unknown: 0, notAttempted: 0 };
  for (const request of requests) {
    if (request.state === 'not-attempted') summary.notAttempted += 1;
    else summary[request.state] += 1;
  }
  return summary;
};

/**
 * Envelope-safe projection of a report. Cell contents and formulas never appear here by
 * construction - requests carry states, coordinates and counts; the only prose fields
 * (`causeSummary`, `gridGrowth.details`) are sanitized and dropped entirely in redacted mode.
 * `retryGuidanceReason` survives redaction: it is generated prose about operation semantics,
 * never cell data, and without it the guidance classification is not actionable.
 */
export function serializeMutationOutcome(
  report: MutationOutcomeReport,
  options: SerializeMutationOutcomeOptions = {}
): SerializedMutationOutcome {
  const redacted = Boolean(options.redacted);
  const clean = options.sanitize ?? ((message: string): string => message);
  return {
    operation: clean(report.operation),
    phase: report.phase,
    summary: summarize(report.requests),
    requests: report.requests.map((request) => ({
      requestIndex: request.requestIndex,
      kind: clean(request.kind),
      a1Ranges: [...request.a1Ranges],
      state: request.state,
      ...(request.httpStatus !== undefined ? { httpStatus: request.httpStatus } : {}),
      ...(!redacted && request.causeSummary !== undefined ? { causeSummary: clean(request.causeSummary) } : {}),
    })),
    gridGrowth: {
      attempted: report.gridGrowth.attempted,
      completed: report.gridGrowth.completed,
      ...(!redacted && report.gridGrowth.details !== undefined ? { details: clean(report.gridGrowth.details) } : {}),
    },
    retryGuidance: report.retryGuidance,
    retryGuidanceReason: clean(report.retryGuidanceReason),
  };
}
