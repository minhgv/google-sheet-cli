/**
 * The shared CLI error contract.
 *
 * Commands and library methods throw either plain `Error`s (existing behaviour, unchanged) or
 * `GSheetError` when a failure has a stable code and retry information a caller could act on.
 * The shared base-class `catch` is the only place anything is serialized: in JSON failure mode
 * (`--rawOutput` or `--json`) it prints exactly one envelope on stderr and exits 1; ordinary
 * mode keeps oclif's human error output. Nothing in this module ever copies an arbitrary error
 * object into the envelope - credentials, request/response internals and causes stay behind.
 */

import { Errors } from '@oclif/core';
import { MutationOutcomeReport, SerializedMutationOutcome, serializeMutationOutcome } from './mutation-outcome';

export const GSheetErrorCode = {
  /** argv could not be parsed (unknown flag, missing required flag, bad flag value) */
  USAGE: 'USAGE',
  /** local authentication could not be established (missing or unreadable credentials) */
  AUTH_REQUIRED: 'AUTH_REQUIRED',
  /** the remote API rejected the authorization (HTTP 401) */
  UNAUTHORIZED: 'UNAUTHORIZED',
  /** the remote API refused the operation (HTTP 403) */
  FORBIDDEN: 'FORBIDDEN',
  /** the requested resource does not exist (HTTP 404) */
  NOT_FOUND: 'NOT_FOUND',
  /** the operation conflicts with remote state (HTTP 409/412) */
  CONFLICT: 'CONFLICT',
  /** quota or rate limiting (HTTP 429); retryable */
  RATE_LIMITED: 'RATE_LIMITED',
  /** a Google-side server error (HTTP 5xx); retryable */
  UPSTREAM: 'UPSTREAM',
  /** the request never reached a usable HTTP status (DNS, connect, timeout) */
  NETWORK: 'NETWORK',
  /** a well-formed request was rejected (other 4xx) */
  REQUEST_INVALID: 'REQUEST_INVALID',
  /** command input parsed but failed a local check */
  VALIDATION: 'VALIDATION',
  /** a validation schema itself is malformed */
  SCHEMA_INVALID: 'SCHEMA_INVALID',
  /** data violates a valid schema or layout contract */
  DATA_INVALID: 'DATA_INVALID',
  /** anything without a better classification */
  INTERNAL: 'INTERNAL',
} as const;

export type GSheetErrorCodeValue = (typeof GSheetErrorCode)[keyof typeof GSheetErrorCode];

const CODES: Record<string, true> = Object.fromEntries(
  Object.values(GSheetErrorCode).map((code) => [code, true as const])
);

const CODES_RETRYABLE: Partial<Record<GSheetErrorCodeValue, true>> = {
  [GSheetErrorCode.RATE_LIMITED]: true,
  [GSheetErrorCode.UPSTREAM]: true,
  [GSheetErrorCode.NETWORK]: true,
};

/** One structured, coordinate-linked problem; the only detail shape the envelope carries. */
export interface ErrorIssue {
  /** 1-based row in the source the issue came from */
  row?: number;
  /** 1-based column in the source the issue came from */
  column?: number;
  /** A1 coordinate in the worksheet, when the source is a sheet range */
  a1?: string;
  /** field / header the issue belongs to */
  field?: string;
  /** stable issue code (e.g. report validation codes, DUPLICATE_HEADER) */
  code?: string;
  /** human-readable description */
  message: string;
  /** the offending value, only when it is a bounded primitive */
  value?: string | number | boolean | null;
}

export interface GSheetErrorOptions {
  /** defaults to the code's inherent retryability */
  retryable?: boolean;
  /** server-provided retry hint in milliseconds */
  retryAfterMs?: number;
  /** structured, coordinate-linked details (e.g. data validation issues) */
  details?: { issues?: ErrorIssue[] };
  /**
   * Typed mutation-outcome report for multi-request write operations (see
   * `./mutation-outcome`). Serialized into the envelope as a sanitized `error.mutation`
   * block - states, ranges, counts, phase and retry guidance only, never cell contents.
   */
  mutationOutcome?: MutationOutcomeReport;
  /**
   * Message shown INSTEAD of `message` when redaction is enabled. Use when the full
   * message embeds cell contents or formulas (e.g. formula-conflict refusals); keep
   * coordinates and counts, drop the values. Ignored when redaction is off.
   */
  redactedMessage?: string;
  /** preserved on the Error for humans/debuggers; never serialized into the envelope */
  cause?: unknown;
}

/**
 * An `Error` with safe, stable code/retry metadata the shared error boundary understands.
 * Everything set on the instance is safe to serialize; anything unsafe belongs on `cause`.
 */
export class GSheetError extends Error {
  public readonly code: GSheetErrorCodeValue;
  public readonly retryable: boolean;
  public readonly retryAfterMs?: number;
  public readonly details?: { issues?: ErrorIssue[] };
  public readonly mutationOutcome?: MutationOutcomeReport;
  public readonly redactedMessage?: string;

  constructor(code: GSheetErrorCodeValue, message: string, options: GSheetErrorOptions = {}) {
    // target is es2017, so the two-argument Error constructor (and its typed `cause`) is not
    // available; the property is still what Node's runtime Error#cause spells, just defined
    // by hand and non-enumerable like the native one, so spreads never pick it up.
    super(message);
    this.name = 'GSheetError';
    this.code = code;
    this.retryable = options.retryable ?? Boolean(CODES_RETRYABLE[code]);
    if (options.cause !== undefined) {
      Object.defineProperty(this, 'cause', { value: options.cause, configurable: true, writable: true });
    }
    if (options.retryAfterMs !== undefined) this.retryAfterMs = options.retryAfterMs;
    if (options.details !== undefined) this.details = options.details;
    if (options.mutationOutcome !== undefined) this.mutationOutcome = options.mutationOutcome;
    if (options.redactedMessage !== undefined) this.redactedMessage = options.redactedMessage;
  }
}

export interface CliErrorMeta {
  code: GSheetErrorCodeValue;
  retryable: boolean;
  retryAfterMs?: number;
}

const NETWORK_CODES: Record<string, true> = {
  ENOTFOUND: true,
  EAI_AGAIN: true,
  ECONNREFUSED: true,
  ETIMEDOUT: true,
  ECONNRESET: true,
  EHOSTUNREACH: true,
  ENETUNREACH: true,
  EPIPE: true,
  UND_ERR_CONNECT_TIMEOUT: true,
  UND_ERR_SOCKET: true,
};

const NETWORK_MESSAGE = /\b(fetch failed|network error|socket hang up|getaddrinfo|ECONNRESET|ETIMEDOUT|ENOTFOUND|ECONNREFUSED|EAI_AGAIN)\b/i;

const asNumber = (value: unknown): number | undefined =>
  typeof value === 'number' && Number.isFinite(value) ? value : undefined;

const asHttpStatus = (value: unknown): number | undefined => {
  const n = typeof value === 'number' ? value : typeof value === 'string' && /^\d+$/.test(value) ? Number(value) : NaN;
  return Number.isInteger(n) && n >= 100 && n <= 599 ? n : undefined;
};

const retryAfterMsFrom = (err: Record<string, any>): number | undefined => {
  const direct = asNumber(err.retryAfterMs);
  if (direct !== undefined && direct > 0) return Math.min(direct, 24 * 60 * 60 * 1000);
  const headers = err.response?.headers;
  const hint = headers?.['retry-after'] ?? headers?.['Retry-After'];
  const seconds = typeof hint === 'string' && /^\d+$/.test(hint) ? Number(hint) : asNumber(hint);
  return seconds !== undefined && seconds > 0 ? Math.min(seconds * 1000, 24 * 60 * 60 * 1000) : undefined;
};

/**
 * Map any thrown value to stable code/retry metadata. `GSheetError`s carry their own; oclif
 * parse failures are USAGE; Google API errors are classified by HTTP status only; everything
 * else falls back to NETWORK (recognizable transport failures) or INTERNAL.
 */
export function classifyError(err: unknown): CliErrorMeta {
  if (err instanceof GSheetError) {
    return err.retryAfterMs !== undefined
      ? { code: err.code, retryable: err.retryable, retryAfterMs: err.retryAfterMs }
      : { code: err.code, retryable: err.retryable };
  }
  if (err instanceof Errors.ExitError) return { code: GSheetErrorCode.INTERNAL, retryable: false };
  if (err instanceof Errors.CLIError) return { code: GSheetErrorCode.USAGE, retryable: false };
  if (err && typeof err === 'object') {
    const e = err as Record<string, any>;
    if (typeof e.code === 'string' && CODES[e.code]) {
      const code = e.code as GSheetErrorCodeValue;
      const retryable = typeof e.retryable === 'boolean' ? e.retryable : Boolean(CODES_RETRYABLE[code]);
      const retryAfterMs = asNumber(e.retryAfterMs);
      return retryAfterMs !== undefined ? { code, retryable, retryAfterMs } : { code, retryable };
    }
    const status = asHttpStatus(e.response?.status) ?? asHttpStatus(e.status) ?? asHttpStatus(e.code);
    if (status !== undefined) {
      if (status === 401) return { code: GSheetErrorCode.UNAUTHORIZED, retryable: false };
      if (status === 403) return { code: GSheetErrorCode.FORBIDDEN, retryable: false };
      if (status === 404) return { code: GSheetErrorCode.NOT_FOUND, retryable: false };
      if (status === 409 || status === 412) return { code: GSheetErrorCode.CONFLICT, retryable: false };
      if (status === 429) {
        const retryAfterMs = retryAfterMsFrom(e);
        return retryAfterMs !== undefined
          ? { code: GSheetErrorCode.RATE_LIMITED, retryable: true, retryAfterMs }
          : { code: GSheetErrorCode.RATE_LIMITED, retryable: true };
      }
      if (status >= 500) return { code: GSheetErrorCode.UPSTREAM, retryable: true };
      return { code: GSheetErrorCode.REQUEST_INVALID, retryable: false };
    }
    if (NETWORK_CODES[e.code] || NETWORK_MESSAGE.test(typeof e.message === 'string' ? e.message : '')) {
      return { code: GSheetErrorCode.NETWORK, retryable: true };
    }
  }
  return { code: GSheetErrorCode.INTERNAL, retryable: false };
}

/**
 * Whether a thrown value still carries transport evidence: an HTTP status in any of the usual
 * shapes, or a recognizable network failure. `resolveGoogleSheetAuth` uses this to wrap only
 * genuinely local authentication problems as AUTH_REQUIRED - anything that still carries a
 * status or a network signature is rethrown untouched, so the envelope keeps its upstream
 * classification (FORBIDDEN, RATE_LIMITED, NETWORK, ...) instead of hiding it behind auth.
 */
export function hasTransportShape(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  const e = err as Record<string, unknown>;
  const response = e.response;
  const responseStatus = response && typeof response === 'object' ? (response as Record<string, unknown>).status : undefined;
  if (asHttpStatus(responseStatus) !== undefined) return true;
  if (asHttpStatus(e.status) !== undefined) return true;
  if (asHttpStatus(e.code) !== undefined) return true;
  if (typeof e.code === 'string' && NETWORK_CODES[e.code]) return true;
  return NETWORK_MESSAGE.test(typeof e.message === 'string' ? e.message : '');
}

const MAX_MESSAGE_LENGTH = 400;
const MAX_VALUE_LENGTH = 200;

const SECRET_PATTERNS: [RegExp, string][] = [
  [/-----BEGIN [A-Z ]*PRIVATE KEY-----[\s\S]*?(-----END [A-Z ]*PRIVATE KEY-----|$)/g, '[REDACTED_KEY]'],
  [/eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}/g, '[REDACTED_JWT]'],
  [/ya29\.[A-Za-z0-9._-]+/g, 'ya29.[REDACTED]'],
  [/\bBearer\s+[A-Za-z0-9._~+/=-]{8,}/gi, 'Bearer [REDACTED]'],
  [/"(access_token|refresh_token|private_key|client_secret|id_token)"\s*:\s*"[^"]*"/gi, '"$1": "[REDACTED]"'],
];

/**
 * Make an error message safe for a machine-readable envelope: single line, secrets redacted,
 * hard length cap. Error text is kept otherwise - it is the part a caller can act on.
 */
export function sanitizeMessage(message: string): string {
  let safe = String(message);
  for (const [pattern, replacement] of SECRET_PATTERNS) {
    safe = safe.replace(pattern, replacement);
  }
  safe = safe.replace(/\s+/g, ' ').trim();
  if (safe.length > MAX_MESSAGE_LENGTH) {
    safe = `${safe.slice(0, MAX_MESSAGE_LENGTH)}...`;
  }
  return safe;
}

/**
 * Opt-in redacted-diagnostics mode. When enabled, every envelope and diagnostic this module
 * emits keeps coordinates, counts, statuses and outcome states but drops cell contents,
 * formulas, incoming values and `issues[].value` - applied before serialization, so nothing
 * sensitive reaches stdout/stderr or a caller. It never strips anything extra from error
 * messages themselves (those are already secret-scrubbed by `sanitizeMessage` and are the
 * actionable part); it removes the structured value fields instead.
 *
 * Off by default: existing output shapes are byte-identical until a command turns it on via
 * the shared `--redacted` flag (GSHEET_REDACTED). Unattended deployments should enable it.
 */
let redactionEnabled = false;

/** Turn redacted-diagnostics mode on/off. Idempotent; commands call this in `init()`. */
export function setRedactionEnabled(enabled: boolean): void {
  redactionEnabled = enabled;
}

/** Whether redacted-diagnostics mode is currently on. */
export function isRedactionEnabled(): boolean {
  return redactionEnabled;
}

/**
 * Keys whose values are data content or credentials, not structure. Matched case-insensitively
 * on the exact key name after stripping `-`/`_` separators: only the bare names below and
 * their snake/kebab spellings match - compound camelCase keys like `inputData` do not. Keep
 * this list explicit; an over-broad match would redact structural fields (like `a1Ranges`)
 * and an over-narrow one would leak content.
 */
const REDACTED_KEYS: Record<string, true> = {
  value: true,
  values: true,
  data: true,
  rows: true,
  rowvalues: true,
  cellvalues: true,
  cellcontents: true,
  formula: true,
  formulas: true,
  input: true,
  inputs: true,
  incoming: true,
  incomingvalues: true,
  content: true,
  contents: true,
  before: true,
  after: true,
  formulasoverwritten: true,
  payload: true,
  privatekey: true,
  clientemail: true,
  clientsecret: true,
  accesstoken: true,
  refreshtoken: true,
  idtoken: true,
  credentials: true,
};

const redactedKey = (key: string): string => key.replace(/[-_]/g, '').toLowerCase();

/**
 * Strip cell contents, formulas, incoming values and credential material from an arbitrary
 * diagnostic object (dry-run payloads in later waves), returning a new object - the input is
 * never mutated. Policy:
 * - a key in `REDACTED_KEYS` keeps its place but loses its content: scalar values become
 *   `'[REDACTED]'`, arrays keep their length with `null` elements (counts survive), and
 *   nested objects are redacted recursively so shape stays inspectable;
 * - every string that survives is run through `sanitizeMessage`, so the shared secret
 *   patterns (private keys, JWTs, bearer and access tokens) scrub prose too;
 * - structural fields (coordinates, codes, counts, states, messages) pass through untouched;
 * - cyclic references are cut, not crashed on.
 *
 * The generic signature preserves the caller's payload type at the public boundary; the
 * recursion itself walks an `unknown` tree so every access is runtime-narrowed.
 */
export function redactDiagnostic<T>(value: T, seen: WeakSet<object> = new WeakSet()): T {
  // public API boundary: the recursion below validates by typeof/in narrowing, not by cast
  return redactDiagnosticValue(value, seen) as T;
}

function redactDiagnosticValue(value: unknown, seen: WeakSet<object>): unknown {
  if (typeof value === 'string') return sanitizeMessage(value);
  if (typeof value !== 'object' || value === null) return value;
  if (seen.has(value)) return null;
  seen.add(value);
  if (Array.isArray(value)) {
    return value.map((element) => redactDiagnosticValue(element, seen));
  }
  const result: Record<string, unknown> = {};
  for (const [key, inner] of Object.entries(value)) {
    if (REDACTED_KEYS[redactedKey(key)]) {
      if (Array.isArray(inner)) {
        // keep the count, drop the content
        result[key] = inner.map(() => null);
      } else if (inner !== null && typeof inner === 'object') {
        result[key] = redactDiagnosticValue(inner, seen);
      } else if (inner === undefined) {
        continue;
      } else {
        result[key] = '[REDACTED]';
      }
      continue;
    }
    result[key] = redactDiagnosticValue(inner, seen);
  }
  return result;
}

const sanitizeIssueValue = (value: unknown): ErrorIssue['value'] => {
  if (value === null) return null;
  if (typeof value === 'string') return value.length > MAX_VALUE_LENGTH ? `${value.slice(0, MAX_VALUE_LENGTH)}...` : value;
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  return undefined;
};

const sanitizeIssue = (issue: ErrorIssue): ErrorIssue => {
  const safe: ErrorIssue = { message: sanitizeMessage(issue.message) };
  if (issue.row !== undefined) safe.row = issue.row;
  if (issue.column !== undefined) safe.column = issue.column;
  if (issue.a1 !== undefined) safe.a1 = issue.a1;
  if (issue.field !== undefined) safe.field = issue.field;
  if (issue.code !== undefined) safe.code = issue.code;
  // redacted mode keeps the issue's coordinates and code but not the offending value
  if (redactionEnabled) return safe;
  const value = sanitizeIssueValue(issue.value);
  if (value !== undefined) safe.value = value;
  return safe;
};

/** The one failure envelope every command emits in JSON failure mode. */
export interface GSheetErrorEnvelope {
  error: {
    code: GSheetErrorCodeValue;
    message: string;
    retryable: boolean;
    retryAfterMs?: number;
    issues?: ErrorIssue[];
    /**
     * Sanitized mutation-outcome block, present only when the thrown `GSheetError` carries a
     * `mutationOutcome`. States, ranges, counts, phase and retry guidance only - never cell
     * values or formulas. In redacted mode the prose fields (`causeSummary`,
     * `gridGrowth.details`) are dropped too; coordinates/counts/states stay.
     */
    mutation?: SerializedMutationOutcome;
  };
}

/**
 * The central serializer. Builds the envelope from stable metadata and a sanitized message
 * only - causes, request internals and arbitrary error properties never pass through. Issues
 * are carried solely from a `GSheetError`'s own details, and the mutation block solely from a
 * `GSheetError`'s own `mutationOutcome`, projected through `serializeMutationOutcome` (which
 * honors the redaction mode). The `cause` is non-enumerable on `GSheetError` and is never
 * read here, so nested causes - redacted or not - cannot leak into the envelope.
 */
export function toErrorEnvelope(err: unknown): GSheetErrorEnvelope {
  const meta = classifyError(err);
  const rawMessage = err instanceof Error
    ? err.message
    : err !== null && typeof err === 'object' && 'message' in err && typeof err.message === 'string'
      ? err.message
      : '';
  const envelope: GSheetErrorEnvelope = {
    error: {
      code: meta.code,
      message: rawMessage
        ? sanitizeMessage(
            redactionEnabled &&
              err !== null &&
              typeof err === 'object' &&
              'redactedMessage' in err &&
              typeof (err as { redactedMessage?: unknown }).redactedMessage === 'string'
              ? (err as { redactedMessage: string }).redactedMessage
              : rawMessage
          )
        : 'Unexpected error',
      retryable: meta.retryable,
    },
  };
  if (meta.retryAfterMs !== undefined) envelope.error.retryAfterMs = meta.retryAfterMs;
  if (err instanceof GSheetError && err.details?.issues?.length) {
    envelope.error.issues = err.details.issues.map(sanitizeIssue);
  }
  if (err instanceof GSheetError && err.mutationOutcome !== undefined) {
    envelope.error.mutation = serializeMutationOutcome(err.mutationOutcome, {
      redacted: redactionEnabled,
      sanitize: sanitizeMessage,
    });
  }
  return envelope;
}
