/**
 * Dependency-free error type for refusals caused entirely by local input inspection: range
 * parsing, table framing, key matching, formula-guard preflights. It deliberately imports
 * nothing - the CLI layer and any consumer of the `google-sheet-cli/sheet` entry point can
 * recognize it structurally, without pulling this module's dependencies or oclif into their
 * graph, by reading the stable `code` property.
 */

/** Stable structural marker. Own and enumerable so serializers see it with plain reads. */
export const VALIDATION_CODE = 'VALIDATION';

export class ValidationError extends Error {
  /** Structural classification marker; never changes, never carries request-specific data. */
  readonly code = VALIDATION_CODE;
  /**
   * Message shown INSTEAD of `message` when redaction is enabled (see cli-errors'
   * GSheetErrorOptions.redactedMessage). Read structurally by the envelope serializer, so
   * this class stays dependency-free.
   */
  readonly redactedMessage?: string;

  constructor(message: string, redactedMessage?: string) {
    super(message);
    this.name = 'ValidationError';
    if (redactedMessage !== undefined) this.redactedMessage = redactedMessage;
  }
}
