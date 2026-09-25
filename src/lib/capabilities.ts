/**
 * Machine-readable capability contract for the gateway's two storage backends.
 *
 * The document is a static literal assembled from the facts this build ships: which command
 * surface exists, how data reaches each backend, what formulas do on each side, which guards
 * refuse destructive operations, and why neither backend can promise transactional semantics.
 * It probes nothing at runtime and never touches credentials, so it can be printed or consumed
 * by an agent before any authentication exists.
 *
 * Static support is deliberately separated from runtime authorization: an operation listed here
 * means the CLI surface exists in this build, never that a caller is entitled to run it.
 *
 * The schema version is owned by this document, independent of the package version: it moves
 * when the document's shape or the meaning of its fields changes, and for no other reason.
 *
 * The 2026-09 gateway update (spreadsheet:list, data:export-csv, the error.mutation envelope,
 * the --redacted flag) changed only the CONTENT of existing fields - operations, outputForms,
 * visibility and mutationLimits grow entries; no field was added, removed or reinterpreted -
 * so schemaVersion stays 1. Consumers that string-match or enumerate entries keep working;
 * only a shape change will move the version. The later local-XLSX update (--updateRefs reference
 * rewriting, format:cells / grid:freeze / grid:resize / grid:hide and worksheet:add|remove|rename
 * on --workbook, workbook:names) follows the same rule: content-only growth, schemaVersion stays 1.
 */

export const CAPABILITY_SCHEMA_VERSION = 1;

export type BackendId = 'google-sheets' | 'local-xlsx';

export interface CapabilityFormulaSemantics {
  /** How formulas enter the backend. */
  readonly write: string;
  /** Whether, and where, formula results are computed. */
  readonly recalculation: string;
}

export interface CapabilityBackend {
  readonly id: BackendId;
  /** oclif command ids executable against this backend. */
  readonly operations: readonly string[];
  /** How tabular input reaches the commands. */
  readonly inputForms: readonly string[];
  /** What the commands emit on success and on failure. */
  readonly outputForms: readonly string[];
  readonly formulaSemantics: CapabilityFormulaSemantics;
  /** Guards that refuse destructive or ambiguous operations before they happen. */
  readonly destructiveGuards: readonly string[];
  /** Why a caller cannot treat this backend as transactional or exactly-once. */
  readonly mutationLimits: readonly string[];
  /**
   * Workbook features the backend does not preserve. Only meaningful for the local backend;
   * empty for google-sheets, whose data lives in Google's own storage.
   */
  readonly fidelityExclusions: readonly string[];
}

export interface CapabilityDocument {
  readonly schemaVersion: number;
  /** Static support vs runtime authorization, stated once for the whole document. */
  readonly supportModel: string;
  /** What the authorization scopes can and cannot see - the drive.file boundary. */
  readonly visibility: readonly string[];
  readonly backends: Readonly<Record<BackendId, CapabilityBackend>>;
}

const CAPABILITY_DOCUMENT: CapabilityDocument = {
  schemaVersion: CAPABILITY_SCHEMA_VERSION,
  supportModel:
    'Support described here is static: it states which command surface exists in this build, not what a caller may do. ' +
    'Every google-sheets operation still requires valid credentials and sufficient Google-side permissions at run time; ' +
    'local-xlsx operations require read and write access to the local files. Authorization is only checked when a command runs.',
  visibility: [
    'Google authorization uses the drive.file OAuth scope: Drive-based operations (spreadsheet:list discovery, ' +
      'spreadsheet:copy, spreadsheet:export) can only see spreadsheets this application created or opened directly.',
    'Spreadsheets created by other applications are invisible to spreadsheet:list even when they exist, so an empty ' +
      'discovery result does not prove a spreadsheet is absent - address such files by ID instead.',
    'spreadsheet:list never selects a spreadsheet by itself: duplicate titles stay separate rows, a title is never ' +
      'an id, and the caller picks the id.',
  ],
  backends: {
    'google-sheets': {
      id: 'google-sheets',
      operations: [
        'data:append',
        'data:append-table',
        'data:batch-get',
        'data:batch-update',
        'data:clear',
        'data:export-csv',
        'data:find',
        'data:get',
        'data:schema',
        'data:update',
        'data:upsert',
        'data:validate',
        'worksheet:add',
        'worksheet:copy',
        'worksheet:get',
        'worksheet:remove',
        'worksheet:rename',
        'spreadsheet:add',
        'spreadsheet:copy',
        'spreadsheet:export',
        'spreadsheet:get',
        'spreadsheet:list',
        'spreadsheet:permissions',
        'spreadsheet:share',
        'spreadsheet:unshare',
        'grid:delete',
        'grid:freeze',
        'grid:hide',
        'grid:insert',
        'grid:resize',
        'format:cells',
        'format:merge',
        'report:run',
        'auth:login',
        'auth:logout',
        'auth:status',
      ],
      inputForms: [
        'Positional JSON string: a nested 2D array (e.g. [["1","2"]]) for the data commands.',
        '--input <path> pointing at a .json or .csv file, or --input - for standard input; the format comes from ' +
          '--inputFormat json|csv, the file extension, or content sniffing.',
        'Addresses and metadata as flags: --spreadsheetId, --worksheetTitle, --range, or JSON flag values such as ' +
          '--ranges \'["Sheet1!A1:B2"]\'.',
      ],
      outputForms: [
        'Human-readable text on stdout by default.',
        '--rawOutput prints the command receipt as machine-readable JSON on stdout.',
        'On failure under --json or --rawOutput, a single structured envelope ' +
          '{"error":{"code","message","retryable","retryAfterMs?","issues?","mutation?"}} goes to stderr with exit code 1.',
        'For multi-request writes (data:batch-update, data:upsert) the envelope may carry error.mutation: per-request ' +
          'outcomes acknowledged, rejected, unknown (dispatch may have succeeded) or not-attempted, each with the ' +
          "caller's logical A1 ranges, grid-growth effects tracked separately, the operation phase, and conservative " +
          'retryGuidance (safe-replay / verify-then-replay / never-blind-replay). States, ranges and counts only - ' +
          'never cell contents or formulas.',
        'data:export-csv writes deterministic RFC 4180 CSV of one worksheet or bounded range to stdout or --output: ' +
          '--mode raw|formatted|formula selects the cell rendering, --injection safe|preserve selects the ' +
          'formula-injection policy (safe is the default and prefixes dangerous leading characters with an apostrophe; ' +
          'preserve emits them byte-faithfully and says so in the receipt). With --output the CSV is written atomically ' +
          'and a receipt (rows, columns, range, mode, policy, transformed cells, warnings) is printed; without it, ' +
          'stdout carries only the CSV bytes and warnings go to stderr.',
        'The global --redacted flag strips cell contents, formulas, incoming values and credentials from error ' +
          'envelopes and dry-run diagnostics before they are written - coordinates, counts, statuses and mutation ' +
          'outcome states are kept. Normal data reads and exports stay data-bearing.',
        'spreadsheet:export writes binary pdf or xlsx files.',
      ],
      formulaSemantics: {
        write:
          'Values are written through the Sheets values API with --valueInputOption RAW|USER_ENTERED (default RAW): ' +
          'USER_ENTERED parses strings that begin with "=" into formulas and applies number and date parsing; RAW stores the literal string.',
        recalculation:
          'Google Sheets evaluates and recalculates formulas server-side on write; the CLI never computes formula results and reads return the server-computed values.',
      },
      destructiveGuards: [
        '--dryRun previews a mutation without writing it (data mutation and formatting commands).',
        'data:clear requires an explicit A1 range bounded on both axes; there is no unbounded or whole-worksheet clear.',
        'data:clear refuses to overwrite cells that contain formulas unless --overwriteFormulas is passed.',
        'spreadsheet:export refuses to replace an existing destination file without --overwrite.',
      ],
      mutationLimits: [
        'Single-writer: no locking and no concurrent-writer detection; callers must serialize mutations across processes and machines.',
        'Multi-request operations (batch updates, upserts, header-aware appends) are not transactional: a failure mid-sequence can leave partially applied state that the CLI does not roll back.',
        'A mutation whose HTTP response is lost has an unknown outcome - the write may have succeeded. Retry guidance is conservative: reads are replayable, deterministic overwrites can be re-issued, and non-idempotent operations (append, create, copy) must be verified before replay, never blind-replayed. Failing operations surface this as the structured error.mutation block (see outputForms).',
        'Range writes can grow the worksheet grid to fit the data; grid growth is part of a mutation\'s effect.',
      ],
      fidelityExclusions: [],
    },
    'local-xlsx': {
      id: 'local-xlsx',
      operations: [
        'data:clear',
        'data:export-csv',
        'format:cells',
        'format:merge',
        'grid:delete',
        'grid:freeze',
        'grid:hide',
        'grid:insert',
        'grid:resize',
        'workbook:find',
        'workbook:inspect',
        'workbook:names',
        'workbook:read',
        'workbook:write',
        'worksheet:add',
        'worksheet:remove',
        'worksheet:rename',
        'report:run',
      ],
      inputForms: [
        'Positional JSON string: a nested 2D array or a full ReportDocument object (workbook:write).',
        '--cells JSON object {"A1": value} or [{"a1","value"}] list for sparse per-cell writes that leave every other cell untouched (workbook:write).',
        '--input <path> with JSON or CSV data, or --input - for standard input, using the same format detection as the Sheets commands (workbook:write, report:run).',
        'Local .xlsx files: template input via --file, reads via --file plus a required --range (workbook:read), cell search via --file (workbook:find), inspection via --file (workbook:inspect), defined names via --file with a positional list|add|remove action (workbook:names), report sources via --sourceWorkbook (report:run).',
        'Dual-backend mutations: grid:insert, grid:delete, grid:freeze, grid:resize, grid:hide, format:cells, ' +
          'format:merge, worksheet:add, worksheet:remove, worksheet:rename and data:clear accept --workbook <file.xlsx> ' +
          'instead of --spreadsheetId; -t/--worksheetTitle selects the sheet inside the workbook, and the result lands ' +
          'on --output or --inPlace (never implicitly on the source).',
      ],
      outputForms: [
        'Human-readable text on stdout by default.',
        '--rawOutput prints the command receipt as machine-readable JSON on stdout; the same failure envelope as the Sheets commands applies.',
        'data:export-csv writes the same RFC 4180 CSV as the Sheets source, reading through the local engine: formula cells carry their cached results - never recalculated values - and cache-freshness caveats are reported as warnings.',
        'Writes are published atomically: the destination is replaced by rename only after a complete serialization; --inPlace edits keep an automatic .bak backup.',
        'workbook:inspect reports per-sheet merged ranges, formula-cell counts (full address list with --includeFormulaCells), frozen panes, data validations and conditional-formatting ranges.',
        'workbook:find returns the same {matchCount, truncated, matches:[{a1,row,column,columnLetter,value,rowValues?}]} shape as data:find.',
        'grid:insert/grid:delete --workbook with --updateRefs report refsRewritten and refsBroken in the receipt (references that now point at #REF!).',
      ],
      formulaSemantics: {
        write:
          'Formula strings are stored as ExcelJS formula cells; the caller may attach a cached result to the formula, and template cells keep their stored formulas.',
        recalculation:
          'Formulas are never recalculated: there is no calculation engine. Reads return the cached result stored in the file and report per-cell freshness ("cached" vs "unknown"); a formula without a stored cache surfaces as "unknown", never as a recomputed value.',
      },
      destructiveGuards: [
        'Formula-overwrite guard: applying data over cells that already contain formulas refuses the whole document unless --overwriteFormulas is passed; writing an identical (normalized) formula again is allowed as idempotent. The same guard covers --cells sparse writes and data:clear.',
        'Unsupported-feature preflight: loading detects features ExcelJS would drop (see fidelityExclusions), and saving such a workbook refuses unless --discardUnsupported is passed explicitly.',
        'grid:insert/grid:delete on --workbook refuse when a merged range intersects the splice boundary; --force adjusts the merge (insert extends it, delete shrinks or drops it).',
        'worksheet:add/remove/rename on --workbook refuse duplicate sheet titles, removing the last visible sheet, and renaming onto an existing title.',
        'format:cells on --workbook refuses --numberFormatType (the type is inferred from the --numberFormat pattern) and --wrapStrategy other than WRAP (WRAP maps to wrap text); --wrapText is local-only and refuses with --spreadsheetId.',
        '--dryRun previews the mutation without saving (workbook:write, grid:insert, grid:delete, format:merge, data:clear).',
        '--overwrite consent is required to replace an existing destination file.',
        'Saves are atomic: a failed serialization leaves the previous file in place.',
      ],
      mutationLimits: [
        'Single-writer: no file locking; concurrent writers race at the filesystem level.',
        'Changes are applied to an in-memory copy and published by one rename, so on-disk state never shows a partial write; multi-cell applies are validated as a whole and refuse as a whole on conflicts.',
        'Bounded inputs: the loader enforces file-size and archive-entry limits (default 50 MB compressed).',
        'Formula results are never recomputed; consumers depend on caches written by a real spreadsheet engine (see formulaSemantics).',
        'Formula references and defined names are rewritten by grid:insert/grid:delete only with --updateRefs (local backend): same-sheet references are rewritten - references fully inside a deleted span become #REF! - and cross-sheet references are left untouched; the receipt reports refsRewritten and refsBroken. Without --updateRefs nothing is rewritten and the receipt reports how many formula cells may reference the shifted region; data validations / conditional formatting ranges are never adjusted.',
      ],
      fidelityExclusions: [
        'VBA macros, macro sheets and ActiveX control properties (xl/vbaProject.bin, xl/vbaProjectSignature.bin, xl/macrosheets/, xl/ctrlProps/)',
        'Native charts and chartsheets (xl/charts/, xl/chartsheets/, chart drawings)',
        'Pivot tables and pivot caches (xl/pivotTables/, xl/pivotCache/)',
        'External workbook links (xl/externalLinks/)',
        'Embedded OLE objects and packages (xl/embeddings/)',
        'Data connections and Power Query (xl/connections.xml, xl/queryTables/)',
        'Custom XML parts (customXml/)',
        'Slicers and slicer caches (xl/slicers/, xl/slicerCaches/)',
        'Data model and rich-data extensions (xl/model/, xl/richData/, xl/metadata.xml)',
        'These parts are detected by load-time preflight; a save that would drop them refuses instead (see destructiveGuards).',
      ],
    },
  },
};

const deepFreeze = <T>(value: T): T => {
  if (value !== null && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child);
  }
  return value;
};

/**
 * Return the capability document. Pure and deterministic: the same deeply-frozen literal every
 * time, so consumers can cache it, deep-compare it or hash its JSON without any caveat.
 */
export const buildCapabilityDocument = (): CapabilityDocument => deepFreeze(CAPABILITY_DOCUMENT);
