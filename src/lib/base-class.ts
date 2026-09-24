import { Args, Command, Flags } from '@oclif/core';
import { FlagInput } from '@oclif/core/interfaces';
import { ux } from '@oclif/core/ux';
import { createInterface } from 'readline';
import { normalizeCredentials } from './credentials';
import { GSheetError, GSheetErrorCode, hasTransportShape, isRedactionEnabled, redactDiagnostic, setRedactionEnabled, toErrorEnvelope } from './cli-errors';
import * as factory from './factory';
import GoogleSheet, { GoogleSheetCli } from './google-sheet';

export const spreadsheetId = Flags.string({
  char: 's',
  description: 'ID of the spreadsheet to use',
  required: true,
  env: 'SPREADSHEET_ID',
});

export const worksheetTitle = Flags.string({
  char: 't',
  description: 'Title of the worksheet to use',
  required: true,
  env: 'WORKSHEET_TITLE',
});

/**
 * Optional variants of the shared required flags for dual-backend commands: a
 * `--workbook` invocation must not be forced to name a spreadsheet, and `-t`
 * doubles as the sheet selector inside a local workbook. Same chars and env
 * bindings as the shared definitions (the data:export-csv precedent).
 */
export const optionalSpreadsheetId = Flags.string({
  char: 's',
  description: 'ID of the spreadsheet to use (Google Sheets target)',
  required: false,
  env: 'SPREADSHEET_ID',
});

export const optionalWorksheetTitle = Flags.string({
  char: 't',
  description: 'Title of the worksheet to use (Google Sheets target; also selects the sheet inside --workbook)',
  required: false,
  env: 'WORKSHEET_TITLE',
});

/**
 * Local-workbook target flags shared by every dual-backend mutation command.
 * `--workbook` selects the local backend; the mutation lands on `--output` or
 * `--inPlace`, never implicitly on the source file.
 */
export const workbookTargetFlags = {
  workbook: Flags.string({
    description:
      'Path to a local .xlsx workbook to mutate instead of the Google Sheets target. Long name only: the shared short -f belongs to --credentialsFile',
    required: false,
  }),
  output: Flags.string({
    char: 'o',
    description: 'Destination path for the modified XLSX file (without it and without --inPlace the mutation is refused unless --dryRun)',
    required: false,
  }),
  inPlace: Flags.boolean({
    description: 'Modify the --workbook file in place (a .bak backup is written first)',
    required: false,
  }),
  discardUnsupported: Flags.boolean({
    description:
      'Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would be dropped by the local engine',
    required: false,
  }),
};

export const valueInputOption = Flags.string({
  char: 'v',
  description: 'The style of the input ("RAW" or "USER_ENTERED")',
  required: false,
  options: [GoogleSheetCli.ValueInputOption.RAW, GoogleSheetCli.ValueInputOption.USER_ENTERED],
  default: GoogleSheetCli.ValueInputOption.RAW,
  env: 'VALUE_INPUT_OPTION',
});

export const useOauth = Flags.boolean({
  helpGroup: 'Authentication',
  description: 'Use OAuth 2.0 user authentication instead of service account',
  required: false,
  env: 'GSHEET_USE_OAUTH',
});

export const clientSecretFile = Flags.string({
  helpGroup: 'Authentication',
  description: 'Path to OAuth 2.0 client_secret.json (Desktop App type)',
  required: false,
  env: 'GSHEET_CLIENT_SECRET_FILE',
});

export const clientEmail = Flags.string({
  helpGroup: 'Authentication',
  char: 'c',
  env: 'GSHEET_CLIENT_EMAIL',
  description: 'The client email to use for authentication. Uses the GSHEET_CLIENT_EMAIL env variable if not provided.',
  required: false,
});

export const privateKey = Flags.string({
  helpGroup: 'Authentication',
  char: 'p',
  env: 'GSHEET_PRIVATE_KEY',
  description: 'The private key to use for authentication. Uses the GSHEET_PRIVATE_KEY env variable if not provided.',
  required: false,
});

export const credentialsFile = Flags.string({
  helpGroup: 'Authentication',
  char: 'f',
  env: 'GSHEET_CREDENTIALS_FILE',
  description:
    'Path to the service account JSON file to read the credentials from. Uses the GSHEET_CREDENTIALS_FILE env variable if not provided. The clientEmail and privateKey flags take precedence.',
  required: false,
});

export const googleAuthFlags = {
  clientEmail,
  privateKey,
  credentialsFile,
  useOauth,
  clientSecretFile,
};

export interface GoogleAuthFlags {
  clientEmail?: string;
  privateKey?: string;
  credentialsFile?: string;
  useOauth?: boolean;
  clientSecretFile?: string;
}
export const data = Args.string({
  name: 'data',
  description: 'The data to be used as a JSON string - nested array [["1", "2", "3"]]',
  required: true,
});

export const optionalData = Args.string({
  name: 'data',
  description: 'The data to be used as a JSON string - nested array [["1", "2", "3"]]',
  required: false,
  // oclif 5's parser otherwise assigns piped stdin to this positional, which collides with
  // `--input -`: the pipe would reach both sources and resolveDataMatrix rejects that. Stdin
  // must go through `--input -` (see resolveDataMatrix), so the positional never auto-fills.
  ignoreStdin: true,
});
export interface CommonFlags extends GoogleAuthFlags {
  rawOutput?: boolean;
  json?: boolean;
  redacted?: boolean;
  help?: void;
}

/**
 * Ask for one secret on the terminal without echoing it back.
 *
 * `@oclif/core` 2 had `ux.prompt(message, { type: 'hide' })` for this; core 5 dropped the whole
 * prompt module along with its `password-prompt` dependency, so the contract that dependency
 * provided is kept here rather than taking a new one for it. Three parts of that contract are
 * load-bearing and easy to lose:
 *
 * - everything goes to **stderr**. stdout is the command's output, and `--rawOutput` promises it
 *   is JSON; a prompt written there lands in the middle of a `| jq` pipeline.
 * - an empty answer is refused and the question asked again, rather than being handed on as an
 *   empty credential that fails later as an opaque authentication error.
 * - the muted `write` is put back whatever happens, including on an aborted prompt (EOF, Ctrl-D),
 *   which closes the interface without ever calling back. Leaving it swallowed would silence the
 *   process for good.
 *
 * @param {string} message
 * @returns {Promise<string>}
 */
export const hiddenPrompt = (message: string): Promise<string> =>
  new Promise((resolve, reject) => {
    const stderr = process.stderr;
    // the function itself, not a bound copy: restoring a copy would leave a different `write` on
    // the stream every time, which stacks up and defeats anyone else swapping it
    const write = stderr.write;
    let muted = false;
    let answered = false;

    const restore = () => {
      muted = false;
      (stderr as unknown as { write: typeof stderr.write }).write = write;
    };

    // readline echoes what is typed; swallow those writes while an answer is being entered
    (stderr as unknown as { write: (...args: unknown[]) => boolean }).write = (...args: unknown[]) =>
      muted ? true : (write as (...a: unknown[]) => boolean).apply(stderr, args);
    const rl = createInterface({ input: process.stdin, output: stderr, terminal: true });

    rl.on('close', () => {
      restore();
      process.stdin.pause();
      if (!answered) reject(new Error('No input'));
    });

    const ask = () => {
      // the question itself has to reach the terminal, so unmute around writing it
      muted = false;
      rl.question(`${message}: `, (answer) => {
        if (answer === '') {
          ask();
          return;
        }
        answered = true;
        restore();
        stderr.write('\n');
        rl.close();
        resolve(answer);
      });
      muted = true;
    };

    ask();
  });

export async function resolveGoogleSheetAuth(
  flags: GoogleAuthFlags,
  options?: { prompt?: boolean }
): Promise<GoogleSheet> {
  const gsheet = factory.createGoogleSheet();
  const allowPrompt = options?.prompt !== false;

  // OAuth 2.0 takes precedence when explicitly requested or when no service account credentials exist
  const hasServiceAccountCreds = Boolean(flags?.clientEmail || flags?.privateKey || flags?.credentialsFile);
  const useOauth = Boolean(
    flags?.useOauth ||
      (!hasServiceAccountCreds &&
        !process.env.GSHEET_CLIENT_EMAIL &&
        !process.env.GSHEET_PRIVATE_KEY &&
        !process.env.GSHEET_CREDENTIALS_FILE)
  );

  if (useOauth) {
    try {
      await gsheet.authorizeOAuth(flags?.clientSecretFile);
    } catch (err) {
      // Service-account local failures carry AUTH_REQUIRED (below); the OAuth path must not
      // blanket-wrap. A failure that still carries an HTTP status or network evidence is
      // rethrown untouched, so a 403 stays FORBIDDEN, a 429 stays RATE_LIMITED and a dead
      // connection stays NETWORK in the envelope. Only genuinely local problems - a missing
      // or unreadable client secret, no stored token, an expired token with no refresh token -
      // become AUTH_REQUIRED, with the original preserved as the (never-serialized) cause.
      if (err instanceof GSheetError || hasTransportShape(err)) throw err;
      throw new GSheetError(
        GSheetErrorCode.AUTH_REQUIRED,
        err instanceof Error && err.message ? err.message : 'Google OAuth authentication could not be established.',
        { cause: err }
      );
    }
  } else {
    // Only prompt for what the flags, the env and the credentials file left missing. Local
    // credential failures carry AUTH_REQUIRED so the JSON envelope classifies them apart from
    // remote authorization rejections (UNAUTHORIZED/FORBIDDEN come from HTTP status later).
    let credentials;
    try {
      credentials = normalizeCredentials({
        client_email: flags?.clientEmail,
        private_key: flags?.privateKey,
        credentialsFile: flags?.credentialsFile,
      });
    } catch (err) {
      throw new GSheetError(
        GSheetErrorCode.AUTH_REQUIRED,
        err instanceof Error && err.message ? err.message : 'Google service account credentials could not be read.',
        { cause: err }
      );
    }

    let client_email = credentials.client_email;
    let private_key = credentials.private_key;

    const promptSecret = async (message: string): Promise<string> => {
      try {
        return await hiddenPrompt(message);
      } catch (err) {
        throw new GSheetError(GSheetErrorCode.AUTH_REQUIRED, err instanceof Error && err.message ? err.message : 'No input', { cause: err });
      }
    };

    if (!client_email && allowPrompt) {
      client_email = await promptSecret('What is your client email?');
    }
    if (!private_key && allowPrompt) {
      private_key = await promptSecret('What is your private key?');
    }

    if (!client_email || !private_key) {
      throw new GSheetError(GSheetErrorCode.AUTH_REQUIRED, 'Google Sheets authentication requires client_email and private_key.');
    }

    await gsheet.authorize({
      client_email,
      private_key,
    });
  }

  return gsheet;
}

const JSON_FAILURE_LONG: Record<string, true> = {
  '--json': true,
  '--json=true': true,
  '--rawOutput': true,
  '--rawOutput=true': true,
};

/**
 * Decide JSON failure mode from raw argv alone, for failures that happen before any flag was
 * parsed. The scan mirrors what @oclif/core's parser does with each token so operator data is
 * never mistaken for a trigger: everything after the `--` passthrough marker is data; a
 * value-taking long option without `=` consumes the next element even when that element looks
 * like a flag; and inside a short cluster an option char owns the rest of the cluster (`-tj`
 * is title "j") or, at the cluster's end, the next argv element (`-t -j` is always a failed
 * invocation whose `-j` sits in the value slot). Only a `j`/`r` in an actual flag position
 * counts, and explicit `--json=false` / `--rawOutput=false` overrides do not trigger.
 */
const jsonFailureModeFromArgv = (argv: string[], flags: FlagInput | undefined): boolean => {
  const optionChars = new Set<string>();
  const optionNames = new Set<string>();
  for (const [name, flag] of Object.entries(flags ?? {})) {
    const f = flag as { type?: string; char?: string } | undefined;
    if (f && typeof f === 'object' && f.type === 'option') {
      optionNames.add(name);
      if (typeof f.char === 'string') optionChars.add(f.char);
    }
  }
  const passThrough = argv.indexOf('--');
  const limit = passThrough === -1 ? argv.length : passThrough;
  let skipNext = false;
  for (let i = 0; i < limit; i++) {
    if (skipNext) {
      skipNext = false;
      continue;
    }
    const arg = argv[i];
    if (JSON_FAILURE_LONG[arg]) return true;
    if (arg.startsWith('--')) {
      if (arg.indexOf('=') === -1 && optionNames.has(arg.slice(2))) skipNext = true;
      continue;
    }
    if (arg.length < 2 || !arg.startsWith('-')) continue;
    for (let c = 1; c < arg.length; c++) {
      const ch = arg[c];
      if (optionChars.has(ch)) {
        // the value is the rest of this cluster, or the next element when the option char
        // ends it - either way a j/r in that position is the value, not a trigger
        if (c === arg.length - 1) skipNext = true;
        break;
      }
      if (ch === 'j' || ch === 'r') return true;
    }
  }
  return false;
};

export default abstract class extends Command {
  private rawLogs: boolean = false;
  private jsonFailures: boolean = false;
  public gsheet!: GoogleSheet;

  /**
   * Whether failures must come out as the machine-readable envelope instead of oclif's human
   * error. Opt-in through `--json`/`-j` or `--rawOutput`/`-r` - the same flags that ask for
   * JSON on the success side - so a caller never has to guess which mode errors arrive in.
   *
   * This must work even when argv parsing failed before any flag was read (an unknown flag,
   * a missing required one), so alongside the parsed state it peeks at `this.argv`. Unlike
   * @oclif/core's own naive `--json` indexOf scan, the peek knows how the parser assigns
   * values, so data is never mistaken for a trigger - see jsonFailureModeFromArgv.
   */
  private failureJsonMode(): boolean {
    if (this.rawLogs || this.jsonFailures) return true;
    return jsonFailureModeFromArgv(this.argv, this.ctor.flags);
  }

  static flags = {
    help: Flags.help({ char: 'h' }),
    rawOutput: Flags.boolean({
      char: 'r',
      description: 'Get the raw output as a JSON string',
      default: false,
      required: false,
    }),
    json: Flags.boolean({
      char: 'j',
      description:
        'Report failures as a machine-readable JSON envelope on stderr (exit code 1 on failure). Success output is unchanged - use --rawOutput for JSON success.',
      default: false,
      required: false,
    }),
    redacted: Flags.boolean({
      description:
        'Strip cell contents, formulas, incoming values and credentials from error envelopes and dry-run diagnostics before they are written. Coordinates, counts, statuses and outcome states are kept.',
      default: false,
      required: false,
      env: 'GSHEET_REDACTED',
    }),
    ...googleAuthFlags,
  } as FlagInput<CommonFlags>;
  async start(message: string) {
    if (!this.rawLogs && !this.failureJsonMode()) {
      ux.action.start(message);
    }
  }

  async stop(message?: string) {
    if (!this.rawLogs && !this.failureJsonMode()) {
      ux.action.stop(message);
    }
  }

  async logRaw(message: string, raw?: unknown) {
    if (this.rawLogs) {
      // redaction strips cell contents/formulas/incoming values from receipts (dry-run
      // previews included) before serialization; coordinates, counts and states survive
      const payload = isRedactionEnabled() ? redactDiagnostic(raw) : raw;
      this.log(JSON.stringify(payload, null, 2));
    } else {
      this.log(message);
    }
  }

  async init() {
    // do some initialization
    const parsed = await this.parse(this.constructor as typeof Command);
    const flags = parsed.flags as unknown as CommonFlags;
    this.rawLogs = !!flags?.rawOutput;
    this.jsonFailures = !!flags?.json;
    // redaction must be on before any command logic (and before the auth errors init itself
    // can throw) so envelopes for those failures are already stripped
    setRedactionEnabled(Boolean(flags?.redacted));
    this.gsheet = await resolveGoogleSheetAuth(flags, { prompt: true });
  }

  async catch(err: Error) {
    // `ux.action.start()` replaces process.stdout.write and process.stderr.write with buffering
    // stubs, and only `stop()` puts them back and flushes. An error thrown while the spinner is
    // running is therefore written into a buffer nobody empties, and the process exits with the
    // right code and nothing printed. oclif's own `Command.catch` stops the action for exactly
    // this reason; overriding it took that away, and `@oclif/core` 2's process-exit hook, which
    // used to flush the buffer regardless, is gone in 5. Stopping an action that was never
    // started returns immediately, so this is correct on every path into `catch` - a parse
    // failure, a credential failure, or a command that already called `stop()`.
    try {
      ux.action.stop();
    } catch {
      // a rendering failure must never swallow the error we are here to report
    }
    if (this.failureJsonMode()) {
      // The one envelope for every failure in this mode - parse, auth, API and network alike -
      // on stderr, where --rawOutput callers already keep their JSON on stdout unmixed.
      process.stderr.write(`${JSON.stringify(toErrorEnvelope(err), null, 2)}\n`);
      // Errors.exit throws an ExitError: the top-level runner turns it into exit code 1
      // without printing anything, so the envelope above is the only error output.
      this.exit(1);
    }
    this.error(err, { exit: 1 });
  }
}
