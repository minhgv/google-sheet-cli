import { Args, Command, Flags } from '@oclif/core';
import { FlagInput } from '@oclif/core/interfaces';
import { ux } from '@oclif/core/ux';
import { createInterface } from 'readline';
import { normalizeCredentials } from './credentials';
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
});
export interface CommonFlags extends GoogleAuthFlags {
  rawOutput?: boolean;
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
    await gsheet.authorizeOAuth(flags?.clientSecretFile);
  } else {
    // Only prompt for what the flags, the env and the credentials file left missing.
    const credentials = normalizeCredentials({
      client_email: flags?.clientEmail,
      private_key: flags?.privateKey,
      credentialsFile: flags?.credentialsFile,
    });

    let client_email = credentials.client_email;
    let private_key = credentials.private_key;

    if (!client_email && allowPrompt) {
      client_email = await hiddenPrompt('What is your client email?');
    }
    if (!private_key && allowPrompt) {
      private_key = await hiddenPrompt('What is your private key?');
    }

    if (!client_email || !private_key) {
      throw new Error('Google Sheets authentication requires client_email and private_key.');
    }

    await gsheet.authorize({
      client_email,
      private_key,
    });
  }

  return gsheet;
}

export default abstract class extends Command {
  private rawLogs: boolean = false;
  public gsheet!: GoogleSheet;

  static flags = {
    help: Flags.help({ char: 'h' }),
    rawOutput: Flags.boolean({
      char: 'r',
      description: 'Get the raw output as a JSON string',
      default: false,
      required: false,
    }),
    ...googleAuthFlags,
  } as FlagInput<CommonFlags>;
  async start(message: string) {
    if (!this.rawLogs) {
      ux.action.start(message);
    }
  }

  async stop(message?: string) {
    if (!this.rawLogs) {
      ux.action.stop(message);
    }
  }

  async logRaw(message: string, raw?: unknown) {
    if (this.rawLogs) {
      this.log(JSON.stringify(raw, null, 2));
    } else {
      this.log(message);
    }
  }

  async init() {
    // do some initialization
    const parsed = await this.parse(this.constructor as typeof Command);
    const flags = parsed.flags as unknown as CommonFlags;
    this.rawLogs = !!flags?.rawOutput;
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
    this.error(err, { exit: 1 });
  }
}
