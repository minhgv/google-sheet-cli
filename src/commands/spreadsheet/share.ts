import { Flags } from '@oclif/core';
import Command, { spreadsheetId } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class SpreadsheetShare extends Command {
  static description =
    'Share the spreadsheet with users, a domain, or anyone with the link. ' +
    'Uses the Drive API (drive.file scope) — OAuth tokens issued before that scope was added must re-run `gsheet auth:login`. ' +
    'Under drive.file only files this app created or has opened are visible.';

  static examples = [
    `$ gsheet spreadsheet:share --spreadsheetId=<id> --email user@example.com --role writer
`,
    `$ gsheet spreadsheet:share --spreadsheetId=<id> --email a@x.com --email b@x.com --role reader
`,
    `$ gsheet spreadsheet:share --spreadsheetId=<id> --domain example.com --role commenter
`,
    `$ gsheet spreadsheet:share --spreadsheetId=<id> --anyone --role reader
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    email: Flags.string({
      description: 'Email address to grant access (repeatable)',
      required: false,
      multiple: true,
      exclusive: ['domain', 'anyone'],
    }),
    domain: Flags.string({
      description: 'Google Workspace domain to grant access',
      required: false,
      exclusive: ['email', 'anyone'],
    }),
    anyone: Flags.boolean({
      description: 'Grant access to anyone with the link',
      required: false,
      exclusive: ['email', 'domain'],
    }),
    type: Flags.string({
      description: 'Grantee type (inferred from --email/--domain/--anyone when omitted)',
      options: ['user', 'group', 'domain', 'anyone'],
      required: false,
    }),
    role: Flags.string({
      description: 'Access role',
      options: ['reader', 'commenter', 'writer'],
      default: 'reader',
      required: false,
    }),
    notify: Flags.boolean({
      description: 'Send Google notification email (default: off — agents should not spam)',
      required: false,
    }),
    message: Flags.string({
      description: 'Message attached to the notification email (requires --notify)',
      required: false,
      dependsOn: ['notify'],
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, email, domain, anyone, type, role, notify, message, rawOutput },
    } = await this.parse(SpreadsheetShare);

    this.start('Sharing spreadsheet');
    const result = await this.gsheet.shareSpreadsheet(
      {
        emails: email,
        domain,
        anyone,
        type: type as GoogleSheetCli.ShareOptions['type'],
        role: role as GoogleSheetCli.ShareOptions['role'],
        notify,
        message,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else {
      for (const granted of result.granted) {
        const target = granted.emailAddress || granted.domain || 'anyone with the link';
        this.log(`Granted ${granted.role} to ${target} (${granted.type}, permission ${granted.id})`);
      }
    }

    return result;
  }
}
