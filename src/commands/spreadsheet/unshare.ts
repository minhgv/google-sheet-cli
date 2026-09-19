import { Flags } from '@oclif/core';
import Command, { spreadsheetId } from '../../lib/base-class';

export default class SpreadsheetUnshare extends Command {
  static description =
    'Remove a sharing permission from the spreadsheet, by permission id or grantee email ' +
    '(Drive API, drive.file scope).';

  static examples = [
    `$ gsheet spreadsheet:unshare --spreadsheetId=<id> --permissionId <id>
`,
    `$ gsheet spreadsheet:unshare --spreadsheetId=<id> --email user@example.com
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    permissionId: Flags.string({
      description: 'Permission id from spreadsheet:permissions',
      required: false,
      exclusive: ['email'],
    }),
    email: Flags.string({
      description: 'Grantee email to remove (resolved via the permission list)',
      required: false,
      exclusive: ['permissionId'],
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, permissionId, email, rawOutput },
    } = await this.parse(SpreadsheetUnshare);

    this.start('Removing permission');
    const result = await this.gsheet.unshareSpreadsheet({ permissionId, email }, spreadsheetId);
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else {
      this.log(`Removed permission ${result.permissionId} from ${result.spreadsheetId}`);
    }

    return result;
  }
}
