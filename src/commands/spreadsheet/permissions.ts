import Command, { spreadsheetId } from '../../lib/base-class';

export default class SpreadsheetPermissions extends Command {
  static description = 'List the sharing permissions on the spreadsheet (Drive API, drive.file scope).';

  static examples = [
    `$ gsheet spreadsheet:permissions --spreadsheetId=<id>
`,
    `$ gsheet spreadsheet:permissions --spreadsheetId=<id> --rawOutput
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
  };

  async run() {
    const {
      flags: { spreadsheetId, rawOutput },
    } = await this.parse(SpreadsheetPermissions);

    this.start('Listing permissions');
    const permissions = await this.gsheet.listPermissions(spreadsheetId);
    this.stop();

    if (rawOutput) {
      this.logRaw('', permissions);
    } else {
      for (const p of permissions) {
        const target = p.emailAddress || p.domain || p.displayName || p.type;
        this.log(`${p.id}\t${p.type}\t${p.role}\t${target}`);
      }
    }

    return permissions;
  }
}
