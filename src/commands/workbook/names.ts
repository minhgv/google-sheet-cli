import { Args, Command, Flags } from '@oclif/core';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class WorkbookNames extends Command {
  static description =
    'List, add or remove defined names (named ranges) in a local XLSX workbook. ' +
    'Run without an action to list every defined name, including the managed ranges the CLI records. ' +
    'add requires --name and --refersTo (a sheet-qualified A1 range); remove requires --name.';

  static examples = [
    `$ gsheet workbook:names --file=book.xlsx
`,
    `$ gsheet workbook:names --file=book.xlsx add --name Amounts --refersTo "Data!$B$2:$B$4" --inPlace
`,
    `$ gsheet workbook:names --file=book.xlsx remove --name Amounts --inPlace
`,
  ];

  static flags = {
    help: Flags.help({ char: 'h' }),
    file: Flags.string({
      char: 'f',
      description: 'Path to the local XLSX workbook',
      required: true,
    }),
    output: Flags.string({
      char: 'o',
      description: 'Destination path for the modified XLSX file (without it and without --inPlace the mutation is refused unless --dryRun)',
      required: false,
    }),
    inPlace: Flags.boolean({
      description: 'Modify the --file workbook in place (a .bak backup is written first)',
      required: false,
    }),
    discardUnsupported: Flags.boolean({
      description:
        'Allow saving a workbook whose unsupported features (charts, pivot tables, macros) would be dropped by the local engine',
      required: false,
    }),
    name: Flags.string({
      description: 'Defined name to add or remove (add/remove only)',
      required: false,
    }),
    refersTo: Flags.string({
      description: 'Sheet-qualified A1 range the name refers to, e.g. "Data!$B$2:$B$4" (add only)',
      required: false,
    }),
    dryRun: Flags.boolean({
      description: 'Preview the mutation without applying it',
      required: false,
    }),
    rawOutput: Flags.boolean({ char: 'r', description: 'Get the raw output as a JSON string', default: false, required: false }),
  };

  static args = {
    action: Args.string({
      description: 'list (default), add or remove',
      options: ['list', 'add', 'remove'],
      default: 'list',
      required: false,
    }),
  };

  async run() {
    const { args, flags } = await this.parse(WorkbookNames);
    const { file, output, inPlace, discardUnsupported, name, refersTo, dryRun, rawOutput } = flags;
    const action = args.action ?? 'list';

    if (action !== 'list') {
      if (!name) {
        throw new Error(`workbook:names ${action} requires --name.`);
      }
      if (action === 'add' && !refersTo) {
        throw new Error('workbook:names add requires --refersTo, e.g. --refersTo "Data!$B$2:$B$4".');
      }
    }

    // list is read-only: the mutation-destination requirement does not apply to it
    const target = await resolveWorkbookTarget(
      { workbook: file, output, inPlace, discardUnsupported, dryRun: dryRun || action === 'list' },
      this.id ?? 'workbook:names'
    );
    if (!target) {
      throw new Error('workbook:names works on local XLSX files; --file is required.');
    }

    if (action === 'list') {
      const receipt = { operation: this.id, names: target.workbook.listDefinedNames() };
      if (rawOutput) {
        this.log(JSON.stringify(receipt, null, 2));
        return receipt;
      }
      if (receipt.names.length === 0) {
        this.log('No defined names in this workbook.');
        return receipt;
      }
      for (const entry of receipt.names) {
        this.log(`${entry.name}\t${entry.ranges.join(', ')}`);
      }
      return receipt;
    }

    if (action === 'add') {
      const added = target.workbook.addDefinedName({ name: name ?? '', refersTo: refersTo ?? '', dryRun });
      const saved = await saveWorkbookTarget(target, { workbook: file, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...added, operation: this.id, saved };

      if (rawOutput) {
        this.log(JSON.stringify(receipt, null, 2));
      } else if (dryRun) {
        this.log(`Dry run: add defined name "${receipt.name}" = ${receipt.refersTo} in (${file})`);
      } else {
        this.log(`Added defined name "${receipt.name}" = ${receipt.refersTo} -> ${saved?.savedPath}`);
      }
      return receipt;
    }

    const removed = target.workbook.removeDefinedName({ name: name ?? '', dryRun });
    const saved = await saveWorkbookTarget(target, { workbook: file, output, inPlace, discardUnsupported, dryRun });
    const receipt = { ...removed, operation: this.id, saved };

    if (rawOutput) {
      this.log(JSON.stringify(receipt, null, 2));
    } else if (dryRun) {
      this.log(`Dry run: remove defined name "${receipt.name}" in (${file})`);
    } else {
      this.log(`Removed defined name "${receipt.name}" (was ${receipt.removedRanges.join(', ')}) -> ${saved?.savedPath}`);
    }
    return receipt;
  }
}
