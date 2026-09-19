import { Command, Flags } from '@oclif/core';
import { XlsxWorkbook } from '../../lib/xlsx';

export default class WorkbookInspect extends Command {
  static description = 'Inspect metadata, sheets, defined names, and capabilities of a local XLSX workbook';

  static examples = [
    `$ gsheet workbook:inspect --file=report.xlsx
`,
    `$ gsheet workbook:inspect --file=report.xlsx --rawOutput
`,
  ];

  static flags = {
    help: Flags.help({ char: 'h' }),
    file: Flags.string({
      char: 'f',
      description: 'Path to the local XLSX file to inspect',
      required: true,
    }),
    rawOutput: Flags.boolean({
      char: 'r',
      description: 'Get the raw output as a JSON string',
      default: false,
      required: false,
    }),
  };

  async run() {
    const {
      flags: { file, rawOutput },
    } = await this.parse(WorkbookInspect);

    const workbook = await XlsxWorkbook.load(file);
    const inspection = workbook.inspect();

    const result = {
      operation: this.id,
      ...inspection,
    };

    if (rawOutput) {
      this.log(JSON.stringify(result, null, 2));
    } else {
      this.log(`XLSX Workbook Inspection: "${inspection.filePath || file}"`);
      this.log(`  - File size: ${inspection.fileSize ? `${inspection.fileSize} bytes` : 'unknown'}`);
      this.log(`  - SHA-256: ${inspection.sha256Hash || 'none'}`);
      this.log(`  - Sheets (${inspection.sheetCount}):`);
      for (const sheet of inspection.sheets) {
        this.log(`    • "${sheet.name}" (ID: ${sheet.id}, Rows: ${sheet.rowCount}, Cols: ${sheet.columnCount}, Has formulas: ${sheet.hasFormulas})`);
      }
      if (inspection.definedNames && inspection.definedNames.length > 0) {
        this.log(`  - Defined Names (${inspection.definedNames.length}):`);
        for (const dn of inspection.definedNames) {
          this.log(`    • ${dn.name} => ${dn.ranges.join(', ')}`);
        }
      }
      this.log(`  - Capabilities:`);
      this.log(`    • Can read: ${inspection.capabilities.canRead}`);
      this.log(`    • Can write safely: ${inspection.capabilities.canWriteSafely}`);
      this.log(`    • General formula recalculation: ${inspection.capabilities.canRecalculateFormulas} (ExcelJS does not evaluate arbitrary formulas)`);
      if (inspection.hasUnsupportedFeatures) {
        this.log(`  ⚠️ Unsupported features detected (${inspection.unsupportedFeatures.length}):`);
        for (const uf of inspection.unsupportedFeatures) {
          this.log(`    • [${uf.type}] ${uf.description}`);
        }
      }
      if (inspection.warnings && inspection.warnings.length > 0) {
        for (const w of inspection.warnings) {
          this.log(`  ⚠️ Warning: ${w}`);
        }
      }
    }

    return result;
  }
}
