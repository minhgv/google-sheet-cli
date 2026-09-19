import * as fs from 'fs';
import { Command, Flags, Args } from '@oclif/core';
import { parseRangesFlag, validate2DMatrix } from '../../lib/cli-input';
import { parseInput, parseJsonSafely, readInput } from '../../lib/report/input';
import { ReportDocument } from '../../lib/report/types';
import { XlsxWorkbook } from '../../lib/xlsx';

export default class WorkbookWrite extends Command {
  static description = 'Write tabular data or a ReportDocument into a local XLSX workbook with atomic persistence';

  static examples = [
    `$ gsheet workbook:write --output=new.xlsx '[["Name", "Amount"], ["Alice", 100], ["Bob", 200]]'
`,
    `$ gsheet workbook:write --file=template.xlsx --output=filled.xlsx --input=data.csv --worksheetTitle="Summary"
`,
    `$ gsheet workbook:write --file=existing.xlsx --inPlace --input=rows.json --startCell="B5" --overwriteFormulas
`,
  ];

  static flags = {
    help: Flags.help({ char: 'h' }),
    file: Flags.string({
      char: 'f',
      description: 'Path to existing template XLSX file to load and modify',
      required: false,
    }),
    output: Flags.string({
      char: 'o',
      description: 'Destination path for the output XLSX file',
      required: false,
    }),
    inPlace: Flags.boolean({
      description: 'Modify the --file workbook in place (creates automatic .bak backup)',
      default: false,
      required: false,
    }),
    input: Flags.string({
      char: 'i',
      description: 'Path to input data file (JSON or CSV) or "-" for stdin',
      required: false,
    }),
    inputFormat: Flags.string({
      description: 'Format of input file ("json" or "csv")',
      options: ['json', 'csv'],
      required: false,
    }),
    worksheetTitle: Flags.string({
      char: 't',
      description: 'Target worksheet name (default: "Sheet1")',
      default: 'Sheet1',
      required: false,
    }),
    startCell: Flags.string({
      description: 'Top-left cell address to place data (e.g. "A1", "B2")',
      default: 'A1',
      required: false,
    }),
    dryRun: Flags.boolean({
      description: 'Preview changes without modifying or saving files',
      default: false,
      required: false,
    }),
    overwrite: Flags.boolean({
      description: 'Allow overwriting existing destination output file',
      default: false,
      required: false,
    }),
    overwriteFormulas: Flags.boolean({
      description: 'Allow overwriting existing formula cells in the workbook',
      default: false,
      required: false,
    }),
    rawOutput: Flags.boolean({
      char: 'r',
      description: 'Get the raw output as a JSON string',
      default: false,
      required: false,
    }),
  };

  static args = {
    data: Args.string({
      name: 'data',
      description: 'Data as a JSON string (nested 2D array or ReportDocument object)',
      required: false,
      // Same contract as optionalData: stdin goes through `--input -`, never auto-fills here.
      ignoreStdin: true,
    }),
  };

  async run() {
    const {
      args: { data },
      flags: {
        file,
        output,
        inPlace,
        input,
        inputFormat,
        worksheetTitle,
        startCell,
        dryRun,
        overwrite,
        overwriteFormulas,
        rawOutput,
      },
    } = await this.parse(WorkbookWrite);

    // Validate destination
    if (inPlace && !file) {
      throw new Error('The --inPlace flag requires --file to specify the workbook to modify in place');
    }
    if (!output && !inPlace) {
      throw new Error('Must specify either --output destination path or --inPlace flag');
    }

    const destination = output || (inPlace ? file : undefined);

    if (output && file && isSamePathOrFile(output, file) && !inPlace) {
      throw new Error(
        `Destination output "${output}" collides with source template file "${file}". To modify in place, use the --inPlace flag.`
      );
    }
    if (output && input && input !== '-' && isSamePathOrFile(output, input) && !inPlace) {
      throw new Error(
        `Destination output "${output}" collides with input data file "${input}". Specify a different output path.`
      );
    }

    // Validate mutually exclusive input sources
    if (data !== undefined && data !== '' && input !== undefined && input !== '') {
      throw new Error('Mutually exclusive input sources: specify either positional data argument or --input flag, not both');
    }
    if ((data === undefined || data === '') && (input === undefined || input === '')) {
      throw new Error('No data provided. Specify either positional data argument or --input flag');
    }
    // Resolve input payload
    let rawInput: unknown;
    if (data !== undefined && data !== '') {
      let parsed: unknown = data;
      if (typeof parsed === 'string') {
        try {
          parsed = parseJsonSafely(parsed);
          while (typeof parsed === 'string' && (parsed.trim().startsWith('[') || parsed.trim().startsWith('{'))) {
            parsed = parseJsonSafely(parsed);
          }
        } catch (err: unknown) {
          const msg = err instanceof Error ? err.message : String(err);
          throw new Error(`"data" input has to be valid JSON (${msg})`);
        }
      }
      rawInput = parsed;
    } else if (input !== undefined && input !== '') {
      if (input === '-') {
        // Read stdin
        const rawContent = await readStdin();
        const detectedFormat = inputFormat || (rawContent.trim().startsWith('{') || rawContent.trim().startsWith('[') ? 'json' : 'csv');
        rawInput = parseInput(rawContent, detectedFormat as 'json' | 'csv', { matrix: true });
      } else {
        const stat = await fs.promises.stat(input).catch((err) => {
          throw new Error(`Cannot read input file "${input}": ${err.message}`);
        });
        if (stat.size > 20 * 1024 * 1024) {
          throw new Error(`Input file size (${stat.size} bytes) exceeds limit of 20MB`);
        }
        const content = await fs.promises.readFile(input, 'utf8');
        let detectedFormat = inputFormat;
        if (!detectedFormat) {
          if (input.toLowerCase().endsWith('.json')) {
            detectedFormat = 'json';
          } else if (input.toLowerCase().endsWith('.csv')) {
            detectedFormat = 'csv';
          } else {
            detectedFormat = content.trim().startsWith('{') || content.trim().startsWith('[') ? 'json' : 'csv';
          }
        }
        rawInput = parseInput(content, detectedFormat as 'json' | 'csv', { matrix: true });
      }
    }

    const document = this.toReportDocument(rawInput, worksheetTitle, startCell);

    // Load or create workbook
    const workbook = file ? await XlsxWorkbook.load(file) : XlsxWorkbook.create();

    // Apply document
    const applyResult = workbook.apply(document, {
      dryRun,
      overwriteFormulas,
      defaultStartCell: startCell,
    });

    let saveResult = undefined;
    if (!dryRun && destination) {
      saveResult = await workbook.save(destination, {
        overwrite,
        inPlace,
      });
    }

    const result = {
      operation: this.id,
      destination: destination || 'dry-run',
      dryRun,
      appliedChanges: applyResult.appliedChanges,
      affectedSheets: applyResult.diff.affectedSheets,
      totalConflicts: applyResult.diff.totalConflicts,
      saved: saveResult,
      warnings: applyResult.warnings,
    };

    if (rawOutput) {
      this.log(JSON.stringify(result, null, 2));
    } else if (dryRun) {
      this.log(`Workbook write preview (dry run):`);
      this.log(`  - Target destination: ${destination}`);
      this.log(`  - Changes to apply: ${applyResult.appliedChanges} cells across ${applyResult.diff.affectedSheets.join(', ')}`);
      this.log(`  - Conflicts detected: ${applyResult.diff.totalConflicts}`);
    } else {
      this.log(`Successfully wrote ${applyResult.appliedChanges} cell(s) to "${destination}":`);
      if (saveResult) {
        this.log(`  - File size: ${saveResult.bytesWritten} bytes`);
        this.log(`  - SHA-256: ${saveResult.sha256Hash}`);
        if (saveResult.backupCreated) {
          this.log(`  - Backup created: ${saveResult.backupCreated}`);
        }
      }
    }

    return result;
  }

  private toReportDocument(rawInput: unknown, defaultSheetTitle: string, defaultStartCell: string): ReportDocument {
    if (rawInput && typeof rawInput === 'object') {
      if ('sheets' in rawInput) {
        const rawSheets = rawInput.sheets;
        if (Array.isArray(rawSheets)) {
          for (let i = 0; i < rawSheets.length; i++) {
            const s = rawSheets[i];
            if (!s || typeof s !== 'object' || !('name' in s) || typeof s.name !== 'string' || !('rows' in s) || !Array.isArray(s.rows)) {
              throw new Error(`ReportDocument sheet at index ${i} is invalid (must have "name" string and "rows" 2D array)`);
            }
          }
          const provenance = 'provenance' in rawInput && rawInput.provenance && typeof rawInput.provenance === 'object'
            ? (rawInput.provenance as unknown as ReportDocument['provenance'])
            : {
                templateId: 'cli-direct-write',
                templateVersion: 1,
                sourceHash: 'direct-write',
              };
          return {
            sheets: rawSheets as unknown as ReportDocument['sheets'],
            provenance,
          };
        }
      }
    }

    if (Array.isArray(rawInput)) {
      const matrix = validate2DMatrix(rawInput);
      return {
        sheets: [
          {
            name: defaultSheetTitle || 'Sheet1',
            rows: matrix,
            startCell: defaultStartCell || 'A1',
          },
        ],
        provenance: {
          templateId: 'cli-direct-write',
          templateVersion: 1,
          sourceHash: 'direct-write',
        },
      };
    }

    throw new Error(
      'Input data must be either a 2D array [[...]] or a valid ReportDocument object ({ sheets: [...], provenance: {...} })'
    );
  }
}

function isSamePathOrFile(pathA?: string, pathB?: string): boolean {
  if (!pathA || !pathB) return false;
  if (pathA === pathB) return true;
  try {
    const realA = fs.realpathSync(pathA);
    const realB = fs.realpathSync(pathB);
    if (realA === realB) return true;
    const statA = fs.statSync(pathA);
    const statB = fs.statSync(pathB);
    if (statA.ino && statB.ino && statA.dev === statB.dev && statA.ino === statB.ino) {
      return true;
    }
  } catch {
    // One or both paths might not exist on disk
  }
  return false;
}

async function readStdin(maxBytes = 20 * 1024 * 1024): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const chunks: Buffer[] = [];
    let totalBytes = 0;

    const onData = (chunk: Buffer) => {
      totalBytes += chunk.length;
      if (totalBytes > maxBytes) {
        process.stdin.removeListener('data', onData);
        process.stdin.removeListener('end', onEnd);
        process.stdin.removeListener('error', onError);
        reject(new Error(`Stdin input exceeded maximum limit of ${maxBytes} bytes`));
        return;
      }
      chunks.push(chunk);
    };

    const onEnd = () => {
      resolve(Buffer.concat(chunks).toString('utf8'));
    };

    const onError = (err: Error) => {
      reject(new Error(`Failed to read from stdin: ${err.message}`));
    };

    if (process.stdin.isTTY) {
      resolve('');
      return;
    }

    process.stdin.on('data', onData);
    process.stdin.on('end', onEnd);
    process.stdin.on('error', onError);
    process.stdin.resume();
  });
}
