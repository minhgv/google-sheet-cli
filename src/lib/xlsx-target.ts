import * as fs from 'fs';
import * as path from 'path';
import { GSheetError, GSheetErrorCode } from './cli-errors';
import { XlsxWorkbook } from './xlsx';
import { XlsxSaveResult } from './xlsx-types';

/**
 * Shared plumbing for dual-backend mutation commands (`--workbook` local path).
 *
 * One source rule mirrors data:export-csv: `--workbook` cannot be combined with
 * `--spreadsheetId`; `-t/--worksheetTitle` stays legal and selects the sheet
 * inside the workbook. A mutation never writes back to the source file unless
 * `--inPlace` is passed; `--output` writes a copy. `--dryRun` performs the
 * in-memory operation and skips the save entirely.
 */
export interface WorkbookTargetFlags {
  workbook?: string;
  spreadsheetId?: string;
  worksheetTitle?: string;
  output?: string;
  inPlace?: boolean;
  discardUnsupported?: boolean;
  dryRun?: boolean;
}

export interface WorkbookTarget {
  workbook: XlsxWorkbook;
  /** sheet selector inside the workbook (-t), or undefined for the first sheet */
  worksheetTitle?: string;
}

/**
 * Validates the flag combination and loads the workbook. Returns null when the
 * command should take the Google Sheets path instead.
 */
export async function resolveWorkbookTarget(
  flags: WorkbookTargetFlags,
  commandId: string
): Promise<WorkbookTarget | null> {
  if (!flags.workbook) return null;

  if (flags.spreadsheetId) {
    throw new GSheetError(
      GSheetErrorCode.VALIDATION,
      `--workbook targets a local file and cannot be combined with --spreadsheetId. Pass exactly one target.`
    );
  }
  if (flags.inPlace && flags.output) {
    throw new GSheetError(
      GSheetErrorCode.VALIDATION,
      `--inPlace and --output are mutually exclusive for ${commandId}.`
    );
  }
  if (!flags.dryRun && !flags.inPlace && !flags.output) {
    throw new GSheetError(
      GSheetErrorCode.VALIDATION,
      `${commandId} with --workbook needs a destination: --inPlace to modify the file, --output <path> for a copy, or --dryRun to preview without saving.`
    );
  }
  if (flags.output && !flags.inPlace) {
    const [realSrc, realDst] = await Promise.all([
      fs.promises.realpath(flags.workbook).catch(() => null),
      fs.promises.realpath(flags.output).catch(() => null),
    ]);
    const same = flags.output === flags.workbook || (realSrc !== null && realSrc === realDst);
    if (same) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        `--output "${flags.output}" collides with the source workbook. To modify it in place, use --inPlace.`
      );
    }
  }

  const workbook = await XlsxWorkbook.load(flags.workbook);
  return { workbook, worksheetTitle: flags.worksheetTitle };
}

/**
 * Persists a mutated workbook according to the target flags. No-op on dryRun.
 * The unsupported-feature gate lives in XlsxWorkbook.save; --discardUnsupported
 * maps onto allowUnsupportedFeatures.
 */
export async function saveWorkbookTarget(
  target: WorkbookTarget,
  flags: WorkbookTargetFlags
): Promise<XlsxSaveResult | undefined> {
  if (flags.dryRun) return undefined;
  const destination = flags.inPlace ? target.workbook.filePath : flags.output;
  if (!destination) return undefined;
  return target.workbook.save(path.resolve(destination), {
    overwrite: Boolean(flags.output),
    inPlace: Boolean(flags.inPlace),
    backup: Boolean(flags.inPlace),
    allowUnsupportedFeatures: Boolean(flags.discardUnsupported),
  });
}
