import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { ErrorIssue, GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { cellScalar, columnIndexByHeader, discoverSchema, findUniqueViolations, mapReportIssues, parseValidationSchema } from '../../lib/data-schema';
import { readInput } from '../../lib/report/input';
import { normalizeTableWithErrors } from '../../lib/report/schema';
import { TableSchema } from '../../lib/report/types';
import { formatBoundedA1Range } from '../../lib/sheet-batch';
import { table, tableFlags } from '../../lib/table';

export default class ValidateData extends Command {
  static description = `Validates worksheet data against a TableSchema (read-only)

Reuses the report validation semantics (required, type, enum, unique, min/max, ...),
reads the sheet with RAW cell values, and reports every violation at its exact sheet
row and A1 coordinate. Nothing is ever written. The run exits nonzero when the data
violates a valid schema; a malformed schema is reported as SCHEMA_INVALID instead.
With --json or --rawOutput the issues arrive inside the error envelope on stderr.`;

  static examples = [
    `$ gsheet data:validate --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --schema='{"fields":[{"name":"Name","type":"string","required":true},{"name":"Amount","type":"decimal","min":0},{"name":"Code","type":"string","unique":true}]}'`,
  ];

  static flags = {
    ...Command.flags,
    ...tableFlags(),
    spreadsheetId,
    worksheetTitle,
    schema: Flags.string({ description: 'The TableSchema as a JSON string', required: false, exclusive: ['schemaFile'] }),
    schemaFile: Flags.string({ description: 'Path to a TableSchema JSON file, or "-" for stdin', required: false }),
    minRow: Flags.integer({ description: 'The first row of the validated range (holds the header row)', default: 1, required: false }),
    minCol: Flags.integer({ description: 'The first column of the validated range', default: 1, required: false }),
    maxRow: Flags.integer({ description: 'The last row of the validated range', default: 100, required: false }),
    maxCol: Flags.integer({ description: 'The last column of the validated range', default: 26, required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, rawOutput, json, worksheetTitle, schema, schemaFile, minRow, minCol, maxRow, maxCol, ...tableOptions },
    } = await this.parse(ValidateData);

    if (!schema && !schemaFile) {
      throw new GSheetError(GSheetErrorCode.VALIDATION, 'No schema provided. Specify either --schema or --schemaFile.');
    }
    if (minRow < 1 || minCol < 1) throw new GSheetError(GSheetErrorCode.VALIDATION, '--minRow and --minCol must be at least 1');
    if (maxRow < minRow) throw new GSheetError(GSheetErrorCode.VALIDATION, `--maxRow (${maxRow}) must not be smaller than --minRow (${minRow})`);
    if (maxCol < minCol) throw new GSheetError(GSheetErrorCode.VALIDATION, `--maxCol (${maxCol}) must not be smaller than --minCol (${minCol})`);

    // Fail fast on a malformed schema, before any request leaves the process. readInput and
    // parseInput throw plain Errors for local input problems (unreadable file, oversized
    // input, bad JSON); surfaced as VALIDATION so a caller can tell its own input apart from
    // a CLI defect instead of seeing INTERNAL. parseValidationSchema below throws the more
    // specific SCHEMA_INVALID for a structurally wrong schema.
    let schemaInput: unknown;
    try {
      schemaInput = await readInput({ data: schema, file: schemaFile, format: 'json' });
    } catch (err) {
      if (err instanceof GSheetError) throw err;
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        err instanceof Error && err.message ? err.message : 'The validation schema could not be read.',
        { cause: err }
      );
    }
    const { schema: tableSchema, uniqueFields } = parseValidationSchema(schemaInput);

    this.start('Validating data');
    const metadata = await this.gsheet.getWorksheetMetadata(
      { worksheetTitle, range: formatBoundedA1Range(undefined, minCol, minRow, maxCol, maxRow) },
      spreadsheetId
    );
    const discovery = discoverSchema(metadata);
    const byHeader = columnIndexByHeader(discovery);
    const issues: ErrorIssue[] = [];
    const warnings: string[] = [];

    // The sheet header row must carry every schema field exactly once - otherwise
    // normalization would either mis-map columns or silently validate the header row as data.
    for (const field of tableSchema.fields) {
      if (!byHeader.has(field.name)) {
        issues.push({ field: field.name, code: 'MISSING_HEADER', message: `Schema field "${field.name}" not found in the sheet header row` });
      }
    }
    for (const column of discovery.columns) {
      if (column.emptyHeader) warnings.push(`Empty header at ${column.headerA1}`);
      if (column.duplicate && byHeader.has(column.header)) {
        issues.push({ field: column.header, a1: column.headerA1, code: 'DUPLICATE_HEADER', message: `Duplicate header "${column.header}" at ${column.headerA1}` });
      }
    }

    const gridRows = metadata.gridData[0]?.rowData ?? [];
    const rowsChecked = Math.max(gridRows.length - 1, 0);

    if (issues.length === 0) {
      const headerWidth = discovery.columns.length;
      const dataRows = gridRows.slice(1).map((row) => Array.from({ length: headerWidth }, (_, j) => cellScalar(row.values?.[j])));
      const matrix: (string | number | boolean | null)[][] = [discovery.columns.map((column) => column.header), ...dataRows];
      // One pass reports row-level and uniqueness problems together. matrix rows are always
      // arrays, so no row is ever skipped as INVALID_ROW_SHAPE and normalizedRows[k] stays
      // index-aligned with dataRows[k] - the sheet-row math in both mappers holds even when
      // some fields of a row failed. Failed fields are absent from their row record, so the
      // null/undefined guard keeps them out of uniqueness keys.
      const { rows: normalizedRows, errors: rowErrors } = normalizeTableWithErrors(matrix, tableSchema, worksheetTitle);
      issues.push(...mapReportIssues(rowErrors, discovery));
      issues.push(...findUniqueViolations(normalizedRows, uniqueFields, discovery));
    }

    if (issues.length > 0) {
      if (!rawOutput && !json) {
        table(
          issues.map((issue) => ({
            row: String(issue.row ?? '-'),
            a1: issue.a1 ?? '-',
            field: issue.field ?? '-',
            code: issue.code ?? '-',
            message: issue.message,
          })),
          { row: {}, a1: {}, field: {}, code: {}, message: {} },
          tableOptions
        );
      }
      throw new GSheetError(GSheetErrorCode.DATA_INVALID, `Data validation failed with ${issues.length} issue(s)`, {
        details: { issues },
      });
    }

    const result = {
      operation: this.id,
      spreadsheetId: discovery.spreadsheetId,
      worksheetTitle: discovery.worksheetTitle,
      range: discovery.sample.range,
      rowsChecked,
      columnsChecked: tableSchema.fields.length,
      valid: true,
      issues: [] as ErrorIssue[],
      warnings,
    };
    if (rawOutput) {
      this.logRaw('', result);
      return result;
    }
    this.stop();
    this.log(
      `OK: ${result.rowsChecked} row(s) validated against ${tableSchema.fields.length} field(s) in ${discovery.sample.range}`
    );
    for (const warning of warnings) this.log(`Warning: ${warning}`);
    return result;
  }
}
