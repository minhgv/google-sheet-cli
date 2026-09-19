import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { readInput } from '../../lib/report/input';
import { FormatSpec } from '../../lib/sheet-format';

const STYLE_FLAG_NAMES = [
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'fontSize',
  'fontFamily',
  'textColor',
  'backgroundColor',
  'horizontalAlignment',
  'verticalAlignment',
  'wrapStrategy',
  'numberFormat',
  'numberFormatType',
  'borders',
] as const;

export default class FormatCells extends Command {
  static description =
    'Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. ' +
    'Only formatting is touched - cell values and formulas are never overwritten.';

  static examples = [
    `$ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --bold --backgroundColor="#1a73e8" --textColor="#ffffff"
`,
    `$ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A2:J50 --numberFormat="#,##0.00" --borders=all
`,
    `$ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J50 --clear
`,
    `$ gsheet format:cells --spreadsheetId=<id> --worksheetTitle=Report --input=style.json --dryRun
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    range: Flags.string({ description: 'The A1 range to format (e.g. "A1:J1"); required unless --input carries "ranges"', required: false }),
    bold: Flags.boolean({ description: 'Bold text', required: false }),
    italic: Flags.boolean({ description: 'Italic text', required: false }),
    underline: Flags.boolean({ description: 'Underline text', required: false }),
    strikethrough: Flags.boolean({ description: 'Strikethrough text', required: false }),
    fontSize: Flags.integer({ description: 'Font size in points', required: false }),
    fontFamily: Flags.string({ description: 'Font family name (e.g. "Roboto")', required: false }),
    textColor: Flags.string({ description: 'Text color as #RRGGBB', required: false }),
    backgroundColor: Flags.string({ description: 'Cell background color as #RRGGBB', required: false }),
    horizontalAlignment: Flags.string({
      description: 'Horizontal alignment',
      options: ['LEFT', 'CENTER', 'RIGHT'],
      required: false,
    }),
    verticalAlignment: Flags.string({
      description: 'Vertical alignment',
      options: ['TOP', 'MIDDLE', 'BOTTOM'],
      required: false,
    }),
    wrapStrategy: Flags.string({
      description: 'Text wrap strategy',
      options: ['OVERFLOW_CELL', 'CLIP', 'WRAP'],
      required: false,
    }),
    numberFormat: Flags.string({ description: 'Number format pattern (e.g. "#,##0.00", "0.0%", "YYYY-MM-DD")', required: false }),
    numberFormatType: Flags.string({
      description: 'Explicit number format type (inferred from the pattern when omitted)',
      options: ['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC'],
      required: false,
      dependsOn: ['numberFormat'],
    }),
    borders: Flags.string({
      description: 'Border sides: comma list of top,bottom,left,right,innerHorizontal,innerVertical or "all"/"inner"',
      required: false,
    }),
    borderStyle: Flags.string({
      description: 'Border line style (only used with --borders)',
      options: ['DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'],
      default: 'SOLID',
      required: false,
    }),
    borderColor: Flags.string({ description: 'Border color as #RRGGBB (only used with --borders)', default: '#000000', required: false }),
    clear: Flags.boolean({ description: 'Reset all formatting on the range (cannot be combined with style flags)', required: false }),
    input: Flags.string({ char: 'i', description: 'Path to a JSON format spec file (or "-" for stdin)', required: false }),
    dryRun: Flags.boolean({ description: 'Preview the batchUpdate requests without applying them', required: false }),
  };

  async run() {
    const { flags } = await this.parse(FormatCells);
    const {
      spreadsheetId,
      worksheetTitle,
      range,
      clear,
      input,
      dryRun,
      rawOutput,
      borders,
      borderStyle,
      borderColor,
      ...styleFlags
    } = flags;

    let spec: FormatSpec;

    if (input !== undefined && input !== '') {
      const styleFlagsSet = clear || STYLE_FLAG_NAMES.some((name) => (flags as Record<string, unknown>)[name] !== undefined);
      if (styleFlagsSet) {
        throw new Error('The --input spec cannot be combined with style flags or --clear');
      }
      const parsed = (await readInput({ file: input, format: 'json' })) as FormatSpec;
      if (!parsed || typeof parsed !== 'object') {
        throw new Error('Format spec input must be a JSON object');
      }
      spec = { ...parsed, worksheetTitle: parsed.worksheetTitle || worksheetTitle };
      if (range && !(Array.isArray(parsed.ranges) && parsed.ranges.length > 0)) {
        spec.ranges = [range];
      }
    } else {
      if (!range) {
        throw new Error('The --range flag is required when no --input spec is provided');
      }
      if (clear && STYLE_FLAG_NAMES.some((name) => (flags as Record<string, unknown>)[name] !== undefined)) {
        throw new Error('The --clear flag cannot be combined with style flags');
      }
      const style: Record<string, unknown> = {};
      for (const name of STYLE_FLAG_NAMES) {
        if (name === 'borders') continue;
        const value = (styleFlags as Record<string, unknown>)[name];
        if (value !== undefined) style[name] = value;
      }
      if (borders) {
        const sides = borders
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean);
        style.borders =
          sides.length === 1 && (sides[0] === 'all' || sides[0] === 'inner')
            ? { sides: sides[0], style: borderStyle, color: borderColor }
            : { sides, style: borderStyle, color: borderColor };
      }
      spec = {
        ranges: [range],
        worksheetTitle,
        ...(clear ? { clear: true } : { style: style as FormatSpec['style'] }),
      };
    }

    this.start(dryRun ? `Previewing format (${spec.ranges.length} range${spec.ranges.length > 1 ? 's' : ''})` : 'Applying format');
    const result = await this.gsheet.formatCells(spec, Boolean(dryRun), spreadsheetId);
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: ${result.requestCount} request(s) over ${result.ranges.join(', ')}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`Formatted ${result.ranges.join(', ')} (${result.fields.join(', ')})`);
    }

    return result;
  }
}
