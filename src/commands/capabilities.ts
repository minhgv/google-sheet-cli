import { Command } from '@oclif/core';
import { buildCapabilityDocument, CapabilityDocument } from '../lib/capabilities';

/**
 * Credential-free meta command: prints the machine-readable capability document.
 *
 * Deliberately extends plain Command rather than the shared base class - the base class
 * authenticates in init(), and the whole point of this surface is that an agent can learn what
 * the CLI can do before any credential exists. There is also deliberately no human mode and no
 * flags to parse: the output is always the JSON document, so stdout stays a stable, parseable
 * contract (use `| jq`, not a flag, to shape it).
 */
export default class Capabilities extends Command {
  static description =
    'Print the machine-readable capability document: per backend (google-sheets, local-xlsx) the supported operations, input and output forms, formula write-vs-recalculation semantics, destructive guards, mutation limits and local fidelity exclusions. Credential-free.';

  static examples = [
    `$ gsheet capabilities
{
  "schemaVersion": 1,
  "supportModel": "...",
  "visibility": [...],
  "backends": { "google-sheets": {...}, "local-xlsx": {...} }
}
`,
    `$ gsheet capabilities | jq '.backends["local-xlsx"].formulaSemantics'
`,
  ];

  async run(): Promise<CapabilityDocument> {
    // no flags on purpose; parse() still runs so oclif knows the argv was consumed
    await this.parse(Capabilities);
    const document = buildCapabilityDocument();
    this.log(JSON.stringify(document, null, 2));
    return document;
  }
}
