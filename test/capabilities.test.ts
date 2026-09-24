import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { CAPABILITY_SCHEMA_VERSION, buildCapabilityDocument, CapabilityDocument } from '../src/lib/capabilities';

/**
 * The capability contract and its credential-free surface.
 *
 * `capabilities` must be printable before any authentication exists - it is the document an
 * agent reads first - so these cases run with no factory stub, no credentials and no network.
 * The pins below are the AC-04 contract: versioned schema, both backends identified, the XLSX
 * formula-write vs recalculation distinction, and the drive.file visibility boundary.
 */
describe('capabilities', () => {
  const doc = (): CapabilityDocument => buildCapabilityDocument();

  describe('document', () => {
    it('carries its own schema version, independent of the package version', () => {
      expect(doc().schemaVersion).to.equal(CAPABILITY_SCHEMA_VERSION);
      expect(doc().schemaVersion).to.be.a('number');
      expect(CAPABILITY_SCHEMA_VERSION).to.equal(1);
    });

    it('lists both backends', () => {
      expect(Object.keys(doc().backends).sort()).to.eql(['google-sheets', 'local-xlsx']);
      expect(doc().backends['google-sheets'].id).to.equal('google-sheets');
      expect(doc().backends['local-xlsx'].id).to.equal('local-xlsx');
    });

    it('names real operations for each backend', () => {
      const { backends } = doc();
      expect(backends['google-sheets'].operations).to.contain('data:get');
      expect(backends['google-sheets'].operations).to.contain('data:clear');
      expect(backends['local-xlsx'].operations).to.contain('workbook:write');
      expect(backends['local-xlsx'].operations).to.contain('workbook:read');
      expect(backends['local-xlsx'].operations).to.contain('workbook:inspect');
      // report:run spans both backends
      expect(backends['google-sheets'].operations).to.contain('report:run');
      expect(backends['local-xlsx'].operations).to.contain('report:run');
      // T-09: scope-honest discovery and CSV interchange shipped
      expect(backends['google-sheets'].operations).to.contain('spreadsheet:list');
      expect(backends['google-sheets'].operations).to.contain('data:export-csv');
      // data:export-csv spans both backends, like report:run
      expect(backends['local-xlsx'].operations).to.contain('data:export-csv');
    });

    it('lists each operation at most once per backend', () => {
      for (const id of ['google-sheets', 'local-xlsx'] as const) {
        const ops = doc().backends[id].operations;
        expect(new Set(ops).size, id).to.equal(ops.length);
      }
    });

    it('distinguishes formula write from recalculation on both backends', () => {
      const { backends } = doc();
      const sheets = backends['google-sheets'].formulaSemantics;
      const xlsx = backends['local-xlsx'].formulaSemantics;

      // Sheets: formulas are written, then evaluated server-side
      expect(sheets.write).to.contain('USER_ENTERED');
      expect(sheets.recalculation).to.contain('recalculates formulas server-side');
      // XLSX: formulas are written but never evaluated locally - cached results only
      expect(xlsx.write).to.contain('formula cells');
      expect(xlsx.recalculation).to.contain('never recalculated');
      expect(xlsx.recalculation).to.contain('cached');
    });

    it('discloses the drive.file visibility boundary', () => {
      expect(doc().visibility.join(' ')).to.contain('drive.file');
      expect(doc().visibility.join(' ')).to.contain('by ID');
    });

    it('discloses that discovery names its command and never auto-selects a target', () => {
      const visibility = doc().visibility.join(' ');
      expect(visibility).to.contain('spreadsheet:list');
      expect(visibility).to.contain('never selects a spreadsheet by itself');
    });

    it('documents the CSV output form and retracts the old no-CSV claim', () => {
      const sheets = doc().backends['google-sheets'].outputForms.join(' ');
      expect(sheets).to.contain('data:export-csv');
      expect(sheets).to.contain('RFC 4180');
      expect(sheets).to.contain('--injection');
      expect(sheets).to.not.contain('No command currently offers CSV');
      const xlsx = doc().backends['local-xlsx'].outputForms.join(' ');
      expect(xlsx).to.contain('data:export-csv');
      expect(xlsx).to.contain('never recalculated');
    });

    it('documents the structured mutation-outcome envelope', () => {
      const forms = doc().backends['google-sheets'].outputForms.join(' ');
      expect(forms).to.contain('"mutation?"');
      expect(forms).to.contain('error.mutation');
      expect(forms).to.contain('acknowledged, rejected, unknown');
      expect(forms).to.contain('never-blind-replay');
      expect(forms).to.contain('never cell contents');
    });

    it('documents the --redacted diagnostics mode', () => {
      const forms = doc().backends['google-sheets'].outputForms.join(' ');
      expect(forms).to.contain('--redacted');
      expect(forms).to.contain('mutation outcome states are kept');
      expect(forms).to.contain('stay data-bearing');
    });

    it('states the destructive guards', () => {
      const { backends } = doc();
      const sheets = backends['google-sheets'].destructiveGuards.join(' ');
      expect(sheets).to.contain('--dryRun');
      expect(sheets).to.contain('bounded on both axes');
      expect(sheets).to.contain('--overwriteFormulas');

      const xlsx = backends['local-xlsx'].destructiveGuards.join(' ');
      expect(xlsx).to.contain('--overwriteFormulas');
      expect(xlsx).to.contain('--discardUnsupported');
      expect(xlsx).to.contain('--overwrite');
    });

    it('states the mutation limits honestly', () => {
      const { backends } = doc();
      for (const id of ['google-sheets', 'local-xlsx'] as const) {
        expect(backends[id].mutationLimits.join(' '), id).to.contain('Single-writer');
      }
      const xlsxLimits = backends['local-xlsx'].mutationLimits.join(' ');
      expect(xlsxLimits).to.contain('never recomputed');
      const sheets = backends['google-sheets'].mutationLimits.join(' ');
      expect(sheets).to.contain('not transactional');
      expect(sheets).to.contain('unknown outcome');
      expect(sheets).to.contain('never blind-replayed');
    });

    it('lists the local fidelity exclusions and none for Sheets', () => {
      const xlsx = doc().backends['local-xlsx'].fidelityExclusions.join(' ');
      expect(xlsx).to.contain('VBA macros');
      expect(xlsx).to.contain('charts');
      expect(xlsx).to.contain('Pivot tables');
      expect(doc().backends['google-sheets'].fidelityExclusions).to.eql([]);
    });

    it('separates static support from runtime authorization', () => {
      expect(doc().supportModel).to.contain('Authorization is only checked when a command runs');
    });

    it('is deterministic across calls', () => {
      expect(JSON.stringify(doc())).to.equal(JSON.stringify(buildCapabilityDocument()));
    });

    it('is deeply frozen so consumers cannot corrupt the shared literal', () => {
      const d = doc();
      expect(Object.isFrozen(d)).to.be.true;
      expect(Object.isFrozen(d.visibility)).to.be.true;
      expect(Object.isFrozen(d.backends)).to.be.true;
      expect(Object.isFrozen(d.backends['google-sheets'])).to.be.true;
      expect(Object.isFrozen(d.backends['google-sheets'].operations)).to.be.true;
      expect(Object.isFrozen(d.backends['local-xlsx'].formulaSemantics)).to.be.true;
    });
  });

  describe('command', () => {
    it('prints the document as JSON to stdout without any credentials', async () => {
      const { error, result, stdout } = await runCommand(['capabilities']);
      if (error) throw error;

      const printed = JSON.parse(stdout);
      expect(printed).to.eql(doc());
      expect(printed.schemaVersion).to.equal(CAPABILITY_SCHEMA_VERSION);
      expect(result).to.eql(doc());
    });

    it('prints byte-identical output on repeated runs', async () => {
      const first = await runCommand(['capabilities']);
      const second = await runCommand(['capabilities']);
      if (first.error) throw first.error;
      if (second.error) throw second.error;
      expect(second.stdout).to.equal(first.stdout);
    });

    it('answers --help without credentials', async () => {
      const { error, stdout } = await runCommand(['capabilities', '--help']);
      if (error) throw error;
      expect(stdout).to.contain('capabilities');
    });
  });
});
