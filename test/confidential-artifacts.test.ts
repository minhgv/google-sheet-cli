import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { expect } from 'chai';
import { saveBufferAtomic } from '../src/lib/xlsx-file';

const IS_POSIX = process.platform !== 'win32';

interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  scope: string;
  token_type: string;
  expiry_date: number;
}

interface OAuthModule {
  saveTokens(tokens: OAuthTokens): void;
  loadTokens(): OAuthTokens | null;
  TOKEN_PATH: string;
}

/** Fresh fixture per call so no test ever mutates a shared object. */
const makeTokens = (): OAuthTokens => ({
  access_token: 'ya29.test-access-token',
  refresh_token: '1//test-refresh-token',
  scope: 'https://www.googleapis.com/auth/spreadsheets',
  token_type: 'Bearer',
  expiry_date: Date.now() + 60 * 60 * 1000,
});

/** Any atomic-save leftovers in dir; sorted so deep.equal is deterministic. */
const tempFilesIn = (dir: string): string[] =>
  fs.readdirSync(dir).filter((name) => name.endsWith('.tmp')).sort();

describe('Confidential artifacts (token persistence + atomic export)', () => {
  // oauth.ts resolves CONFIG_DIR from os.homedir() at module load. The module
  // under test is (re)loaded with HOME pointed at a private tmpdir so the
  // suite never touches the real ~/.config/google-sheet-cli. Any pre-existing
  // cache entry is preserved and restored afterwards so other test files in a
  // combined run keep the instance they loaded.
  const OAUTH_MODULE = require.resolve('../src/lib/oauth');
  let realHome: string | undefined;
  let realUserProfile: string | undefined;
  let cachedModule: NodeJS.Module | undefined;
  let tmpHome: string;
  let oauth: OAuthModule;
  let configDir: string;
  let tokenPath: string;

  let realConfigDir: string | undefined;

  before(() => {
    realHome = process.env.HOME;
    realUserProfile = process.env.USERPROFILE;
    realConfigDir = process.env.GSHEET_CONFIG_DIR;
    // oauth.ts resolves CONFIG_DIR at module load; GSHEET_CONFIG_DIR overrides
    // it (os.homedir() ignores process.env.HOME on POSIX, so env HOME alone
    // cannot redirect). Re-require the module under the redirected dir.
    cachedModule = require.cache[OAUTH_MODULE];
    tmpHome = fs.mkdtempSync(path.join(os.tmpdir(), 'gsheet-confidential-'));
    process.env.HOME = tmpHome;
    process.env.USERPROFILE = tmpHome;
    process.env.GSHEET_CONFIG_DIR = path.join(tmpHome, '.config', 'google-sheet-cli');
    delete require.cache[OAUTH_MODULE];
    oauth = require('../src/lib/oauth');
    tokenPath = oauth.TOKEN_PATH;
    configDir = path.dirname(tokenPath);
    expect(configDir.startsWith(tmpHome)).to.equal(true);
  });

  after(() => {
    if (realConfigDir === undefined) {
      delete process.env.GSHEET_CONFIG_DIR;
    } else {
      process.env.GSHEET_CONFIG_DIR = realConfigDir;
    }
    if (realHome === undefined) {
      delete process.env.HOME;
    } else {
      process.env.HOME = realHome;
    }
    if (realUserProfile === undefined) {
      delete process.env.USERPROFILE;
    } else {
      process.env.USERPROFILE = realUserProfile;
    }
    if (cachedModule) {
      require.cache[OAUTH_MODULE] = cachedModule;
    } else {
      delete require.cache[OAUTH_MODULE];
    }
    fs.rmSync(tmpHome, { recursive: true, force: true });
  });


  beforeEach(() => {
    // Every case starts with no config dir, so creation/tightening behavior
    // is exercised from scratch regardless of test order.
    fs.rmSync(configDir, { recursive: true, force: true });
  });

  describe('token persistence', () => {
    it('saveTokens creates a 0700 dir and 0600 token file, leaves no temp files, and loadTokens reads it back', () => {
      const tokens = makeTokens();
      oauth.saveTokens(tokens);

      if (IS_POSIX) {
        expect(fs.statSync(configDir).mode & 0o777).to.equal(0o700);
        expect(fs.statSync(tokenPath).mode & 0o777).to.equal(0o600);
      }
      expect(tempFilesIn(configDir)).to.deep.equal([]);
      expect(oauth.loadTokens()).to.deep.equal(tokens);
    });

    it('saveTokens tightens an existing permissive config dir to 0700', function () {
      if (!IS_POSIX) this.skip();
      fs.mkdirSync(configDir, { recursive: true });
      fs.chmodSync(configDir, 0o755);

      oauth.saveTokens(makeTokens());

      expect(fs.statSync(configDir).mode & 0o777).to.equal(0o700);
      expect(fs.existsSync(tokenPath)).to.equal(true);
    });

    it('saveTokens refuses a symlinked token path without touching the target', function () {
      if (!IS_POSIX) this.skip();
      fs.mkdirSync(configDir, { recursive: true });
      const victim = path.join(configDir, 'victim.json');
      fs.writeFileSync(victim, 'do-not-overwrite', 'utf8');
      fs.symlinkSync(victim, tokenPath);

      expect(() => oauth.saveTokens(makeTokens())).to.throw(/symbolic link/);
      expect(fs.readFileSync(victim, 'utf8')).to.equal('do-not-overwrite');
      expect(tempFilesIn(configDir)).to.deep.equal([]);
    });

    it('saveTokens cleans up the temp file when the final rename fails', () => {
      fs.mkdirSync(configDir, { recursive: true });
      // A directory occupies the destination, so rename(temp, TOKEN_PATH) must fail.
      fs.mkdirSync(tokenPath);

      expect(() => oauth.saveTokens(makeTokens())).to.throw();
      expect(tempFilesIn(configDir)).to.deep.equal([]);
    });

    it('loadTokens tightens an over-permissive existing token file to 0600 before reading', function () {
      if (!IS_POSIX) this.skip();
      fs.mkdirSync(configDir, { recursive: true });
      fs.writeFileSync(tokenPath, JSON.stringify(makeTokens()), { mode: 0o644 });

      const loaded = oauth.loadTokens();

      expect(loaded?.access_token).to.equal('ya29.test-access-token');
      expect(fs.statSync(tokenPath).mode & 0o777).to.equal(0o600);
    });

    it('loadTokens returns null when no token file exists', () => {
      fs.mkdirSync(configDir, { recursive: true });
      expect(oauth.loadTokens()).to.equal(null);
    });
  });

  describe('saveBufferAtomic export artifacts', () => {
    let tmpDir: string;

    beforeEach(async () => {
      tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gsheet-atomic-'));
    });

    afterEach(async () => {
      await fs.promises.rm(tmpDir, { recursive: true, force: true }).catch(() => {});
    });

    it('writes the output file owner-only (0600) and leaves no temp files', async () => {
      const buffer = Buffer.from('workbook-bytes');
      const target = path.join(tmpDir, 'out.xlsx');

      const result = await saveBufferAtomic(buffer, target);

      expect(result.savedPath).to.equal(target);
      expect(result.bytesWritten).to.equal(buffer.length);
      expect(result.isOverwritten).to.equal(false);
      if (IS_POSIX) {
        expect(fs.statSync(target).mode & 0o777).to.equal(0o600);
      }
      expect(tempFilesIn(tmpDir)).to.deep.equal([]);
    });

    it('keeps overwrite consent and then replaces the file without widening access', async () => {
      const target = path.join(tmpDir, 'out.xlsx');
      await saveBufferAtomic(Buffer.from('v1'), target);

      let consentError: Error | undefined;
      try {
        await saveBufferAtomic(Buffer.from('v2'), target);
      } catch (error) {
        consentError = error as Error;
      }
      expect(consentError?.message).to.match(/already exists/);

      const overwritten = await saveBufferAtomic(Buffer.from('v2'), target, { overwrite: true });

      expect(overwritten.isOverwritten).to.equal(true);
      expect(fs.readFileSync(target, 'utf8')).to.equal('v2');
      if (IS_POSIX) {
        expect(fs.statSync(target).mode & 0o777).to.equal(0o600);
      }
      expect(tempFilesIn(tmpDir)).to.deep.equal([]);
    });

    it('rejects on hash conflict and leaves the on-disk file untouched', async () => {
      const target = path.join(tmpDir, 'out.xlsx');
      const first = await saveBufferAtomic(Buffer.from('original'), target);
      fs.writeFileSync(target, 'tampered');

      let conflictError: Error | undefined;
      try {
        await saveBufferAtomic(Buffer.from('next'), target, {
          overwrite: true,
          expectedHash: first.sha256Hash,
        });
      } catch (error) {
        conflictError = error as Error;
      }

      expect(conflictError?.message).to.match(/Conflict detected/);
      expect(fs.readFileSync(target, 'utf8')).to.equal('tampered');
    });
  });
});
