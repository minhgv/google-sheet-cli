import { expect } from 'chai';
import { existsSync, mkdirSync, rmSync, writeFileSync } from 'fs';
import { join } from 'path';
const os = require('os');
const originalHomedir = os.homedir;
const originalHomeEnv = process.env.HOME;

const TEST_HOME = join(__dirname, '..', 'tmp', 'test-oauth-home');
os.homedir = () => TEST_HOME;
process.env.HOME = TEST_HOME;

// Clear require cache to ensure oauth initializes with the stubbed homedir
delete require.cache[require.resolve('../src/lib/oauth')];
const {
  loadTokens,
  saveTokens,
  deleteTokens,
  readClientSecret,
  isTokenExpired,
  TOKEN_PATH,
  CLIENT_SECRET_PATH,
} = require('../src/lib/oauth');
import type { OAuthTokens } from '../src/lib/oauth';
const CONFIG_DIR = join(TEST_HOME, '.config', 'google-sheet-cli');
const TEST_TOKEN_PATH = TOKEN_PATH;
const TEST_CLIENT_SECRET_PATH = CLIENT_SECRET_PATH;

const VALID_TOKENS: OAuthTokens = {
  access_token: 'test-access-token',
  refresh_token: 'test-refresh-token',
  scope: 'https://www.googleapis.com/auth/spreadsheets',
  token_type: 'Bearer',
  expiry_date: Date.now() + 3600 * 1000,
};

const VALID_CLIENT_SECRET = {
  installed: {
    client_id: 'test-client-id.apps.googleusercontent.com',
    client_secret: 'test-client-secret',
    redirect_uris: ['http://localhost'],
    auth_uri: 'https://accounts.google.com/o/oauth2/auth',
    token_uri: 'https://oauth2.googleapis.com/token',
  },
};
describe('oauth', () => {
  before(() => {
    if (!existsSync(CONFIG_DIR)) {
      mkdirSync(CONFIG_DIR, { recursive: true });
    }
  });
  after(() => {
    os.homedir = originalHomedir;
    if (originalHomeEnv !== undefined) {
      process.env.HOME = originalHomeEnv;
    } else {
      delete process.env.HOME;
    }
    delete require.cache[require.resolve('../src/lib/oauth')];
    if (existsSync(TEST_HOME)) {
      rmSync(TEST_HOME, { recursive: true, force: true });
    }
  });


  afterEach(() => {
    if (existsSync(TEST_TOKEN_PATH)) rmSync(TEST_TOKEN_PATH);
    if (existsSync(TEST_CLIENT_SECRET_PATH)) rmSync(TEST_CLIENT_SECRET_PATH);
  });

  describe('saveTokens / loadTokens / deleteTokens', () => {
    it('saves and loads tokens', () => {
      saveTokens(VALID_TOKENS);
      const loaded = loadTokens();
      expect(loaded).to.deep.equal(VALID_TOKENS);
    });

    it('returns null when no tokens exist', () => {
      expect(loadTokens()).to.be.null;
    });

    it('returns null for corrupted token file', () => {
      writeFileSync(TEST_TOKEN_PATH, 'not-json', 'utf8');
      expect(loadTokens()).to.be.null;
    });

    it('returns null for token file missing access_token', () => {
      writeFileSync(TEST_TOKEN_PATH, JSON.stringify({ refresh_token: 'x' }), 'utf8');
      expect(loadTokens()).to.be.null;
    });

    it('deletes tokens', () => {
      saveTokens(VALID_TOKENS);
      deleteTokens();
      expect(loadTokens()).to.be.null;
    });

    it('deleteTokens is idempotent when no file exists', () => {
      expect(() => deleteTokens()).to.not.throw();
    });
  });

  describe('readClientSecret', () => {
    it('reads a valid client secret file', () => {
      writeFileSync(TEST_CLIENT_SECRET_PATH, JSON.stringify(VALID_CLIENT_SECRET), 'utf8');
      const secret = readClientSecret();
      expect(secret.installed.client_id).to.equal('test-client-id.apps.googleusercontent.com');
    });

    it('throws when file does not exist', () => {
      expect(() => readClientSecret()).to.throw('Client secret file not found');
    });

    it('throws for invalid JSON', () => {
      writeFileSync(TEST_CLIENT_SECRET_PATH, 'not-json', 'utf8');
      expect(() => readClientSecret()).to.throw('Cannot read client secret file');
    });

    it('throws for missing installed object', () => {
      writeFileSync(TEST_CLIENT_SECRET_PATH, JSON.stringify({ web: {} }), 'utf8');
      expect(() => readClientSecret()).to.throw('Invalid client secret format');
    });

    it('throws for missing client_id', () => {
      writeFileSync(
        TEST_CLIENT_SECRET_PATH,
        JSON.stringify({ installed: { client_secret: 'x' } }),
        'utf8'
      );
      expect(() => readClientSecret()).to.throw('Invalid client secret format');
    });

    it('accepts custom path', () => {
      const customPath = join(CONFIG_DIR, 'custom-secret.json');
      writeFileSync(customPath, JSON.stringify(VALID_CLIENT_SECRET), 'utf8');
      const secret = readClientSecret(customPath);
      expect(secret.installed.client_id).to.equal('test-client-id.apps.googleusercontent.com');
      rmSync(customPath);
    });
  });

  describe('isTokenExpired', () => {
    it('returns false for valid token', () => {
      expect(isTokenExpired(VALID_TOKENS)).to.be.false;
    });

    it('returns true for expired token', () => {
      const expired = { ...VALID_TOKENS, expiry_date: Date.now() - 1000 };
      expect(isTokenExpired(expired)).to.be.true;
    });

    it('returns true for token expiring within buffer', () => {
      const soonExpired = { ...VALID_TOKENS, expiry_date: Date.now() + 60 * 1000 };
      expect(isTokenExpired(soonExpired)).to.be.true;
    });

    it('returns false for token expiring after buffer', () => {
      const valid = { ...VALID_TOKENS, expiry_date: Date.now() + 10 * 60 * 1000 };
      expect(isTokenExpired(valid)).to.be.false;
    });
  });
});
