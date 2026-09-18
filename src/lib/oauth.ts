import { OAuth2Client } from 'google-auth-library';
import { createServer, IncomingMessage, Server, ServerResponse } from 'http';
import { readFileSync, writeFileSync, existsSync, mkdirSync, unlinkSync } from 'fs';
import { join } from 'path';
import { homedir } from 'os';
import { URL } from 'url';
import { log } from './log';

const DEBUG_NAMESPACE = 'gsheet:oauth';
const CONFIG_DIR = join(homedir(), '.config', 'google-sheet-cli');
export const TOKEN_PATH = join(CONFIG_DIR, 'token.json');
export const CLIENT_SECRET_PATH = join(CONFIG_DIR, 'client_secret.json');

export interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  scope: string;
  token_type: string;
  expiry_date: number;
}

export interface ClientSecret {
  installed: {
    client_id: string;
    client_secret: string;
    redirect_uris: string[];
    auth_uri: string;
    token_uri: string;
  };
}

/**
 * Log OAuth debug messages.
 */
const debug = (message: string): void => log(DEBUG_NAMESPACE, message, true);

/**
 * Ensure the config directory exists.
 */
const ensureConfigDir = (): void => {
  if (!existsSync(CONFIG_DIR)) {
    mkdirSync(CONFIG_DIR, { recursive: true });
  }
};

/**
 * Read the client secret file.
 */
export const readClientSecret = (path?: string): ClientSecret => {
  const secretPath = path || CLIENT_SECRET_PATH;
  if (!existsSync(secretPath)) {
    throw new Error(
      `Client secret file not found at ${secretPath}. ` +
      `Please download it from Google Cloud Console (APIs & Services > Credentials > OAuth 2.0 Client IDs > Desktop App) ` +
      `and save it to ${CLIENT_SECRET_PATH} or provide the path via --clientSecretFile.`
    );
  }

  try {
    const parsed = JSON.parse(readFileSync(secretPath, 'utf8'));
    if (!parsed.installed || !parsed.installed.client_id || !parsed.installed.client_secret) {
      throw new Error('Invalid client secret format. Expected "installed" object with client_id and client_secret.');
    }
    return parsed as ClientSecret;
  } catch (error) {
    throw new Error(`Cannot read client secret file ${secretPath}: ${(error as Error).message}`);
  }
};

/**
 * Save OAuth tokens to disk.
 */
export const saveTokens = (tokens: OAuthTokens): void => {
  ensureConfigDir();
  writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2), 'utf8');
  debug(`Tokens saved to ${TOKEN_PATH}`);
};

/**
 * Load OAuth tokens from disk.
 */
export const loadTokens = (): OAuthTokens | null => {
  if (!existsSync(TOKEN_PATH)) {
    return null;
  }

  try {
    const parsed = JSON.parse(readFileSync(TOKEN_PATH, 'utf8'));
    if (!parsed.access_token) {
      return null;
    }
    return parsed as OAuthTokens;
  } catch (error) {
    debug(`Failed to load tokens: ${(error as Error).message}`);
    return null;
  }
};

/**
 * Delete stored OAuth tokens.
 */
export const deleteTokens = (): void => {
  if (existsSync(TOKEN_PATH)) {
    unlinkSync(TOKEN_PATH);
    debug(`Tokens deleted from ${TOKEN_PATH}`);
  }
};

/**
 * Create an OAuth2Client from stored tokens.
 */
export const createOAuth2Client = (clientSecret: ClientSecret, tokens: OAuthTokens): OAuth2Client => {
  const { client_id, client_secret, redirect_uris } = clientSecret.installed;
  const redirectUri = redirect_uris?.[0] || 'http://localhost';
  const client = new OAuth2Client(client_id, client_secret, redirectUri);
  client.setCredentials(tokens);
  return client;
};

/**
 * Check if tokens are expired or about to expire.
 */
export const isTokenExpired = (tokens: OAuthTokens, bufferMs = 5 * 60 * 1000): boolean => {
  return Date.now() >= tokens.expiry_date - bufferMs;
};

/**
 * Refresh the access token using the refresh token.
 */
export const refreshTokens = async (client: OAuth2Client): Promise<OAuthTokens> => {
  try {
    const { credentials } = await client.refreshAccessToken();
    const newTokens: OAuthTokens = {
      access_token: credentials.access_token!,
      refresh_token: credentials.refresh_token ?? client.credentials.refresh_token ?? undefined,
      scope: credentials.scope || '',
      token_type: credentials.token_type || 'Bearer',
      expiry_date: credentials.expiry_date || Date.now() + 3600 * 1000,
    };
    saveTokens(newTokens);
    return newTokens;
  } catch (error) {
    throw new Error(`Failed to refresh access token: ${(error as Error).message}`);
  }
};

/**
 * Get a valid OAuth2Client, refreshing tokens if necessary.
 */
export const getAuthenticatedClient = async (clientSecretPath?: string): Promise<OAuth2Client> => {
  const clientSecret = readClientSecret(clientSecretPath);
  let tokens = loadTokens();

  if (!tokens) {
    throw new Error(
      'No OAuth tokens found. Please run "google-sheet auth:login" first to authenticate with your Google account.'
    );
  }

  const client = createOAuth2Client(clientSecret, tokens);

  if (isTokenExpired(tokens)) {
    if (!tokens.refresh_token) {
      throw new Error('Access token expired and no refresh token available. Please run "google-sheet auth:login" again.');
    }
    debug('Access token expired, refreshing...');
    tokens = await refreshTokens(client);
    client.setCredentials(tokens);
  }

  return client;
};

/**
 * Start the OAuth 2.0 local server flow.
 * Opens the browser for user consent and waits for the callback.
 */
export const authenticate = async (clientSecretPath?: string): Promise<OAuthTokens> => {
  const clientSecret = readClientSecret(clientSecretPath);
  const { client_id, client_secret } = clientSecret.installed;

  const getPort = (await import('get-port')).default;
  const port = await getPort({ port: [3000, 3001, 3002, 8080, 8081] });
  const redirectUri = `http://localhost:${port}/oauth2callback`;

  const client = new OAuth2Client(client_id, client_secret, redirectUri);

  const authUrl = client.generateAuthUrl({
    access_type: 'offline',
    scope: ['https://www.googleapis.com/auth/spreadsheets'],
    prompt: 'consent', // Force consent to get refresh_token
  });

  return new Promise((resolve, reject) => {
    const server: Server = createServer(async (req: IncomingMessage, res: ServerResponse) => {
      try {
        if (!req.url?.startsWith('/oauth2callback')) {
          res.writeHead(404);
          res.end('Not found');
          return;
        }

        const url = new URL(req.url, `http://localhost:${port}`);
        const code = url.searchParams.get('code');
        const error = url.searchParams.get('error');

        if (error) {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end(`<h1>Authentication failed</h1><p>Error: ${error}</p><p>You can close this window.</p>`);
          server.close();
          reject(new Error(`OAuth error: ${error}`));
          return;
        }

        if (!code) {
          res.writeHead(400, { 'Content-Type': 'text/html' });
          res.end('<h1>Authentication failed</h1><p>No authorization code received.</p>');
          server.close();
          reject(new Error('No authorization code received'));
          return;
        }

        const { tokens } = await client.getToken(code);
        const oauthTokens: OAuthTokens = {
          access_token: tokens.access_token!,
          refresh_token: tokens.refresh_token ?? undefined,
          scope: tokens.scope || '',
          token_type: tokens.token_type || 'Bearer',
          expiry_date: tokens.expiry_date || Date.now() + 3600 * 1000,
        };

        saveTokens(oauthTokens);

        res.writeHead(200, { 'Content-Type': 'text/html' });
        res.end('<h1>Authentication successful!</h1><p>You can close this window and return to the terminal.</p>');
        server.close();
        resolve(oauthTokens);
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'text/html' });
        res.end(`<h1>Authentication failed</h1><p>${(error as Error).message}</p>`);
        server.close();
        reject(error);
      }
    });

    server.listen(port, async () => {
      debug(`OAuth callback server listening on http://localhost:${port}`);
      try {
        const open = (await import('open')).default;
        await open(authUrl);
      } catch (err) {
        debug(`Failed to open browser: ${(err as Error).message}`);
        console.log(`Please open this URL in your browser:\n${authUrl}`);
      }
    });

    // Timeout after 5 minutes
    setTimeout(() => {
      server.close();
      reject(new Error('Authentication timed out. Please try again.'));
    }, 5 * 60 * 1000);
  });
};

export default {
  authenticate,
  getAuthenticatedClient,
  loadTokens,
  saveTokens,
  deleteTokens,
  readClientSecret,
  TOKEN_PATH,
  CLIENT_SECRET_PATH,
};
