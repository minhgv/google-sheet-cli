import { Command, Flags } from '@oclif/core';
import { authenticate, CLIENT_SECRET_PATH } from '../../lib/oauth';

export default class AuthLogin extends Command {
  static description = 'Authenticate with your Google account via OAuth 2.0';

  static examples = [
    `$ gsheet auth:login
Opening browser for authentication...
Authentication successful! Tokens saved to ~/.config/google-sheet-cli/token.json
`,
    `$ gsheet auth:login --clientSecretFile=./client_secret.json
Opening browser for authentication...
Authentication successful! Tokens saved to ~/.config/google-sheet-cli/token.json
`,
  ];

  static flags = {
    clientSecretFile: Flags.string({
      description: `Path to OAuth 2.0 client_secret.json (Desktop App type). Defaults to ${CLIENT_SECRET_PATH}`,
      required: false,
      env: 'GSHEET_CLIENT_SECRET_FILE',
    }),
  };

  async run() {
    const { flags } = await this.parse(AuthLogin);

    this.log('Starting OAuth 2.0 authentication...');
    this.log(`Using client secret: ${flags.clientSecretFile || CLIENT_SECRET_PATH}`);

    try {
      const tokens = await authenticate(flags.clientSecretFile);
      this.log('Authentication successful!');
      this.log(`Access token expires: ${new Date(tokens.expiry_date).toLocaleString()}`);
      if (tokens.refresh_token) {
        this.log('Refresh token saved - you will not need to re-authenticate.');
      }
    } catch (error) {
      this.error(`Authentication failed: ${(error as Error).message}`, { exit: 1 });
    }
  }
}
