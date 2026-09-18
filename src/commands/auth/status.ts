import { Command } from '@oclif/core';
import { loadTokens, TOKEN_PATH, isTokenExpired } from '../../lib/oauth';

export default class AuthStatus extends Command {
  static description = 'Check OAuth 2.0 authentication status';

  static examples = [
    `$ gsheet auth:status
Authenticated: yes
Token expires: 2026-09-18 15:30:00
Refresh token: available
`,
  ];

  async run() {
    const tokens = loadTokens();

    if (!tokens) {
      this.log('Authenticated: no');
      this.log(`No tokens found at ${TOKEN_PATH}`);
      this.log('Run "google-sheet auth:login" to authenticate.');
      return;
    }

    this.log('Authenticated: yes');
    this.log(`Token expires: ${new Date(tokens.expiry_date).toLocaleString()}`);
    this.log(`Refresh token: ${tokens.refresh_token ? 'available' : 'not available'}`);

    if (isTokenExpired(tokens)) {
      this.log('Status: expired (will auto-refresh on next use)');
    } else {
      this.log('Status: valid');
    }
  }
}
