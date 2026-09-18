import { Command } from '@oclif/core';
import { deleteTokens, TOKEN_PATH } from '../../lib/oauth';

export default class AuthLogout extends Command {
  static description = 'Remove stored OAuth 2.0 tokens';

  static examples = [
    `$ gsheet auth:logout
OAuth tokens removed from ~/.config/google-sheet-cli/token.json
`,
  ];

  async run() {
    deleteTokens();
    this.log(`OAuth tokens removed from ${TOKEN_PATH}`);
  }
}
