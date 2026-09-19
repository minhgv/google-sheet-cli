# OAuth 2.0 Upgrade Plan for google-sheet-cli

## Context
The `google-sheet-cli` currently only supports Service Account JWT authentication (`client_email` + `private_key` or `credentialsFile`). This requires users to manually share their Google Sheets with a service account email. To improve UX for local CLI usage, we need to add OAuth 2.0 User Authentication, allowing users to log in with their personal Google account via a browser popup.

## Approach
We will implement the **OAuth 2.0 Local Server Flow** (Loopback IP address flow) for desktop/CLI applications:
1. **New Commands**: `auth:login`, `auth:logout`, `auth:status`.
2. **Local Server**: During `auth:login`, a temporary HTTP server starts on `localhost` (e.g., port 3000 or dynamic) to catch the OAuth redirect URI.
3. **Token Storage**: Access and refresh tokens are stored securely in `~/.config/google-sheet-cli/token.json` (or OS-specific config dir).
4. **Auth Resolution**: `GoogleSheet.authorize()` will be updated to check for OAuth tokens first (if `--useOauth` flag or env var is set, or if no service account credentials are provided), falling back to Service Account JWT.

## Critical files and ownership
- `src/lib/oauth.ts`: New file handling the local server, token exchange, storage, and refresh logic.
- `src/commands/auth/login.ts`: New command to trigger the browser flow.
- `src/commands/auth/logout.ts`: New command to clear tokens.
- `src/commands/auth/status.ts`: New command to check auth state.
- `src/lib/google-sheet.ts`: Update `authorize()` to accept `OAuth2Client`.
- `src/lib/base-class.ts`: Update flag parsing to include OAuth options.
- `package.json`: Add dependencies (`google-auth-library`, `open`, `get-port` or similar).

## Verification
- **AC-01**: `google-sheet auth:login` opens browser, catches callback, and saves `token.json`.
- **AC-02**: `google-sheet data:get` works using the saved OAuth token without requiring `-f` or `-c/-p`.
- **AC-03**: `google-sheet auth:logout` deletes `token.json`.
- **AC-04**: Edge cases handled: expired token auto-refresh, missing client_secret.json, user denying consent, port already in use.

## Execution checklist
- [ ] T-01: Write durable work record (this file).
- [ ] T-02: Install dependencies (`google-auth-library`, `open`, `get-port`).
- [ ] T-03: Implement `src/lib/oauth.ts` (token storage, local server, refresh).
- [ ] T-04: Implement `auth:login`, `auth:logout`, `auth:status` commands.
- [ ] T-05: Update `GoogleSheet.authorize()` and `base-class.ts` to support OAuth.
- [ ] T-06: Write unit tests for OAuth flow and edge cases.
- [ ] T-07: Run `npm run test:unit` and `npm run build`.
- [ ] T-08: Smoke test with real Google account.

## Evidence and handoff
- Pending implementation.

## Assumptions and contingencies
- User must provide their own `client_secret.json` (Desktop App type) from Google Cloud Console.
- The CLI will default to OAuth if `GSHEET_OAUTH_TOKEN` or `~/.config/google-sheet-cli/token.json` exists and no service account flags are passed.
