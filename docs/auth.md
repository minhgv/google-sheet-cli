`google-sheet auth`
===================

Authenticate with your Google account via OAuth 2.0

* [`google-sheet auth:login`](#google-sheet-authlogin)
* [`google-sheet auth:logout`](#google-sheet-authlogout)
* [`google-sheet auth:status`](#google-sheet-authstatus)

## `google-sheet auth:login`

Authenticate with your Google account via OAuth 2.0

```
USAGE
  $ google-sheet auth:login [--clientSecretFile <value>]

FLAGS
  --clientSecretFile=<value>  [env: GSHEET_CLIENT_SECRET_FILE] Path to OAuth 2.0 client_secret.json (Desktop App type).
                              Defaults to /Users/giapminh79/.config/google-sheet-cli/client_secret.json

DESCRIPTION
  Authenticate with your Google account via OAuth 2.0

EXAMPLES
  $ gsheet auth:login
  Opening browser for authentication...
  Authentication successful! Tokens saved to ~/.config/google-sheet-cli/token.json

  $ gsheet auth:login --clientSecretFile=./client_secret.json
  Opening browser for authentication...
  Authentication successful! Tokens saved to ~/.config/google-sheet-cli/token.json
```

_See code: [src/commands/auth/login.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/auth/login.ts)_

## `google-sheet auth:logout`

Remove stored OAuth 2.0 tokens

```
USAGE
  $ google-sheet auth:logout

DESCRIPTION
  Remove stored OAuth 2.0 tokens

EXAMPLES
  $ gsheet auth:logout
  OAuth tokens removed from ~/.config/google-sheet-cli/token.json
```

_See code: [src/commands/auth/logout.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/auth/logout.ts)_

## `google-sheet auth:status`

Check OAuth 2.0 authentication status

```
USAGE
  $ google-sheet auth:status

DESCRIPTION
  Check OAuth 2.0 authentication status

EXAMPLES
  $ gsheet auth:status
  Authenticated: yes
  Token expires: 2026-09-18 15:30:00
  Refresh token: available
```

_See code: [src/commands/auth/status.ts](https://github.com/jroehl/google-sheet-cli/blob/v3.1.0/src/commands/auth/status.ts)_
