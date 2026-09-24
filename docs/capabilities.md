`google-sheet capabilities`
===========================

Print the machine-readable capability document: per backend (google-sheets, local-xlsx) the supported operations, input and output forms, formula write-vs-recalculation semantics, destructive guards, mutation limits and local fidelity exclusions. Credential-free.

* [`google-sheet capabilities`](#google-sheet-capabilities)

## `google-sheet capabilities`

Print the machine-readable capability document: per backend (google-sheets, local-xlsx) the supported operations, input and output forms, formula write-vs-recalculation semantics, destructive guards, mutation limits and local fidelity exclusions. Credential-free.

```
USAGE
  $ google-sheet capabilities

DESCRIPTION
  Print the machine-readable capability document: per backend (google-sheets, local-xlsx) the supported operations,
  input and output forms, formula write-vs-recalculation semantics, destructive guards, mutation limits and local
  fidelity exclusions. Credential-free.

EXAMPLES
  $ gsheet capabilities
  {
    "schemaVersion": 1,
    "supportModel": "...",
    "visibility": [...],
    "backends": { "google-sheets": {...}, "local-xlsx": {...} }
  }

  $ gsheet capabilities | jq '.backends["local-xlsx"].formulaSemantics'
```

_See code: [src/commands/capabilities.ts](https://github.com/jroehl/google-sheet-cli/blob/master/src/commands/capabilities.ts)_
