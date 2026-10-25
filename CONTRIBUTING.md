# Contributing to onenote-interactor

## Setup

```bash
pnpm install
```

Microsoft Graph work needs an Azure app and a `.env.local`, created through the onboarding wizard in the interactive CLI (`pnpm cli`). Never commit `.env.local` or the persisted MSAL token cache.

## Checks

```bash
pnpm test
```

This runs `node --test tests/**/*.test.js`. There is no separate lint or typecheck script.

## Try it

- Interactive CLI: `pnpm cli`
- Direct commands, for example: `pnpm run graph:list`, `pnpm run graph:status`, `pnpm run graph:audit`

## Pull requests

Create a branch from `main`. For Graph changes, say which command you ran against a real or test notebook (redact notebook names and tokens where needed). For local-only changes, the test suite above is enough.
