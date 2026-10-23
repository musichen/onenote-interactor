# Security Policy

## Reporting a vulnerability

Open a GitHub issue in this repository for non-sensitive reports. For a sensitive vulnerability, contact the maintainer privately first through the GitHub profile at https://github.com/musichen instead of posting details publicly.

## Notes for users

- Keep `.env.local` private. It holds your Azure app and Microsoft Graph settings from the onboarding wizard. Never commit it.
- The tool persists an MSAL token cache so exports can resume. Treat that cache like a password: do not share or commit it.
- Exports contain your notebook content (pages as HTML and Markdown, plus downloaded resources). Store the `exports/` output somewhere you trust.
