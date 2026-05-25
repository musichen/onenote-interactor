#!/usr/bin/env node
import React from "react";
import { render } from "ink";
import { App } from "./app.js";

const args = process.argv.slice(2);

if (args.includes("--help") || args.includes("-h")) {
  console.log(`
  📓 OneNote Interactor — Interactive CLI

  Usage:
    ./cli/bin/onenote-interactor          Launch interactive TUI
    pnpm cli                              Launch interactive TUI
    pnpm cli --help                       Show this help

  The interactive CLI provides:
    • Onboarding wizard (Azure app setup, .env.local creation)
    • Authentication with Microsoft Graph
    • Notebook listing and selection
    • Export, post-process, diff, resync
    • Markdown-only book creation
    • Stats server lifecycle management
    • Sync status dashboard

  Direct commands (non-interactive):
    pnpm run auth:login
    pnpm run graph:export
    pnpm run graph:postprocess
    pnpm run graph:diff
    pnpm run graph:resync
    pnpm run graph:status
    pnpm run stats:server:silent
    pnpm run stats:server:stop
`);
  process.exit(0);
}

render(<App />);
