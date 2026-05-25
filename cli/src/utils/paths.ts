import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const ROOT_DIR = path.resolve(__dirname, "../../..");
export const EXPORTS_DIR = path.join(ROOT_DIR, "exports", "graph");
export const LOGS_DIR = path.join(ROOT_DIR, "logs");
export const INTERACTOR_PATH = path.join(ROOT_DIR, "src", "onenote-interactor.js");
export const STATS_MANAGER_PATH = path.join(ROOT_DIR, "scripts", "stats-server-manager.cjs");
export const MARKDOWN_EXPORTER_PATH = path.join(ROOT_DIR, "scripts", "create-markdown-only-book.cjs");
