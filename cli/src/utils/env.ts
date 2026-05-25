import fs from "node:fs/promises";
import path from "node:path";

const ROOT_DIR = path.resolve(import.meta.dirname, "../../..");
const ENV_PATH = path.join(ROOT_DIR, ".env.local");

export async function envFileExists(): Promise<boolean> {
  try {
    await fs.access(ENV_PATH);
    return true;
  } catch {
    return false;
  }
}

export async function readEnvFile(): Promise<Record<string, string>> {
  const env: Record<string, string> = {};
  try {
    const content = await fs.readFile(ENV_PATH, "utf-8");
    for (const line of content.split("\n")) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) continue;
      const eq = trimmed.indexOf("=");
      if (eq === -1) continue;
      const key = trimmed.slice(0, eq).trim();
      const value = trimmed.slice(eq + 1).trim().replace(/^["']|["']$/g, "");
      env[key] = value;
    }
  } catch {
    // ignore
  }
  return env;
}

export async function writeEnvFile(values: Record<string, string>): Promise<void> {
  const lines = Object.entries(values).map(([k, v]) => `${k}=${v}`);
  await fs.writeFile(ENV_PATH, lines.join("\n") + "\n", "utf-8");
}

export function isEnvValid(env: Record<string, string>): boolean {
  return !!env.ONENOTE_CLIENT_ID && env.ONENOTE_CLIENT_ID.length > 10;
}
