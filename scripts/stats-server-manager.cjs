#!/usr/bin/env node
/**
 * stats-server-manager.cjs
 * Start/stop/status helper for the stats dashboard server.
 * Tracks PID in .stats-server.pid for clean lifecycle management.
 */

const fs = require("fs").promises;
const { spawn } = require("child_process");
const path = require("path");

const PID_FILE = path.join(process.cwd(), ".stats-server.pid");
const LOG_FILE = path.join(process.cwd(), ".stats-server.log");
const DEFAULT_ROOT = "exports/graph/A";
const DEFAULT_PORT = "9876";

async function readPid() {
  try {
    const pid = await fs.readFile(PID_FILE, "utf-8");
    return parseInt(pid.trim(), 10);
  } catch {
    return null;
  }
}

function isRunning(pid) {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function start() {
  const existingPid = await readPid();
  if (existingPid && isRunning(existingPid)) {
    console.log(`Stats server already running (PID: ${existingPid})`);
    return;
  }

  // Clean up stale PID file
  if (existingPid) {
    try { await fs.unlink(PID_FILE); } catch {}
  }

  const out = await fs.open(LOG_FILE, "a");
  const child = spawn("node", [
    "src/stats-server.js",
    "--root", DEFAULT_ROOT,
    "--port", DEFAULT_PORT
  ], {
    detached: true,
    stdio: ["ignore", out, out]
  });

  child.unref();
  await fs.writeFile(PID_FILE, String(child.pid), "utf-8");

  // Verify process is still alive after a moment
  await delay(500);
  if (!isRunning(child.pid)) {
    console.log(`Stats server failed to start (PID: ${child.pid})`);
    console.log(`Check log: ${LOG_FILE}`);
    try { await fs.unlink(PID_FILE); } catch {}
    try { out.close(); } catch {}
    return;
  }

  out.close();
  console.log(`Stats server started (PID: ${child.pid})`);
  console.log(`Dashboard: http://127.0.0.1:${DEFAULT_PORT}`);
}

async function status() {
  const pid = await readPid();
  if (!pid) {
    console.log("Stats server: not running");
    return;
  }
  if (isRunning(pid)) {
    console.log(`Stats server: running (PID: ${pid})`);
  } else {
    console.log("Stats server: not running (stale PID file removed)");
    try { await fs.unlink(PID_FILE); } catch {}
  }
}

async function stop() {
  const pid = await readPid();
  if (!pid) {
    console.log("Stats server: not running");
    return;
  }
  if (isRunning(pid)) {
    try {
      process.kill(pid, "SIGTERM");
      console.log(`Stats server stopped (PID: ${pid})`);
    } catch (err) {
      console.log(`Failed to stop stats server: ${err.message}`);
    }
  } else {
    console.log("Stats server: not running (stale PID file removed)");
  }
  try { await fs.unlink(PID_FILE); } catch {}
}

const command = process.argv[2];
if (command === "start") {
  start().catch(console.error);
} else if (command === "status") {
  status().catch(console.error);
} else if (command === "stop") {
  stop().catch(console.error);
} else {
  console.log("Usage: node scripts/stats-server-manager.cjs {start|status|stop}");
  process.exit(1);
}
