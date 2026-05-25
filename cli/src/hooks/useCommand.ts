import { useState, useCallback, useRef, useEffect } from "react";
import { spawn, ChildProcess } from "node:child_process";
import type { CommandResult, CommandStatus } from "../types.js";

export function useCommand() {
  const [result, setResult] = useState<CommandResult>({
    status: "idle",
    output: [],
    error: [],
    exitCode: null,
  });
  const childRef = useRef<ChildProcess | null>(null);

  const run = useCallback((cmd: string, args: string[], options?: { cwd?: string; env?: NodeJS.ProcessEnv }) => {
    setResult({ status: "running", output: [], error: [], exitCode: null });

    const child = spawn(cmd, args, {
      cwd: options?.cwd,
      env: { ...process.env, ...options?.env },
      shell: false,
    });

    childRef.current = child;
    const output: string[] = [];
    const error: string[] = [];

    child.stdout?.on("data", (data) => {
      const lines = data.toString().split("\n").filter(Boolean);
      output.push(...lines);
      setResult((prev) => ({ ...prev, output: [...output] }));
    });

    child.stderr?.on("data", (data) => {
      const lines = data.toString().split("\n").filter(Boolean);
      error.push(...lines);
      setResult((prev) => ({ ...prev, error: [...error] }));
    });

    child.on("close", (code) => {
      const status: CommandStatus = code === 0 ? "success" : "error";
      setResult({ status, output, error, exitCode: code });
      childRef.current = null;
    });

    child.on("error", (err) => {
      error.push(err.message);
      setResult({ status: "error", output, error, exitCode: 1 });
      childRef.current = null;
    });
  }, []);

  const kill = useCallback(() => {
    if (childRef.current) {
      childRef.current.kill("SIGTERM");
    }
  }, []);

  useEffect(() => {
    return () => {
      if (childRef.current) {
        childRef.current.kill("SIGTERM");
      }
    };
  }, []);

  return { result, run, kill };
}
