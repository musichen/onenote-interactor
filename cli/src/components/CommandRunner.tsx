import React, { useEffect } from "react";
import { Box, Text } from "ink";
import { useCommand } from "../hooks/useCommand.js";
import Spinner from "ink-spinner";

export type CommandRunnerProps = {
  cmd: string;
  args: string[];
  cwd?: string;
  env?: NodeJS.ProcessEnv;
  onComplete?: (exitCode: number | null) => void;
  label?: string;
  maxLines?: number;
};

export function CommandRunner({ cmd, args, cwd, env, onComplete, label, maxLines = 50 }: CommandRunnerProps) {
  const { result, run } = useCommand();

  useEffect(() => {
    run(cmd, args, { cwd, env });
  }, []);

  useEffect(() => {
    if (result.status === "success" || result.status === "error") {
      onComplete?.(result.exitCode);
    }
  }, [result.status]);

  const showOutput = result.status === "running" || result.status === "success" || result.status === "error";
  const lines = result.output.slice(-maxLines);

  return (
    <Box flexDirection="column">
      {label && (
        <Box>
          {result.status === "running" && (
            <Text color="cyan">
              <Spinner type="dots" /> {label}
            </Text>
          )}
          {result.status === "success" && <Text color="green">✓ {label}</Text>}
          {result.status === "error" && <Text color="red">✗ {label}</Text>}
          {result.status === "idle" && <Text color="gray">○ {label}</Text>}
        </Box>
      )}
      {showOutput && lines.length > 0 && (
        <Box flexDirection="column" marginLeft={2} marginTop={1}>
          {lines.map((line, i) => (
            <Text key={i} color="gray" dimColor>
              {line}
            </Text>
          ))}
        </Box>
      )}
      {result.error.length > 0 && (
        <Box flexDirection="column" marginLeft={2} marginTop={1}>
          {result.error.slice(-10).map((line, i) => (
            <Text key={i} color="red">
              {line}
            </Text>
          ))}
        </Box>
      )}
    </Box>
  );
}
