import React, { useState, useEffect } from "react";
import { Box, Text } from "ink";
import { CommandRunner } from "../components/CommandRunner.js";
import { StatusBadge } from "../components/StatusBadge.js";
import { ROOT_DIR } from "../utils/paths.js";

export type AuthProps = {
  onComplete: () => void;
};

export function Auth({ onComplete }: AuthProps) {
  const [step, setStep] = useState<"check" | "login" | "done">("check");

  useEffect(() => {
    // Check if token cache exists
    const cachePath = `${process.env.HOME}/.config/onenote-interactor/msal-cache.json`;
    import("node:fs/promises")
      .then((fs) => fs.access(cachePath))
      .then(() => {
        setStep("done");
        setTimeout(onComplete, 500);
      })
      .catch(() => setStep("login"));
  }, []);

  if (step === "check") {
    return (
      <Box>
        <StatusBadge status="info" label="Checking authentication..." />
      </Box>
    );
  }

  if (step === "login") {
    return (
      <Box flexDirection="column">
        <Text bold color="yellow">
          🔐 Authentication Required
        </Text>
        <Text color="gray">A Microsoft device code will appear below. Open the URL and enter the code.</Text>
        <Box marginTop={1}>
          <CommandRunner
            cmd="node"
            args={["src/onenote-interactor.js", "graph-login"]}
            cwd={ROOT_DIR}
            label="Logging in..."
            onComplete={(code) => {
              if (code === 0) {
                setStep("done");
                setTimeout(onComplete, 800);
              }
            }}
          />
        </Box>
      </Box>
    );
  }

  return (
    <Box>
      <StatusBadge status="ok" label="Authenticated" />
    </Box>
  );
}
