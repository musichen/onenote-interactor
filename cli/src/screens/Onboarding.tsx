import React, { useState, useEffect } from "react";
import { Box, Text } from "ink";
import TextInput from "ink-text-input";
import SelectInput from "ink-select-input";
import { envFileExists, readEnvFile, writeEnvFile, isEnvValid } from "../utils/env.js";
import { StatusBadge } from "../components/StatusBadge.js";

export type OnboardingProps = {
  onComplete: () => void;
};

export function Onboarding({ onComplete }: OnboardingProps) {
  const [step, setStep] = useState<"check" | "explain" | "input" | "confirm" | "done">("check");
  const [env, setEnv] = useState<Record<string, string>>({});
  const [clientId, setClientId] = useState("");

  useEffect(() => {
    envFileExists().then(async (exists) => {
      if (exists) {
        const e = await readEnvFile();
        setEnv(e);
        if (isEnvValid(e)) {
          setStep("done");
          setTimeout(onComplete, 500);
        } else {
          setStep("explain");
        }
      } else {
        setStep("explain");
      }
    });
  }, []);

  if (step === "check") {
    return (
      <Box>
        <StatusBadge status="info" label="Checking environment..." />
      </Box>
    );
  }

  if (step === "explain") {
    const items = [
      { label: "I already have an Azure app registration", value: "input" },
      { label: "Show me how to create one", value: "guide" },
    ];
    return (
      <Box flexDirection="column">
        <Text bold color="yellow">
          ⚠️ Environment not configured
        </Text>
        <Text color="gray">You need an Azure AD app registration to access OneNote via Microsoft Graph.</Text>
        <Box marginTop={1}>
          <SelectInput
            items={items}
            onSelect={(item) => {
              if (item.value === "input") setStep("input");
              else setStep("guide");
            }}
          />
        </Box>
      </Box>
    );
  }

  if (step === "guide") {
    return (
      <Box flexDirection="column">
        <Text bold>Azure App Registration Guide</Text>
        <Box flexDirection="column" marginTop={1}>
          <Text>1. Go to <Text color="cyan">https://portal.azure.com</Text></Text>
          <Text>2. Open <Text bold>App registrations</Text> → <Text bold>New registration</Text></Text>
          <Text>3. Name it <Text color="yellow">onenote-interactor</Text></Text>
          <Text>4. Supported account types: <Text color="yellow">Any Entra ID tenant + Personal Microsoft accounts</Text></Text>
          <Text>5. Create → open <Text bold>Authentication</Text> → enable <Text bold>Allow public client flows</Text></Text>
          <Text>6. Open <Text bold>API permissions</Text> → add:</Text>
          <Text>   - Microsoft Graph → Delegated → <Text color="cyan">Notes.Read</Text></Text>
          <Text>   - Microsoft Graph → Delegated → <Text color="cyan">Notes.Read.All</Text></Text>
          <Text>   - Microsoft Graph → Delegated → <Text color="cyan">offline_access</Text></Text>
          <Text>   - Microsoft Graph → Delegated → <Text color="cyan">User.Read</Text></Text>
          <Text>7. Copy the <Text bold>Application (client) ID</Text></Text>
        </Box>
        <Box marginTop={1}>
          <Text color="gray">Press Enter to continue...</Text>
        </Box>
      </Box>
    );
  }

  if (step === "input") {
    return (
      <Box flexDirection="column">
        <Text bold>Enter your Azure Application (client) ID:</Text>
        <Box marginTop={1}>
          <TextInput
            value={clientId}
            onChange={setClientId}
            onSubmit={() => setStep("confirm")}
            placeholder="cd86f4e7-9e5d-4756-937a-081ab93099f0"
          />
        </Box>
        <Text color="gray" dimColor>
          Paste your client ID and press Enter
        </Text>
      </Box>
    );
  }

  if (step === "confirm") {
    const items = [
      { label: "Save and continue", value: "save" },
      { label: "Edit", value: "edit" },
    ];
    return (
      <Box flexDirection="column">
        <Text>Client ID: <Text color="cyan">{clientId}</Text></Text>
        <Box marginTop={1}>
          <SelectInput
            items={items}
            onSelect={async (item) => {
              if (item.value === "save") {
                await writeEnvFile({
                  ONENOTE_CLIENT_ID: clientId,
                  ONENOTE_TENANT_ID: "consumers",
                });
                setStep("done");
                setTimeout(onComplete, 500);
              } else {
                setStep("input");
              }
            }}
          />
        </Box>
      </Box>
    );
  }

  return (
    <Box>
      <StatusBadge status="ok" label="Environment configured" />
    </Box>
  );
}
