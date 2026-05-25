import React, { useState } from "react";
import { Box } from "ink";
import { Header } from "./components/Header.js";
import { Onboarding } from "./screens/Onboarding.js";
import { Auth } from "./screens/Auth.js";
import { MainMenu } from "./screens/MainMenu.js";
import { NotebookList } from "./screens/NotebookList.js";
import { SyncStatus } from "./screens/SyncStatus.js";
import { DiffResync } from "./screens/DiffResync.js";
import { ExportScreen } from "./screens/ExportScreen.js";
import { PostprocessScreen } from "./screens/PostprocessScreen.js";
import { StatsServerScreen } from "./screens/StatsServerScreen.js";
import { MarkdownExportScreen } from "./screens/MarkdownExportScreen.js";
import { HelpScreen } from "./screens/HelpScreen.js";
import type { Screen } from "./types.js";

export function App() {
  const [screen, setScreen] = useState<Screen>("onboarding");
  const [activeNotebook, setActiveNotebook] = useState<string | undefined>(undefined);

  const navigate = (s: Screen) => setScreen(s);
  const goBack = () => {
    if (screen === "main-menu") return;
    setScreen("main-menu");
  };

  return (
    <Box flexDirection="column" padding={1}>
      <Header activeNotebook={activeNotebook} />
      {screen === "onboarding" && <Onboarding onComplete={() => navigate("auth")} />}
      {screen === "auth" && <Auth onComplete={() => navigate("main-menu")} />}
      {screen === "main-menu" && (
        <MainMenu
          onSelect={navigate}
          onSelectNotebook={setActiveNotebook}
          activeNotebook={activeNotebook}
        />
      )}
      {screen === "notebook-list" && (
        <NotebookList
          onSelect={(name) => {
            setActiveNotebook(name);
            navigate("main-menu");
          }}
          onBack={goBack}
        />
      )}
      {screen === "sync-status" && <SyncStatus onBack={goBack} />}
      {screen === "diff-resync" && <DiffResync notebook={activeNotebook || "A"} onBack={goBack} />}
      {screen === "export" && <ExportScreen notebook={activeNotebook || "A"} onBack={goBack} />}
      {screen === "postprocess" && <PostprocessScreen notebook={activeNotebook || "A"} onBack={goBack} />}
      {screen === "stats-server" && <StatsServerScreen onBack={goBack} />}
      {screen === "markdown-export" && <MarkdownExportScreen onBack={goBack} />}
      {screen === "help" && <HelpScreen onBack={goBack} />}
    </Box>
  );
}
