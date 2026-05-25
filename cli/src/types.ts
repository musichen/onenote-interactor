export type Screen =
  | "onboarding"
  | "auth"
  | "main-menu"
  | "notebook-list"
  | "sync-status"
  | "diff-resync"
  | "export"
  | "postprocess"
  | "stats-server"
  | "markdown-export"
  | "help";

export type AppState = {
  screen: Screen;
  activeNotebook?: string;
  previousScreen?: Screen;
};

export type CommandStatus = "idle" | "running" | "success" | "error";

export type CommandResult = {
  status: CommandStatus;
  output: string[];
  error: string[];
  exitCode: number | null;
};
