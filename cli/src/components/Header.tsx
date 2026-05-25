import React from "react";
import { Box, Text } from "ink";

export type HeaderProps = {
  activeNotebook?: string;
};

export function Header({ activeNotebook }: HeaderProps) {
  return (
    <Box flexDirection="column" marginBottom={1}>
      <Box>
        <Text bold color="cyan">
          📓 OneNote Interactor
        </Text>
        <Text color="gray"> — Interactive CLI</Text>
      </Box>
      {activeNotebook && (
        <Box>
          <Text color="gray">Active notebook: </Text>
          <Text bold color="yellow">
            {activeNotebook}
          </Text>
        </Box>
      )}
    </Box>
  );
}
