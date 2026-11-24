# Moving the AdminConsole workspace from Cowork to Claude Code

Written 2026-09-29. Do steps 1-3 before October 6.

## 1. Copy the workspace somewhere permanent

The whole workspace lives in a Cowork session folder on this Mac:

    /Users/glenn.armbruster/Library/Application Support/Claude/local-agent-mode-sessions/42575061-9c3b-4220-af4f-d2ed7c5e1703/99f73e5b-1573-466b-b875-c4b871ec0ec0/local_2ec1a5a4-b4ad-4b56-be11-cbe39e5eb759/outputs

1. Finder > Go > Go to Folder (Shift+Cmd+G), paste the path above, press Return.
2. Select everything inside (Cmd+A) and copy it to a new folder you own, e.g.
   `~/Documents/AdminConsole-Workspace` (or an iCloud Drive folder, so it is also backed up).
3. Keep the zip `AdminConsole-Workspace-2026-09-29.zip` from the same folder as a second,
   frozen copy.
4. Check the copy: it should contain `CLAUDE.md`, `AdminConsole/`, `ScriptNotes/`, `shipEX/`,
   `shipFH/`, `shipGC/`, `undo/` and `Handoff/`.

## 2. Save the standards document into the new folder

Your "PowerShell WinForms Scripting Standards" document lives in the Claude.ai project
(Project knowledge, file `ex - script project.md`). Download it and save it as:

    ~/Documents/AdminConsole-Workspace/Handoff/PowerShell-WinForms-Scripting-Standards.md

CLAUDE.md points Claude Code at that path.

## 3. Ship anything still waiting

Check `shipEX/`, `shipFH/`, `shipGC/` in the new copy and ship what is there as usual.

## 4. Open the folder in Claude Code

- Open Claude Code and point it at `~/Documents/AdminConsole-Workspace`
  (desktop app: open/choose that folder; terminal: `cd ~/Documents/AdminConsole-Workspace` then `claude`).
- Claude Code reads `CLAUDE.md` in that folder automatically at the start of every session.

## 5. Reconnect Microsoft Learn

In a terminal, inside the workspace folder:

    claude mcp add --transport http microsoft-learn https://learn.microsoft.com/api/mcp

(Or, inside Claude Code: `/plugin marketplace add microsoftdocs/mcp` then
`/plugin install microsoft-docs@microsoft-docs-marketplace`, and restart.)

## 6. First message in Claude Code

    Read CLAUDE.md, Handoff/PROJECT-MEMORY.md and ScriptNotes/ARCHITECTURE.md sections 99-108,
    then summarise the standing rules back to me before we change anything.

If the summary misses a rule, correct it and ask Claude to add it to CLAUDE.md.

## 7. Optional: real PowerShell syntax checks on the Mac

    brew install powershell

Then Claude Code can parse every script with the real PowerShell parser before staging,
instead of counting brackets.

## What does NOT carry over

- This Cowork conversation and its task list. The decisions are in ARCHITECTURE.md and CLAUDE.md.
- Claude.ai project memory. The work-related parts are exported to Handoff/PROJECT-MEMORY.md.
- Cowork-only features (scheduled tasks, artifacts). None are used by the toolkit.
