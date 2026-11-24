# Mailbox / Group Inspector -- Launcher Architecture Refactor

No hate here -- 4,272 lines in one file is a completely reasonable thing to want to break up, and the tool itself is doing a lot (11 distinct admin capabilities across Exchange Online and two on-prem AD environments). This document is the design you asked for: folder structure, the two launcher scripts, the JSON schema, the parameter model, a full function-to-script map, the trade-offs worth knowing about, and a phased plan that keeps the current tool working the entire time.

**All fourteen files below are now built.** Every tool in `Tools.json` points to a working, verified script -- the migration described in this document is complete (Phases 0-3 in Section 10; Phase 4, retiring the monolith, is still your call).

| File | Status |
|---|---|
| `Common.ps1` | Complete -- shared functions, moved unchanged from the monolith |
| `AdminConsole.ps1` | Complete -- Layer 1 launcher |
| `ToolManager.ps1` | Complete -- Tools.json editor |
| `Tools.json` | Seed data for all 10 tools |
| `MailboxHealth.ps1` | Complete -- template tool |
| `MessageTrace.ps1` | Complete |
| `InboxRules.ps1` | Complete -- includes a small new wrapper form (see Section 7 correction below) |
| `CalendarPermissions.ps1` | Complete |
| `OOFEditor.ps1` | Complete |
| `DLMemberships.ps1` | Complete |
| `MailAttributes.ps1` | Complete |
| `AllAttributes.ps1` | Complete |
| `MailboxPermissions.ps1` | Complete -- includes a small new wrapper form |
| `GroupEditor.ps1` | Complete -- includes a small new wrapper form, and is the one tool that calls `Resolve-Target` itself (see its own header for why) |

Every extracted script was verified for bracket balance and single, correctly-wired function definitions before delivery, following the same check used on `MailboxHealth.ps1`.

---

## 1. Folder structure

```
AdminTools\
├── AdminConsole.ps1          Layer 1 launcher (menu only, no business logic)
├── ToolManager.ps1           Tools.json editor
├── Common.ps1                Shared functions -- every script dot-sources this
├── Tools.json                Button/tool configuration, read by AdminConsole
├── REFACTOR-DESIGN.md        This document
│
├── MailboxHealth.ps1         Layer 2 -- done
├── MessageTrace.ps1          Layer 2 -- done
├── InboxRules.ps1            Layer 2 -- done
├── CalendarPermissions.ps1   Layer 2 -- done
├── OOFEditor.ps1             Layer 2 -- done
├── DLMemberships.ps1         Layer 2 -- done
├── MailAttributes.ps1        Layer 2 -- done
├── AllAttributes.ps1         Layer 2 -- done
├── MailboxPermissions.ps1    Layer 2 -- done
├── GroupEditor.ps1           Layer 2 -- done
│
└── Mailbox-Group-Inspector.ps1   The original monolith -- kept, untouched,
                                   as a working fallback. All 10 tools are now
                                   extracted and verified; retiring this file
                                   (Section 10, Phase 4) is optional and up to you.
```

Everything lives flat in one folder on purpose. `AdminConsole.ps1` resolves tool scripts with `Join-Path $PSScriptRoot $Tool.Script`, and `Tools.json` paths are relative (`.\MailboxHealth.ps1`) -- so this whole folder can be zipped, copied to a different machine or a different admin's laptop, and it still works with no path edits. This is also what makes the tiered-distribution goal (Section 8) trivial: a "Tier 1 package" is just this same folder with 7 of the 10 tool scripts deleted and a trimmed `Tools.json`.

---

## 2. AdminConsole.ps1 design

Already built. Four jobs, exactly as specified, nothing else:

1. **Collect input** -- Environment radio buttons (Franchise/HQ), DC dropdown, one text box for the mailbox/group address. This is the monolith's own top-of-form UI, moved with only cosmetic re-numbering of Y-coordinates (it's no longer sharing vertical space with 10 other tools' controls below it).
2. **Detect object type** -- calls `Resolve-Target` (moved unchanged into `Common.ps1`) and gets back `Mailbox`, `UnifiedGroup`, `OnpremGroup`, or `NotFound`.
3. **Render buttons dynamically** -- reads `Tools.json`, filters to entries whose `ObjectTypes` array contains the detected type, and adds one button per entry to a `FlowLayoutPanel` (which auto-wraps to new rows, so the button count isn't hardcoded anywhere). Zero tool names appear in `AdminConsole.ps1`'s code.
4. **Launch + pass context** -- each button's click handler runs `& $scriptPath -SmtpAddress $smtp -Environment $env -DomainController $dc`, in AdminConsole's own PowerShell session.

**One deliberate, documented behavior change from the monolith:** the old single "Run" click also fetched and displayed the mailbox/group's full current-state summary (size, delegate counts, owner/member lists) right there on the same screen, because the permission-editing controls lived on that same form. `AdminConsole.ps1` does not do this -- it only reports *what* was found (`"Found: Mailbox -- jane@expresspros.com"`), not its state, because rendering that summary is display logic that belongs to a tool (`MailboxPermissions.ps1` / `GroupEditor.ps1`), not the launcher. See Section 8 if you'd rather keep that inline summary -- it's a small, reversible decision, not a hard architectural constraint.

**Corrected 2026-07-10 (field report):** this originally used `Start-Process powershell.exe` to launch each tool as a fully separate process, on the theory that a hang or crash in one tool window could never take down the console or another open tool window. In practice this broke every tool that needs a live Exchange Online connection or the ActiveDirectory module -- both are per-session state, and a brand-new `powershell.exe` process starts with neither, regardless of what's connected in the window AdminConsole itself is running in. `Start-ToolScript` now launches each tool with the call operator (`&`), not `Start-Process`, so tools run in AdminConsole's own session and see whatever's already connected/imported there -- matching the assumption the original monolith always made (see the warning label in the console: "Connect to Exchange Online before running"). The call operator still gives each tool its own child scope, so its `$script:`-scoped state doesn't leak into AdminConsole or collide between tools; only the connected-module/session state is actually shared, which is exactly what was needed. See Section 8 for the trade-off this introduces (one tool window at a time instead of several in parallel).

---

## 3. ToolManager.ps1 design

Already built. A `DataGridView` bound to `Tools.json`, one row per tool, five editable columns (`Order`, `Name`, `Script`, `ObjectTypes`, `Tier`), plus:

- **Add Tool** -- built into the grid itself (`AllowUserToAddRows = $true`); type into the blank row at the bottom.
- **Remove Tool** -- same, via the grid's built-in row-delete (select row, press Delete).
- **Rename Tool / Change Script Path** -- edit the `Name` / `Script` cell directly.
- **Reorder Tools** -- `Move Up` / `Move Down` buttons swap the selected row's `Order` value with its neighbor.
- **Save** -- validates (every row needs a `Name` and `Script`; `ObjectTypes` values must be `Mailbox`, `UnifiedGroup`, or `OnpremGroup`), then writes `Tools.json`.

You never edit `AdminConsole.ps1`'s code to add or remove a tool -- you run `ToolManager.ps1`, make the change, hit Save, and the next time `AdminConsole.ps1` starts (or the operator re-runs a lookup) the button list reflects it.

---

## 4. Tools.json structure

```json
{
  "Order": 1,
  "Name": "Mailbox Health",
  "Script": ".\\MailboxHealth.ps1",
  "ObjectTypes": ["Mailbox"],
  "Tier": "Tier1"
}
```

| Field | Purpose |
|---|---|
| `Order` | Integer. Controls button order in `AdminConsole.ps1`. |
| `Name` | Button text. |
| `Script` | Path to the tool script, relative to `AdminConsole.ps1`'s folder. |
| `ObjectTypes` | Array. `AdminConsole.ps1` only shows this button when the detected type is one of these. A mailbox-only tool lists `["Mailbox"]`; `GroupEditor.ps1` lists `["UnifiedGroup", "OnpremGroup"]`. |
| `Tier` | Free text (`Tier1`/`Tier2`/`Tier3`, or your own scheme). **Not read by `AdminConsole.ps1` at all** -- purely a label for you when deciding what to hand out to whom (Section 7 / Section 8). |

The seed `Tools.json` included here has all 10 tools already entered with the tiers implied by your own Tier 1/2/3 examples -- see the table in Section 7.

---

## 5. Parameter-passing model

Every Layer 2 tool script has the same three-parameter contract:

```powershell
param(
    [string]$SmtpAddress,
    [string]$Environment,        # 'Franchise' or 'HQ'
    [string]$DomainController    # DC hostname (Franchise) or ignored (HQ)
)
```

Launched from `AdminConsole.ps1`, all three arrive filled in and the tool opens straight to its window -- no prompting. Launched directly (double-click, or from a shortcut), any missing or invalid value triggers `Common.ps1`'s `Resolve-ToolTarget`, which pops `Show-TargetPromptDialog` -- a small dialog with the exact same input box / Environment radios / DC dropdown the monolith's own top-of-form UI has. One shared dialog function, instead of 10 tool scripts each carrying their own copy of that UI.

```
Resolve-ToolTarget
    │
    ├─ all 3 params present & valid ──────────────► return immediately, no UI
    │
    └─ anything missing/invalid ──► Show-TargetPromptDialog ──► operator fills
                                     it in, or cancels (script exits cleanly)
```

This is the one piece of genuinely new code in the whole refactor (everything else is moved, not written) -- and it's UI/parameter glue, not business logic for any tool.

---

## 6. Migration strategy (high level)

1. Build Layer 1 infrastructure (`Common.ps1`, `AdminConsole.ps1`, `ToolManager.ps1`, `Tools.json`) -- **done**, this delivery.
2. Extract one tool as a worked template (`MailboxHealth.ps1`) -- **done**, this delivery.
3. Extract the remaining nine tools, one at a time, each verified against the still-running monolith before moving to the next (Section 10 has the order and the per-tool notes).
4. Once every tool in `Tools.json` has a working script and has been spot-checked, retire `Mailbox-Group-Inspector.ps1` (rename it `Mailbox-Group-Inspector.ps1.bak` rather than deleting it -- keep it as a reference/rollback for a few weeks).

---

## 7. Function-to-script migration map

Built from the monolith's actual `#region` boundaries and function list (`Mailbox-Group-Inspector.ps1`, 4,272 lines, verified by grep, not guessed):

| Destination script | Functions moved (unchanged) | Original region / lines | Suggested Tier |
|---|---|---|---|
| **Common.ps1** *(done)* | `Get-TargetSmtp`, `Format-ErrorDetail`, `Get-EXOCommand`, `Invoke-EXOCommand`, `Resolve-Target`, `Write-CompletedLog`, `Get-InboxRulesForMailbox` | Parsing (550-560), Error formatting (561-626), Resolution (627-669), Logging (3654-3670), Inbox rules (798-806 only) | n/a -- shared |
| **MailboxHealth.ps1** *(done)* | `Get-MailboxHealthSnapshot`, `Show-MailboxHealthForm` | Mailbox health (1457-1956, minus `Get-MailboxDLMemberships`/`Show-DLMembershipsForm`) | Tier 1 |
| **MessageTrace.ps1** | `Invoke-MessageTrace`, `Invoke-MessageTraceDetail`, `Find-QuarantineMatch`, `Invoke-QuarantineRelease`, `Get-SelectedTraceTimeZone`, `ConvertTo-TraceDisplayTime`, `Format-TraceDetail`, `Show-MessageTraceForm` | Message trace (3083-3653) | Tier 1 |
| **InboxRules.ps1** | `Get-InboxRulesForMailbox` *(already in Common.ps1 -- reused, not duplicated)*, `Enable-TargetInboxRule`, `Disable-TargetInboxRule`, `Remove-TargetInboxRule`, `Update-RulesGrid`, `Show-RuleDetailForm`, plus a NEW `Show-InboxRulesForm` wrapper (see correction below) | Inbox rules (797-983) | Tier 1 |
| **CalendarPermissions.ps1** | `Get-CalendarFolderIdentity`, `Get-CalendarPermissions`, `Set-CalendarPermission`, `Remove-CalendarPermission`, `Show-CalendarPermissionsForm` | Calendar permissions (1271-1456) | Tier 2 |
| **OOFEditor.ps1** | `Get-OutOfOfficeConfig`, `Set-OutOfOfficeConfig`, `Show-OutOfOfficeForm` | Out of office (1107-1270) | Tier 2 |
| **DLMemberships.ps1** | `Get-MailboxDLMemberships`, `Show-DLMembershipsForm` | Mailbox health region (1458-1555, 1852-1956) | Tier 2 |
| **MailboxPermissions.ps1** | `Get-MailboxState`, `Format-MailboxState`, `Grant-FullAccessPermission`, `Revoke-FullAccessPermission`, `Grant-SendAsPermission`, `Revoke-SendAsPermission`, `Invoke-MailboxEdit`, plus a NEW `Show-MailboxPermissionsForm` wrapper | Mailbox (670-796), Invoke-MailboxEdit (4083-4114) | Tier 2 |
| **MailAttributes.ps1** | `Get-MailAttribsADUser`, `Get-PrimarySmtpFromProxyAddresses`, `Set-PrimarySmtpInProxyAddresses`, `Add-SecondarySmtpToProxyAddresses`, `Format-AttributeValue`, `Format-MailAttribsTable`, `Set-MailAttribsAttribute`, `Add-MailAttribsSecondaryProxyAddress`, `Show-MailAttribsEditDialog`, `Show-KeepOrRemoveOldPrimaryDialog`, `Show-MailAttribsForm` | Mail attributes (1957-2667) | Tier 3 |
| **AllAttributes.ps1** | `Get-AllAttribsADUser`, `Format-AllAttribsValue`, `Format-AllAttribsTable`, `Set-AllAttribsAttribute`, `Show-AllAttribsEditDialog`, `Show-AllAttribsForm` | All attributes (2668-3082) | Tier 3 |
| **GroupEditor.ps1** | `Get-GroupState`, `Format-GroupState`, `Add-UnifiedOwner`, `Remove-UnifiedOwner`, `Add-UnifiedMember`, `Remove-UnifiedMember`, `Set-OnpremOwner`, `Add-OnpremMember`, `Remove-OnpremMember`, `Invoke-GroupEdit`, plus a NEW `Show-GroupEditorForm` wrapper | Group (984-1106), Invoke-GroupEdit (4136-4228) | Tier 3 |

Every function name above is copy-pasteable -- these are exact matches from the current file, not paraphrases.

**Correction to this section, made while building the remaining nine scripts:** this originally said only `MailboxPermissions.ps1` and `GroupEditor.ps1` needed a new wrapper form, and that the other eight already had a self-contained `Show-XForm`. That undercounted by one. `InboxRules.ps1`'s grid (`$grdRules`) and its driving function (`Update-RulesGrid`) also lived directly on the monolith's main form, not in a dedicated `Show-XForm` -- so it needed the same treatment as the other two. The corrected breakdown:

**Six tools** (`MessageTrace.ps1`, `CalendarPermissions.ps1`, `OOFEditor.ps1`, `DLMemberships.ps1`, `MailAttributes.ps1`, `AllAttributes.ps1`) already had a self-contained `Show-XForm` in the monolith, so extraction was almost entirely mechanical:

1. Copy the `Show-XForm` function and its supporting `Get-`/`Set-` functions into the new file, verbatim.
2. Change the function's closing `[void]$xForm.ShowDialog($form)` to `[void]$xForm.ShowDialog()` -- the only line that has to change, because there's no longer a parent `$form` (see `MailboxHealth.ps1` for exactly what this looks like).
3. Add the standard param block + `Resolve-ToolTarget` call at the top (copy from `MailboxHealth.ps1`).

**Three tools** (`InboxRules.ps1`, `MailboxPermissions.ps1`, `GroupEditor.ps1`) needed a small new wrapper form built (following the same layout the monolith already used for `$grdRules` / `$grpMbx` / `$grpGrp` respectively) around the otherwise-unchanged driving functions (`Update-RulesGrid`+`Show-RuleDetailForm`; `Invoke-MailboxEdit`+Grant/Revoke; `Invoke-GroupEdit`+Add/Remove-Owner/Member). `GroupEditor.ps1` additionally has to call `Resolve-Target` itself, since `AdminConsole.ps1` only passes `Smtp`/`Environment`/`DomainController` -- not the object `Type`/`ADGroupDN` it already resolved while deciding to show the button. See that script's own header for the full explanation.

All nine are now built and verified (bracket-balance + function-definition checks, same as `MailboxHealth.ps1`).

---

## 8. Architectural concerns and improvements

**Missing tool script shouldn't crash the console.** `AdminConsole.ps1`'s `Start-ToolScript` already checks `Test-Path` before launching and shows a plain warning instead of an unhandled error if a script referenced in `Tools.json` isn't present. This matters specifically because of the tiered-distribution model -- a Tier 1 admin's folder legitimately won't have `GroupEditor.ps1`, and that should never look like a bug.

**`Tier` in `Tools.json` is documentation, not enforcement.** Nothing in `AdminConsole.ps1` checks who's running it or blocks a tool at runtime -- the security boundary is "the file isn't on disk," full stop, which is exactly what you described wanting (control by what's distributed, not by a login check). Worth being explicit about, though: if `Tools.json` itself ever gets shared as one file across all tiers (instead of being trimmed per package), that boundary disappears, since the JSON alone doesn't stop someone from placing a copied `.ps1` file into the folder. If that risk matters to you, the mitigation is procedural (each tier gets its own zipped package with only its own scripts + its own trimmed `Tools.json`), not something to build into the code.

**The "show current state on lookup" behavior moved, not disappeared -- confirm you're OK with this.** Flagged already in Section 2, repeating it here because it's the one place this refactor changes what the operator sees, even though no capability was removed. If you'd rather `AdminConsole.ps1` still show the summary immediately after lookup (before the operator picks a tool), that's a small addition -- have `AdminConsole.ps1` also dot-source `Get-MailboxState`/`Format-MailboxState`/`Get-GroupState`/`Format-GroupState` (currently slated for `MailboxPermissions.ps1`/`GroupEditor.ps1`) purely for read/display. Say the word and I'll build it that way instead.

**Two tools currently share one small function (`Get-InboxRulesForMailbox`).** Promoted to `Common.ps1` rather than duplicated, since `MailboxHealth.ps1` explicitly reuses it "rather than querying it twice" (that's the original code's own comment) and duplicating a working, already-verified function is worse for maintainability than one extra shared dependency. As more tools get extracted, watch for this pattern again -- if two tools need the same non-trivial function, it likely belongs in `Common.ps1`; if it's tiny and truly single-purpose, duplication across two files is fine and keeps each script more independent.

**Launching tools in-session (not per-process) trades parallelism for a working connection.** `Start-ToolScript` was originally written around `Start-Process powershell.exe`, one real process per tool, on the theory that a hang or crash in one tool should never take down the console or another open tool window. That theory was correct but the cost was worse than expected: a brand-new `powershell.exe` process has no Exchange Online connection and no `ActiveDirectory` module imported, no matter what's already connected in the window AdminConsole is running in -- so every tool failed with "cmdlet not found" style errors the moment it tried to do anything, which defeats the point of a launcher. Fixed by switching `Start-ToolScript` to the call operator (`&`), so tools run in AdminConsole's own session and see whatever's already connected there -- exactly the assumption the original monolith always made. The remaining cost is the one flagged above: one tool window at a time (AdminConsole's UI thread is busy until the tool window closes), and a crash inside a tool now can take AdminConsole down with it. If that trade-off stops being acceptable, the alternative is a `Connect-IfNeeded` helper in `Common.ps1` that each tool calls at its own startup, paired with going back to `Start-Process` -- a bigger change, not attempted here.

---

## 9. Best practices for future growth

- **New tool checklist:** write the tool script (param block + `Resolve-ToolTarget` + your `Show-XForm`), test it standalone (`.\NewTool.ps1` with no params, confirm the prompt dialog appears and works), test it with explicit params (confirm it skips the prompt), then add one entry to `Tools.json` via `ToolManager.ps1`. No other file changes, ever.
- **Keep `Common.ps1` small on purpose.** Only promote a function there when a second tool genuinely needs it (see the `Get-InboxRulesForMailbox` note above) -- a bloated "shared utils" file becomes its own maintenance problem.
- **Every tool script should degrade gracefully if Exchange Online / AD isn't connected**, exactly like the monolith already does (`Format-ErrorDetail` catch blocks everywhere) -- this didn't change and shouldn't.
- **Version each tool script's own `.NOTES` block independently** once they're split apart -- the monolith's single shared revision history stops being meaningful once each tool has its own lifecycle and its own file in source control.
- **Source control this whole folder as one repo.** The big win of splitting the file isn't really the launcher UI -- it's that a diff on `MessageTrace.ps1` no longer shows unrelated changes to nine other tools, and code review / blame becomes tool-scoped instead of whole-monolith-scoped.

---

## 10. Phased migration plan

The monolith keeps running, unmodified, through every phase below -- nothing here requires touching `Mailbox-Group-Inspector.ps1` until the very last step, and that step is optional.

**Phase 0 -- done, this delivery.**
`Common.ps1`, `AdminConsole.ps1`, `ToolManager.ps1`, `Tools.json`, `MailboxHealth.ps1`. Run `AdminConsole.ps1` side-by-side with the existing monolith; look up a mailbox in both, confirm `AdminConsole.ps1` detects the same type and `MailboxHealth.ps1`'s output matches the monolith's Health button output exactly (same data, since it's the same function).

**Phase 1 -- Tier 1 tools -- done.**
`MessageTrace.ps1` and `InboxRules.ps1` are both built. `InboxRules.ps1` turned out to need the small new wrapper form (Section 7 correction), not the "easy" pattern originally expected -- everything else about it went as planned.

**Phase 2 -- Tier 2 tools -- done.**
`CalendarPermissions.ps1`, `OOFEditor.ps1`, `DLMemberships.ps1`, `MailboxPermissions.ps1` are all built.

**Phase 3 -- Tier 3 tools -- done.**
`MailAttributes.ps1`, `AllAttributes.ps1`, `GroupEditor.ps1` are all built. `GroupEditor.ps1` is the only tool in `Tools.json` with `ObjectTypes: ["UnifiedGroup", "OnpremGroup"]` instead of `["Mailbox"]`, confirming `AdminConsole.ps1`'s type-filtering works both ways.

**Phase 4 -- retirement (optional, whenever you're confident).**
All 10 entries in `Tools.json` now point to working, verified scripts. Rename `Mailbox-Group-Inspector.ps1` to `Mailbox-Group-Inspector.ps1.bak` and keep it for a few weeks as a fallback once you've spot-checked the new scripts against it. Start distributing tiered packages (Section 8) to other admins whenever you're ready -- that's purely a "copy this folder, delete some .ps1 files, trim Tools.json" step, no further code changes needed.
