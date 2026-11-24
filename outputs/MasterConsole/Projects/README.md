# Projects

A **tool** lives at the toolkit root, answers one question, and is used
indefinitely. A **project** lives here: a numbered sequence of steps run in
order against a batch of objects, with a beginning, an end, and its own working
data.

```
AdminConsole\
    Common.ps1
    AdminConsole.ps1
    ...tools...
    Config\
    Projects\
        DLtoSharedMbx\       <- one folder per project
            1-ONPREM-Export.ps1
            2-ONPREM-Decomm.ps1
            3-CLOUD-CheckSync.ps1
            4-CLOUD-create.ps1
            DL-Input.csv
            Exports-Complete\
        DLtoCloudDL\         <- next one
```

Full reasoning is in `ScriptNotes/ARCHITECTURE.md` §24. The short version:

---

## Starting a new project

**1. Make the folder.** `Projects\<ProjectName>\`

**2. Number the steps** in the order they must run — `1-`, `2-`, `3-`. The
numbering is the documentation of the sequence, and it is the first thing
someone reads.

**3. Start each script with this:**

```powershell
#Requires -Version 5.1
# TOOLKIT-HELP-BEGIN
# STEP n of <project>. <what it does>
#
# <what is irreversible, if anything>
# <what must be true before the next step runs>
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Two levels up: Projects\<Name> -> AdminConsole
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'Common.ps1')

# The script lives in the folder it works in.
$basePath = $PSScriptRoot
```

**4. Never hard-code a path.** `$basePath = $PSScriptRoot` is the whole rule.
Every DLConversion script had an absolute `E:\scripts\...` path, and one of the
four had it spelled wrong — so that step could never find its input. A constant
that only works on one machine is a constant nobody checks.

**5. Add a `Section` to `Tools.json` named for the project**, so the console
grouping and the folder share a name.

**6. Decide `OwnProcess` per step, not per project.** Steps often target
different directories:

| | targets | `TOOLKIT-REQUIRES: OwnProcess` |
|---|---|---|
| on-premises step | Exchange on-prem / AD | **yes** — must not share a process with the cloud |
| cloud step | Exchange Online / SPO | **no** — needs the console's live session |

Getting this backwards breaks things quietly. A cloud step forced into its own
process gets an on-premises connection and no cloud one.

**7. Qualify ambiguous cmdlets.** `Get-DistributionGroup`, `Set-Mailbox`,
`Get-Recipient` and friends exist in *both* Exchanges. In a cloud step, route
them through `Invoke-EXOCommand`; in an on-premises step, `Invoke-OnPremCommand`
or `OwnProcess`. A project that touches both worlds is where this goes wrong.

---

## Working data

Code and data share the folder — that is what makes `$PSScriptRoot` work, and
it means **shipping the project ships the data.**

- Completed output goes in a subfolder (`Exports-Complete\` is the pattern).
- **Empty it before shipping.**
- Input templates ship with headers and no rows.
- Anything naming real people — an owners list, a manager CSV — is data, and
  needs the same treatment.

---

## Naming

**Name the transformation, not the subject.**

| | |
|---|---|
| `DLtoSharedMbx` | DL becomes a shared mailbox + access group |
| `DLtoCloudDL` | on-premises DL becomes a cloud DL |

These were nearly called `DLConversion` and `DLMigrations`, which are
synonyms -- neither tells you which is which.

That matters because **they run at the same time, for years.** Two live
projects, two folders, two console sections, two sets of exports. The name has
to be unambiguous at a glance, on the destructive step, under time pressure.

Same test for the `Tools.json` section and for export filenames: someone
opening a CSV six months from now should know which project made it.

---

## Before shipping to another engineer

- No absolute paths anywhere (`grep` for `:\\`).
- Working-data subfolders emptied.
- Every step has a `TOOLKIT-HELP` block.
- Every step is UTF-8 **with BOM**, CRLF, ASCII-only.
- `Test-Toolkit.ps1` run over the toolkit, clean.

That last one matters most for adopted scripts. A script written elsewhere is
usually ANSI with no BOM, and reading it as UTF-8 silently destroys any
bullet or dash in its dialog text — see ARCHITECTURE §14i, which exists because
this has happened twice.
