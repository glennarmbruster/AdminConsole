# DEVIN BUILD PACK — M365 / Exchange hybrid admin toolkit

**THIS FILE IS TEN SEPARATE DOCUMENTS CONCATENATED.**

Treat each numbered section below as a standalone document with its own
authority and its own audience. They are combined into one file only because of
a transfer constraint, not because they are one argument. Cross-references
between them use the original filenames, which are given in each heading.

---

# ▶ STEP 0 — DO THIS BEFORE READING ANYTHING ELSE

You have been given **three files** in a directory:

| File | What it is |
|---|---|
| `DEVIN-PACK.md` | This file. Ten instruction documents. |
| `DEVIN-REFERENCE-CODE.md` | Sixteen working files: the shared layer, the verifier, the console, core tools and config. |
| `DEVIN-REFERENCE-CODE-2.md` | Eight more working tools, one for each remaining tool in the first-pass scope. |

> **Scope, in one line.** `DOCUMENT 1` §5 is the **first pass** — eight tools.
> `DOCUMENT 8` is the **whole toolkit** — 55 buttons over 65 files, in eight
> further stages. Build the first pass, stop for review, then work the stages.
> Where an early document says something is "out of scope", `DOCUMENT 8` §0 says
> whether that means *deferred* or *never*.

**Your first action is to extract the two reference-code files into real files
on disk.** Every instruction in this pack refers to those scripts by path, and
those paths only resolve after you have done this.

Extract them into a `reference/` folder at the repo root, preserving the exact
names in each section heading:

```
reference/
    Common.ps1              Test-Toolkit.ps1        AdminConsole.ps1
    ToolManager.ps1         Connect-ExO.ps1         Connect-ExSE.ps1
    MailboxHealth.ps1       MessageTrace.ps1        Unified-Lookup.ps1
    Group-Manager.ps1       Group-Manager-Editor.ps1  activate.ps1
    MailboxPermissions.ps1  CalendarPermissions.ps1 OOFEditor.ps1
    InboxRules.ps1          MailAttributes.ps1      Protocols.ps1
    ExternalForwarding.ps1
    Config/
        Tenant.TEMPLATE.json    Settings.json
        Sections.json           Tools.json
        Themes.TEMPLATE.json
```

Each section in those files is delimited by a heading of the form
`# ◆ FILE n OF m — \`Name.ps1\`` followed by the content in a fenced code
block. Extract the block, not the heading or the commentary above it.

**Confirm the extraction before proceeding:** report how many files you wrote
and their line counts. The counts are stated in each section heading, so a
mismatch means the extraction is wrong — and a subtly truncated `Common.ps1`
would poison everything built on top of it.

Then read `DOCUMENT 1` below and follow it.

**Do not write any toolkit code until Step 0 is confirmed and you have read
Documents 1 and 2 in full.**

---

