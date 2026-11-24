# REFERENCE FILES — what ships, and how to use them

**Nineteen** working scripts and five config files from the original toolkit
ship in `reference/`, delivered across two concatenated files and extracted in
Step 0. They are **reference implementations, not a starting codebase.** That
distinction is the most important thing in this document and is spelled out
below.

- `DEVIN-REFERENCE-CODE.md` — 11 scripts + 5 config. Infrastructure and core
  tools. Described in the two tables immediately below.
- `DEVIN-REFERENCE-CODE-2.md` — 8 scripts. **An exemplar for every remaining
  tool in the build scope**, so that no tool has to be written from prose alone.
  Listed under "Part 2" further down.

---

## The core five

Take these first. If only five files were available, these are the five.

| File | ~lines | The reference for |
|---|---|---|
| `Common.ps1` | 7,200 | **The spine.** Config readers, connection probes, capability gating, audit log, error display, theming, geometry. Every convention in `02` is realised here. |
| `Test-Toolkit.ps1` | 1,950 | The verifier from `03`, already built. Run and extend it — do not reimplement it from the spec. |
| `AdminConsole.ps1` | 2,230 | Launcher: connection state, capability-driven greying, JSON-driven layout, the modal-loop deferral (B2). |
| `Unified-Lookup.ps1` | 2,760 | The most complex read tool. Two directories searched in one pass, results classified, routing to the right editor, a capped result set filtered. |
| `Group-Manager-Editor.ps1` | 1,250 | **The write reference.** Directory resolution before a write, identity binding, the on-prem/cloud refusal guard, audit rows on success and failure. |

## The recommended six

Added in value order. Cut from the bottom if the set has to shrink.

| # | File | ~lines | Why it earns a slot |
|---|---|---|---|
| 6 | `MailboxHealth.ps1` | 373 | **The canonical small tool.** All five above are large; none shows the whole shape at a glance. This one does, and it is tool #1 in the build order — an exact target rather than a description. |
| 7 | `Connect-ExO.ps1` | 251 | Encodes **B2** (a modal loop blocks any window a tool raises) and **C17** (connections are per-process). Two of the least intuitive constraints in the pack, in one of the smallest files. |
| 8 | `Connect-ExSE.ps1` | 203 | The on-premises half: session setup, `-ViewEntireForest` asserted rather than assumed (C6), and why it avoids `Get-Credential` entirely. |
| 9 | `ToolManager.ps1` | 948 | Config round-trip. Without it you will invent config editing, and **silently dropping fields the editor does not understand is a documented past bug** — it once deleted a live setting on first save. |
| 10 | `Protocols.ps1` | 404 | The tool that taught **C15** (slice a large query; never retry an oversized one) and **C13** (ask the cmdlet what it declares). The best single artefact on bulk queries. |
| 11 | `MessageTrace.ps1` | 688 | A clean mid-size tool: grid, filter-without-re-running, export, and an explicit statement of what the tool is *not* for. |

## Part 2 — one exemplar per remaining tool

These arrive in `DEVIN-REFERENCE-CODE-2.md`. They exist because after part 1,
five of the eight tools in the build scope had no working example, while three
did — and that asymmetry is exactly where conventions drift.

| File | ~lines | The reference for |
|---|---|---|
| `activate.ps1` | 326 | **Tool 0, built with the console in Phase 2.** Rights here are eligible, not standing, so nothing else works until it has been used. Read its ORDER MATTERS header. |
| `MailboxPermissions.ps1` | 347 | Tool 4. Full Access and Send As are separate grants and constantly confused. |
| `CalendarPermissions.ps1` | 246 | Tool 5. Folder permissions; the calendar folder name is localised. |
| `OOFEditor.ps1` | 240 | Tool 6. Two messages, internal and external, usually only one wrong. |
| `InboxRules.ps1` | 320 | Tool 7. Delete is irreversible; disable is not. |
| `Group-Manager.ps1` | 732 | Tool 8's **front door** — classifies, then routes to the editor in part 1. |
| `MailAttributes.ps1` | 739 | **Constraint C11 in real code**: the proxy-address rewrite that destroys X500 if done naively. The most damaging mistake available here. |
| `ExternalForwarding.ps1` | 405 | A complete CSV report: shared report controls, partial-run labelling, never a zero-row file. |

**Every tool in the first pass now has a working example.** If you want another
reference beyond these, ask first — the answer is usually that the pattern already
exists in one of these.

---

Also in `reference/Config/`: `Tenant.TEMPLATE.json`, `Settings.json`,
`Sections.json`, `Tools.json`, `Themes.TEMPLATE.json`. The config split is half the architecture and is
hard to infer from prose.

> `Tools.json` is a reference for **schema first, contents second**. It
> registers tools from the source toolkit, most of which are **deferred out of
> the first pass rather than out of scope** — they are specified in
> `09-FULL-TOOL-CATALOGUE.md` and are built in Stages A–H. For the first pass,
> copy the shape and write your own eight entries. From Stage A onward, `09` §2
> gives the destination sections and `09` §3 the `Group` field this file's copy
> does not exercise.

---

## How to use them

**These are the reference for CONVENTIONS, not a codebase to extend.**

The toolkit they come from served a different tenant with a different shape.
Copying a tool wholesale imports assumptions that stay invisible until they are
wrong — which is the failure mode this entire pack exists to prevent.

### Do

- **Match the conventions exactly.** Naming, parameter shapes, how a tool
  dot-sources `Common.ps1`, how it logs, how it reports partial results, how it
  lays out a form. Consistency across the toolkit is worth more than any
  individual improvement.
- **Take `Common.ps1` functions verbatim — but only the ones you need.** Their
  bodies encode constraints you cannot re-derive, so do not rewrite them from
  scratch. But the file is **7,200 lines and 164 functions**, and porting all of
  it would consume the build.
- **`08-COMMON-CONTRACT.md` is a FLOOR, not a list of everything you need.** It
  names the minimum. There are at least 33 further `Common.ps1` functions called
  by in-scope tools that it does not name. The discovery mechanism is `CMD001`:
  build, run it, and **every unresolved call is a function to take from the
  reference** rather than one to write yourself.
- **Run `Test-Toolkit.ps1` unchanged, first**, then extend it — starting with
  the `PS7002` rule from `03`, which the original does not have because it was
  written on the target platform.
- **Read the comments.** They are dense and load-bearing. Where a comment
  explains why something is done a surprising way, that explanation usually cost
  somebody a day. **Treat a comment that contradicts your instinct as
  information, not noise.**

### Do not

- **Do not copy a tool and adjust it.** Write each new tool from its spec in
  `05`, consulting the reference for *how* rather than *what*.
- **Do not carry over multi-directory machinery.** The source tenant had two AD
  domains and much of the directory-resolution code exists only to disambiguate
  them. This tenant has **one**. Keep the config-driven shape — `Directories`
  stays a list — but do not port cross-directory pinning, per-directory
  controller selection, or side-routing. Dead branches here.
- **Do not port anything you cannot explain.** If a block's purpose is not clear
  from the code and its comment, leave it out and note the omission. Unexplained
  code that is carried forward becomes permanent.
- **Do not assume the reference is bug-free.** It is working production code,
  not scripture. If something looks wrong, say so — several of the constraints
  in `01` exist because a previous version of these very files had the bug.

### Expect to find

Comments describing bugs that were fixed, sometimes at length, sometimes dated.
These are deliberate: they exist to stop the bug being reintroduced by the next
person to edit that line. Preserve the practice — short guardrail at the line,
longer reasoning in notes.

You will also find comments admitting a known limitation or an open question.
Those are honest, not stale. Do not "fix" them by deleting them.

---

## The one thing to do first

Run `Test-Toolkit.ps1` against `reference/` and read its output. In about a
minute that tells you which rules exist, what they are called, what a clean run
looks like, and what the report format is.

It is also the fastest way to learn what this toolkit considers a defect, which
is most of what `01` is trying to convey in prose.
