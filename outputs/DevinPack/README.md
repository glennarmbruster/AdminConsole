# What to send, and what to keep

This folder is a **working copy**. Only a few files leave it.

## If the first pass is already built — send TWO files

The eight tools exist. **Do not resend the reference code** — `reference/` is
untouched and Devin already has it.

| File | Size | What it is |
|---|---|---|
| `CONTINUE-PROMPT.txt` | ~7 KB | Paste into the existing session. Retrofit list, then Stage A. |
| `DEVIN-PACK.md` | ~229 KB | **Replaces** the copy in the repo. Now ten documents. |

The continuation has a deliberate shape: **Devin reports before it edits.** Five
questions, then the retrofit, then Stage A. Three of the corrections apply to
code already written, and one of them — a private fork of a `Common.ps1`
function — is the kind of thing that gets worse the longer it sits.

## For a cold start — send four

| File | Size | What it is |
|---|---|---|
| `KICKOFF-PROMPT.txt` | ~4 KB | Paste this into Devin. Nothing else needs typing. |
| `DEVIN-PACK.md` | ~229 KB | Ten instruction documents, concatenated. |
| `DEVIN-REFERENCE-CODE.md` | ~861 KB | Sixteen working files: shared layer, verifier, console, core tools, config. |
| `DEVIN-REFERENCE-CODE-2.md` | ~164 KB | Eight more: one exemplar per tool in the first-pass scope. |

`.txt` twins exist for all three `.md` files if a mail gateway strips markdown.
Identical content — send whichever survives.

**Do not send anything else.** The numbered source documents, `build-pack.py`,
and the `reference/` and `reference2/` folders are what the shipping files are
built *from*. Including them gives Devin two copies of the same material with no
way to know which is current.

## Scope, in one line

**`00` §5 is the first pass — eight tools. `09` is the whole toolkit — 55
buttons over 65 files, in eight further stages.** Devin builds the first pass,
stops for review, then works the stages one at a time.

Where an early document says something is "out of scope", `09` §0 says whether
that means *deferred to a stage* or *never*. That table exists because `00`,
`04` and `08` were written when eight tools was the whole plan.

## The workflow

1. Put the three content files in the Devin repo.
2. Paste `KICKOFF-PROMPT.txt`.
3. Devin extracts the reference code to `reference/`, reports line counts, and
   stops.
4. It reads the constraints and explains the platform mismatch back to you.
5. It gets the verifier running and proves a rule can fail.
6. Console plus **one tool**, then it stops again for you to run on Windows.
7. Then Stages A–H, one at a time, each ending with a report and a stop.

Steps 3–5 take minutes and are worth not skipping. If Devin cannot explain why
building on PowerShell 7 for a 5.1 target matters, it has not absorbed the thing
most likely to sink the delivery.

## When a build arrives

Work through `07-ACCEPTANCE-CHECKLIST.md` — it is inside `DEVIN-PACK.md` as
Document 10, and it is written for **you**, not the agent.

- **First pass:** eight gates, about thirty minutes, ordered so the cheapest
  checks that can invalidate everything come first. Gate 1 is two minutes and
  can end the review on its own.
- **Each stage after that:** the short list at the end of `07` — Gates 1 and 2
  again, plus the checks that only fail at scale (console height, section
  capacity, the group chooser).

## Regenerating

**If you edit any numbered source document, run `python3 build-pack.py`** before
sending. It rebuilds `DEVIN-PACK.md` and its `.txt` twin from the sources and
prints a line count and MD5. The flat file and the sources disagreeing has
happened during this build more than once, and it is invisible until Devin acts
on the stale copy.

`_PREAMBLE.md` holds the Step 0 block that sits above the concatenated
documents. Document order and the one-line descriptions live in `build-pack.py`.

## What this pack can and cannot do

It carries roughly forty findings that each cost real production time, an
exemplar for every tool in the first pass, a full specification for the other
forty-eight, and a verifier that turns most of the constraints into checks that
fail the build rather than advice competing for attention.

It cannot make anything *tested*. Nothing delivered will have been run — not by
Devin, which has no Windows, no forms and no tenant. The honest ceiling is **a
first pass that runs and needs refinement**, and `UNVERIFIED.md` is the
instrument that tells you where to look first.
