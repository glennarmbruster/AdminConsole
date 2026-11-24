# obsolete

Archived, not deleted. Nothing here is referenced by anything that still runs.

## `ANZ/` and `SRG/` — archived 2026-07-30

Two cloud-only suites, one per tenant. Superseded by the **FH-Agnostic**
toolkit, which does the same job without being tied to a tenant.

Each folder is **self-contained and still runnable**: the scripts reference only
their own `*Common.ps1`, their own `*Tools.json`, and a `Logs` folder beside
themselves. Everything resolves from `$PSScriptRoot`, so the move did not break
them — run either console from inside its folder and it works as it did.

**Why they were retired rather than maintained.** Measured across 1,873 lines,
ANZ and SRG differed by exactly **two lines** — `$script:MailDomain` and
`$script:SpoTenantPrefix` — and five of the seven file pairs were byte-identical.
That duplication was deliberate isolation, not divergent code, so the moment a
genuinely environment-agnostic suite existed there was nothing left for them to do.

**What only exists here.** `ANZDisable.ps1` and `SRGDisable.ps1` are the only
surviving disable workflows in this repository — the joiner/leaver scripts were
deliberately excluded from the agnostic set as too process-specific. If a
disable tool is ever needed there, start from these rather than from scratch.

**Before deleting:** confirm the CloudOnly toolkit has actually run against
a real tenant. Until then these are the only proof of what a working cloud-only
deployment looked like.

## `M365Web/` — paused 2026-07-31

The browser-hosted suite: one PowerShell process serving HTML views instead of
WinForms windows. **Paused, not abandoned** — it works, and it may come back.

**Why it stopped.** By 2026-07-31 it carried 11 views against 46 WinForms
tools, untouched since 28 July, with none of `Disable-User`, `LogViewer`,
`ResourceDuplicator`, `ReportBrowser` or the three new SharePoint reports. It
had stopped being a parallel toolkit and become a snapshot of an older one.
That is the same shape as ANZ and SRG above: duplication kept in separate
places does not stay identical, it drifts, and the drift is invisible until
somebody diffs them. Two front ends over one set of logic means every feature
gets built twice, which does not fit a monthly ship.

**Two things beyond drift.** The audit log added the same week assumes **one
operator per process** — `Operator` comes from the Windows identity of
whoever launched the tool, and a shared web host either records the host's
identity for everybody or needs real authentication. And a local HTTP listener
with hand-rolled CSRF hardening is more attack surface than a WinForms app on
a jump box.

**What it proved, and what is worth keeping.** The per-process Exchange Online
connection lesson (ARCHITECTURE §4) came from here: the original per-tool
launchers each spawned their own process and so each needed its own sign-in,
which is why the single-process host exists at all. That lesson shaped
everything built since.

`M365Logic.psm1` is the part with a future — 62 functions of cloud logic with
no UI attached. If a second front end is ever built, start from that rather
than from the host.

**What would justify reviving it:** a user WinForms cannot serve. Tier-1
helpdesk self-service, franchise managers checking their own mailboxes, anyone
on a device without PowerShell. That is a different product with authentication
and role-based access, and should be scoped as one rather than maintained as a
parallel port of the console.

`superseded/` holds the five original per-tool launchers that the host
replaced.

## `M365_*_Web.ps1` — superseded per-tool web launchers

Now in `M365Web/superseded/`. Replaced by the single-process `M365_Host_Web.ps1`. Kept as reference for the
architecture change: each of these spawned its own PowerShell process, and an
Exchange Online connection is per-process, so every launcher needed its own
sign-in. The host exists to hold one connection and serve every view from it.

## `_archive/` — moved here 2026-07-31

The pre-comment-stripping originals (`_Originals_WithNotes/`), the extraction
tooling that produced `ScriptNotes/` (`_extract_notes.py`, `_pilot/`), and a
stale file inventory.

It sat at the top level because it predates this folder. After enough rewrites
that the originals no longer resemble what runs, keeping it beside the live
toolkits implied a currency it does not have. Nothing reads any of it.

`_Originals_WithNotes/` is still the only place the reasoning exists **inside**
the code rather than beside it, so it stays until `ScriptNotes/` is confirmed
complete — see the outstanding notes gap in ARCHITECTURE §18.

## Related, and NOT here

- `ScriptNotes/ANZ*.notes.md`, `SRG*.notes.md` and the `M365*` notes — the
  reasoning for the archived code, still in `ScriptNotes/`.
- Nothing. `_archive/` moved in here on 2026-07-31; see above.
