# ANZAdminConsole.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    ANZ Admin Console -- Layer 1 launcher for the ANZ (cloud-only,
    expresspros.com.au) tenant. Mailbox / Group lookup + dynamic tool menu.
.DESCRIPTION
    A THIRD, PARALLEL console alongside the original AdminConsole.ps1 and
    SRGAdminConsole.ps1 -- operator request (2026-07-21): "same files, done
    the same way [as SRG]." ANZ is a different, cloud-only engagement with
    its own mail domain (expresspros.com.au) and no on-prem Exchange/AD
    footprint whatsoever. This file, ANZCommon.ps1, ANZTools.json,
    ANZToolManager.ps1, and every ANZ*.ps1 tool script are a complete,
    independent copy of the launcher/tool-script architecture -- nothing
    here is dot-sourced by, or dot-sources, either the original suite's
    files or the SRG suite's files, and neither of those was modified to
    build this.

    Same four jobs as the original AdminConsole.ps1 / SRGAdminConsole.ps1:
      1. Collect the target mailbox/group email address.
      2. Detect the object type (Mailbox / Microsoft 365 Group /
         Distribution Group) via ANZCommon.ps1's Resolve-Target.
      3. Read ANZTools.json and render one button per entry, ALWAYS --
         lookup is optional, not a gate.
      4. Launch the selected tool script IN THIS SAME POWERSHELL SESSION
         (call operator, not a new process) -- same reasoning as the
         original: Connect-ExchangeOnline is per-session state, and a tool
         needs to see whatever's already connected here.

    WHAT'S DIFFERENT FROM AdminConsole.ps1 (identical to SRGAdminConsole.ps1's
    own list, since this is the same architecture applied to a third
    tenant):
      - No Environment radio group, no Domain Controller selector -- not
        commented out, never built at all. ANZ has nothing for them to
        choose between (see ANZCommon.ps1 header). Every ANZ tool script's
        param block is just -SmtpAddress; Start-ToolScript below launches
        with only that one parameter.
      - Object types known: Mailbox, UnifiedGroup, DistributionGroup (not
        the original's Mailbox/UnifiedGroup/OnpremGroup) -- see
        ANZCommon.ps1's Resolve-Target for why DistributionGroup replaces
        OnpremGroup here.
      - ONE "Tools" button panel, not the original's Admin Tools / Daily
        Scripts split. Only 3 tools exist in ANZTools.json today -- a
        single 5-wide x 4-tall FlowLayoutPanel (20-tool nominal capacity,
        same AutoScroll safety net as the original) covers that
        comfortably. If ANZ's tool count grows enough to want two
        categories later, split $flowTools into two panels the same way
        the original's REVISED 2026-07-21 note describes, and add a Section
        column back to ANZTools.json/ANZToolManager.ps1.
      - Layout sized tight to its content from the start, same coordinates
        as SRGAdminConsole.ps1 -- no leftover grey strip on the right or a
        dead zone at the bottom to fix later.
      - Look Up / Clear / Close button row (three buttons) from the start.

    ASSUMPTIONS CARRIED OVER FROM THE SRG SUITE WITHOUT RE-CONFIRMING (see
    ANZCommon.ps1 header for the same list): single Tools section instead
    of a two-way split; the Invoke-EXOCommand collision guard was kept.
    Both are easy to revisit.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (built the same way as SRGAdminConsole.ps1, for a
              third, cloud-only tenant)
```


## region Main window

- **L98-99** REVISED 2026-07-24 -- narrowed 860 -> 600 to make room for the connected- tenant label on the same row. See $lblTenant just below.
- **L105-113** ADDED 2026-07-24 -- shows WHICH tenant Exchange Online is connected to, the same information M365_Host_Web.ps1 puts in its top bar. "Connected" on its own doesn't say connected to what, and this is one of several consoles used in the same day. Right-aligned on the warning row so nothing else had to move. The value is the tenant's LIVE default accepted domain, which is deliberately NOT the same thing as this suite's own $script:MailDomain constant -- if the two disagree, the box is appending a domain from a tenant you are not connected to, and that is exactly what this makes visible.
- **L121-124** Repaints the tenant label. Green when known, grey when not -- a red "unknown" would read as an error when the honest state is simply "nothing connected yet". Cheap after the first call (the lookup is cached), so it is safe to call from the Look Up handler as well as at startup.

## function Update-TenantLabel

- **L131-133** Worth flagging loudly: names typed in the box get $script:MailDomain appended, so a mismatch means this console is resolving into a different tenant than the one it is talking to.
- **L194-198** Dynamic button surface -- FlowLayoutPanel auto-wraps buttons to new rows, so this doesn't need to know in advance how many tools are configured. Sized to fit exactly 5 columns (button 150 + Margin 4*2 = 158px each, 5*158=790 < 860 panel width) x 4 rows (32 + 4*2 = 40px each, 4*40=160 < 172 panel height) before wrapping/scrolling. AutoScroll stays on as a
- **L207-209** One shared ToolTip instance services every tool button -- SetToolTip is called per-button in Update-ToolButtonState to explain why a button is greyed out (or cleared when it isn't).

## region State + lookup

- **L225-231** Launches the tool script with the call operator (&), not Start-Process, so it runs in this same PowerShell session/runspace and sees whatever Exchange Online connection is already established -- same reasoning as the original AdminConsole.ps1's Start-ToolScript (see that file's header for the full trade-off discussion; it applies identically here). Only -SmtpAddress is passed -- ANZ tool scripts have no -Environment / -DomainController parameters to fill in.

## function Start-ToolScript

- **L248-251** Set of object types ANZAdminConsole/Resolve-Target actually knows how to detect. Anything outside this set (most notably 'NotFound', but also plain $null before any lookup has happened) means "no positively-detected type to compare a tool's ObjectTypes against" -- see Update-ToolButtonState.
- **L254-258** Builds the button surface ONCE from $script:AllTools -- every tool gets a button, always, regardless of any lookup. Greying out buttons that don't match a resolved type is Update-ToolButtonState's job below, run separately (and repeatedly, once per lookup) against these same button instances.

## function Build-ToolButtons

- **L274-278** NOTE: no .GetNewClosure() here -- WinForms Add_Click scriptblocks get $this bound to whichever control actually raised the event, so each button already resolves its own .Tag correctly with no closure capture required (same reasoning as the original AdminConsole.ps1's Add-ToolButtonsToPanel).
- **L284-289** Toggles Enabled/tooltip on the already-built buttons -- never adds, removes, or reorders a button. $DetectedType outside $script:KnownObjectTypes (covers $null before any lookup, and 'NotFound'/anything else a failed or no-match lookup produces) means every button is re-enabled: there's no positively-known type to declare a mismatch against, so nothing is greyed.

## function Update-ToolButtonState

- **L294** the "no tools found" fallback label, if present
- **L315-317** Blank input doubles as "clear the last lookup" -- releases any grey-out from a previous lookup instead of just erroring (same as Clear below, just reachable via Look Up too).
- **L342-343** Nothing positively resolved -- release any grey-out from a prior lookup rather than leaving it stale against a failed attempt.
- **L350-351** Same reset $btnRun.Add_Click performs when handed blank input, exposed as its own button.
- **L361-362** Build the full button surface up front -- no lookup has happened yet, so every tool starts enabled.


## Trailing notes

- safety net if more than 20 tools ever land here.
