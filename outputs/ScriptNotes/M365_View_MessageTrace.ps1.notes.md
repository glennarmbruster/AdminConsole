# M365_View_MessageTrace.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Message Trace / Quarantine Release view for M365_Host_Web.ps1.
.DESCRIPTION
    ADDED 2026-07-24 as part of the single-process host rewrite. Replaces
    standalone M365_MessageTrace_Web.ps1.

    Contains no logic -- Invoke-MessageTrace, Invoke-MessageTraceDetail,
    Find-QuarantineMatch, and Invoke-QuarantineRelease all live in
    M365Logic.psm1.

    Client-side by design (no server round trip):
      - the results filter box
      - column sorting
      - timezone display conversion (Intl.DateTimeFormat with IANA zone
        names, mapped from the same five labels the WinForms dropdown uses;
        Windows time-zone IDs aren't available to a browser)
      - CSV export (a Blob download of what's currently on screen, filtered
        or not -- the closest honest analog to the WinForms SaveFileDialog)
    Keeping these off the server matters more here than elsewhere: the host
    handles one request at a time (see M365WebCommon.ps1 on concurrency), so
    anything that can be done in the browser should be, to avoid blocking
    the whole host on a re-render.

    CONFIRMATION IS THIS VIEW'S JOB for quarantine release -- the module's
    Invoke-QuarantineRelease never prompts. Each match gets its own button
    and its own confirm(), matching the WinForms tool's one-at-a-time flow.
```


## function applyFilter

- **L460-462** Serves the admin-portal URLs to this view's JS so they aren't hardcoded in two languages. See M365Logic's Get-M365PortalUrl for the URLs themselves and their verification status.
- **L465-469** Accepts optional sender/recipient/messageId so a prepopulating template (if the operator has configured one -- see Set-M365PortalUrlTemplate) can substitute them. Reports *Prepopulated so the UI can tell the operator honestly whether a filter was actually applied.
- **L508-511** Preview-QuarantineMessage's return shape isn't fully documented and varies by message, so render it generically (Format-List over whatever came back) rather than assuming specific property names.


## Embedded HTML / JavaScript comments

- Copy buttons: the real friction in handing a message off to Threat Explorer isn't the link, it's retyping the Message-ID or address. These work regardless of whether URL prepopulation is configured.
- links just stay inert; not worth interrupting the view
- filled from /api/trace/portal-urls on load
- Portal URLs come from the server (M365Logic's Get-M365PortalUrl) rather
- than being hardcoded in this JS, so every admin-portal link in the whole
- toolkit lives in exactly one place and a portal change is a one-line fix.
- Opens Threat Explorer. If the operator has configured a prepopulating
- template (see Set-M365PortalUrlTemplate in M365Logic), the server builds
- a filtered URL and reports prepopulated=true, and the note says so.
- Otherwise it opens the plain page and tells the operator to paste --
- rather than implying a filter was applied when it wasn't.
- Second button, per operator request: releasing is one option, but
- sometimes you want to eyeball the message in the portal first
- (preview it, check the verdict, block the sender instead). Opens
- the Defender Quarantine Email tab in a new tab. Note this lands on
- the quarantine LIST, not this specific message -- Defender has no
- documented per-message deep link, so filtering to the message is
- still a manual step there. Saying so plainly beats implying it
- jumps straight to the message.
- Inline details -- the better answer to "let me look at it first"
- than a portal link, since the toolkit already has the quarantine
- Identity and Defender has no per-message deep link anyway.
- Read-only message preview via Preview-QuarantineMessage.
