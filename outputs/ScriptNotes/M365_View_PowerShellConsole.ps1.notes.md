# M365_View_PowerShellConsole.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    PowerShell Console view for M365_Host_Web.ps1. OFF BY DEFAULT -- the host
    only loads this file when started with -EnableConsole.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms PowerShellConsole.ps1.

    WHY THIS ONE IS OPT-IN AND THE OTHER SEVEN AREN'T -- read before enabling:

    Every other view exposes a FIXED set of operations. This one executes
    whatever PowerShell you type, inside the host process, which is holding a
    live privileged Exchange Online session (and possibly SPO and a directory
    module too). That is a categorically larger exposure than the rest of the
    toolkit, for two reasons:

      1. Scope. The other views can grant a permission or convert a mailbox.
         This can run anything the connected account can do -- across every
         connected service, plus the local filesystem, plus the network.
      2. Reachability. The host listens on localhost, and any web page open in
         the same browser can attempt requests to localhost. Loopback binding
         stops other MACHINES, not other TABS.

    Reason 2 is why M365WebCommon.ps1's request-authorization layer exists
    (per-run random token + Origin check, required on every /api/ call). That
    layer is what makes this view defensible rather than reckless -- a page that
    never loaded from this host cannot know the token. It is a genuine
    mitigation, not a fig leaf. But it is also the ONLY thing standing between
    a hostile page and arbitrary code execution as you, which is why enabling
    this is a deliberate switch rather than the default.

    NOT ENABLED unless the operator passes -EnableConsole to M365_Host_Web.ps1.
    If you don't need it, don't enable it -- the WinForms PowerShellConsole.ps1
    has none of this exposure because it isn't reachable over HTTP at all.

    HOW IT EXECUTES: the submitted text is run in the host's own runspace via
    the call operator on a scriptblock created from it, so it sees the host's
    live connections and imported modules -- which is the entire point (a
    separate runspace would have no Exchange session, making the view useless).
    Output is captured with Out-String so objects render the way they would in
    a console.

    NO SANDBOX, deliberately and explicitly: there is no allow-list, no
    language-mode restriction, and no attempt to block dangerous cmdlets. A
    partial restriction would give a false sense of safety while being trivial
    to bypass, so the honest design is an unrestricted console that is off by
    default and clearly labelled, rather than a "safe" one that isn't.

    Errors are captured and returned as text rather than thrown, so a typo
    doesn't take down an all-day host.
```

- **L144-147** Log every command BEFORE running it, so the completed-tasks log records what was attempted even if it then fails or the host dies partway through. This is the only audit trail this view has, and an unrestricted console deserves one.
- **L152-154** Executed in the host's own runspace on purpose -- it must see the host's live connections and imported modules. See the header on why there is no sandbox.
- **L159-161** 2>&1 folds error records into the stream above, so they're already in $out. $Error is checked only to set the flag that colours the status line.
- **L171** A parse failure (bad syntax) lands here rather than above.


## Embedded HTML / JavaScript comments

- Ctrl+Enter runs; plain Enter stays a newline so multi-line scripts work.
