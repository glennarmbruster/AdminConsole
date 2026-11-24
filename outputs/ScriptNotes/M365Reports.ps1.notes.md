# M365Reports.ps1

Launcher for the Microsoft 365 admin portal reports.

## Why this exists instead of more report scripts

Mailbox size, storage quota, item counts and per-user activity are all computed
by Microsoft **daily**, aggregated over 7/30/90/180 days, and exportable to CSV.
Rebuilding any of that in PowerShell means thousands of `Get-MailboxStatistics`
calls to produce a worse answer, slower, that is already stale by the time it
finishes.

The reports this toolkit *does* implement are the ones the portal does not
answer: cross-referencing AD state against cloud state, on-prem group structure,
and anything needing a judgement the portal has no opinion about.

## Browser identity

Every link opens with `-Purpose 'Admin'` through the profile mechanism in §15.
These portals need the elevated account, and a plain `Start-Process` would land
in whichever profile Edge used last — usually the everyday one.

The profile in use is **stated up front** in the window. A silent fallback to
the wrong identity is the failure this whole mechanism exists to prevent, so it
should be visible before anything is clicked rather than discovered at a sign-in
prompt.
