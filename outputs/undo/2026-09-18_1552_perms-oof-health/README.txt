Pre-change copies for the 2026-09-18 permissions / OOF / health round.

WHAT IS HERE
  MailboxEditor.ps1   as it was before the Access and Folder Permissions
                      rework, the Health and Statistics tabs, and the
                      Automatic Replies change.
  Common.ps1          as it was before the Out of Office editor moved in and
                      before Enable-ToolkitRemoteMailbox /
                      Get-TenantRemoteRoutingDomain were added.
  Unified-Lookup.ps1  from earlier the same day -- BEFORE the Enable mailbox
                      button AND before the Mailbox Editor hand-off row.
                      Restoring it undoes both.

WHAT IS NOT HERE, AND WHY
  OOFEditor.ps1.  I did not take a copy before moving its editor into
  Common.ps1, which I should have. Nothing is lost -- the moved body is in
  Common.ps1 verbatim, under '#region Out of Office editor -- shared', and the
  only edits were three added parameters (-Smtp, -NoTargetBar,
  -HelpScriptPath) and the two lines that use them. To reverse it: copy that
  region back into OOFEditor.ps1 between its param block and the final call,
  drop the three parameters, and restore
      Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath ...
      $bar = Add-ToolTargetBar ...   (unconditional)
  That is a ten-minute job, not a lost file -- but it is worse than having
  taken the copy.
