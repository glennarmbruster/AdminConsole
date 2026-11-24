# Connect-ExO.ps1 / Connect-SPO.ps1 / Connect-ExSE.ps1

Three buttons that connect **this PowerShell process** to Exchange Online,
SharePoint Online, and on-premises Exchange. Written 2026-08-03 after the first attempt produced sign-in
windows that drew and then never accepted input.

---

## Why the MFA prompt never finished loading

`AdminConsole.ps1` ends in `$form.ShowDialog()`, which puts the console into a
**modal message loop**. A modal form disables every other window belonging to
the same thread. Tools are launched with `&`, in that same process and on that
same thread — so when the sign-in window appeared it was created into a thread
whose input was already owned by the modal form. It rendered, and then sat
there.

"The MFA page never fully loaded" is what that looks like from the outside.

## The fix that would have made it much worse

The obvious next move is to run the connect in a **separate PowerShell process**
— a new window, no modal loop, prompt works.

It would also have been useless, and confusingly so. Exchange Online and
SharePoint Online sessions belong to the **process that created them** (§4).
A separate process would authenticate perfectly, report success, and leave the
console exactly as disconnected as before — with the operator now certain they
had connected. That is a worse bug than the one being fixed, because the first
one at least looks like a fault.

Whatever the fix was, it had to keep the connection in the console's own
process.

## First attempt: move the prompt out of our thread

Both modules can be told to sign in somewhere our message loop cannot reach.
This fixed SharePoint and **not** Exchange Online -- see below.

**Exchange Online — `-DisableWAM`.** From module 3.7.0, sign-in goes through
Web Account Manager, which raises its own broker window on the calling thread —
precisely what a modal form blocks. Microsoft document the switch for this
class of problem: *"If you encounter WAM-related issues during sign in, you can
use this switch to disable WAM."* Without WAM the flow goes to a browser
instead.

**SharePoint Online — `-UseSystemBrowser $true`.** Sends authentication to the
default browser via MSAL rather than the module's own dialog. A browser is a
**separate process**, so nothing the console does to its own message loop can
affect it. That is what makes this a fix rather than a workaround.

**Both are version-guarded** with `(Get-Command X).Parameters.ContainsKey(...)`.
`-DisableWAM` only exists from 3.7.2-Preview1, and `-UseSystemBrowser` is not in
every build of the SharePoint module. Passing an unknown parameter fails the
entire call, and "your connect script is broken" is a worse outcome than a
prompt that needs an extra click. When the parameter is absent the script says
so and connects the old way.

## Ruled out, and why

Checked against Microsoft Learn rather than assumed:

| Option | Verdict |
|---|---|
| `Connect-ExchangeOnline -Device` | Device-code flow, and it would sidestep the window problem entirely — but documented **"only in PowerShell 7"**. The toolkit is 5.1. |
| `Connect-ExchangeOnline -InlineCredential` | Also PowerShell 7 only, and *"does not work with accounts that use MFA."* |
| A parent-window / HWND parameter | Neither cmdlet has one. There is no supported way to reparent the prompt onto our form. |
| `Connect-SPOService -Device` | Does not exist. SharePoint's only non-dialog paths are the system browser, certificate auth, or managed identity. |

## The switches were not enough, and the structural fix was needed

Tested 2026-08-03. **SharePoint worked; Exchange Online did not.**

`Connect-ExO.ps1` printed its banner and then stopped -- no window, no error,
no return. Run from a bare prompt with the console closed, the same script
signs in normally. That is the decisive test, and it says the console is the
cause.

**Why `-DisableWAM` was not enough.** It does what it says: sign-in stops going
through the WAM broker. But in Windows PowerShell 5.1 the fallback is an
**embedded browser** -- still a window created on our thread, still disabled by
the modal loop. One blocked window swapped for another.

`-UseSystemBrowser` worked for SharePoint precisely because it launches the
default browser, which is a genuinely separate **process**. Exchange Online has
no equivalent: no `-UseSystemBrowser`, `-Device` is PowerShell 7 only, and there
is no parent-window parameter. In PowerShell 5.1 there is no way to make its
prompt appear off our thread.

So the loop has to go, not the window.

### Close, connect, relaunch

The connect button no longer runs its script. It records which tool was clicked
and closes the form:

```powershell
if (Get-ConnectButtonService -Tool $tool) {
    $script:PendingConnect = $tool
    $form.Close()
    return
}
```

After `ShowDialog()` returns there is no modal loop, so the sign-in window takes
input normally. Then the console relaunches **itself**:

```powershell
& $PSCommandPath
```

`& ` and not `Start-Process`, and that distinction is the entire design.
Relaunching in the same process keeps the session that was just established;
a new PowerShell would authenticate perfectly and leave this console exactly as
disconnected as before.

From the operator's side: the console disappears, the sign-in prompt behaves
like any other, and the console comes back connected -- which is what Glenn
proposed at the outset. The cheaper switch-based attempt was tried first
because it would have avoided restructuring the console; it half-worked, and
the half that failed was the one that mattered.

### The depth guard

Each connect nests one invocation. Two or three per session is normal and
harmless. `$global:AdminConsoleDepth` stops at 8 and prints how to relaunch by
hand, so a stuck loop ends in a readable message rather than a stack overflow.

### What this does not change

Only connect buttons defer. Every other tool still runs inline with `&` and
still triggers a re-probe when it returns -- greying stays accurate, and the
modal loop is not a problem for tools that raise no external window.

## The bug these scripts nearly shipped with

The capability scan reads each script and infers what it needs.
`Connect-ExO.ps1` calls `Connect-ExchangeOnline`, so it scans — correctly — as
**requiring Exchange Online**. Acting on that would have greyed the button out
whenever Exchange Online was not connected, which is the only moment it is
useful.

`Get-ToolCapabilityBlock` now returns `$null` for anything
`Get-ConnectButtonService` recognises. **A tool whose job is to create a
connection is never gated on that connection.**

Keyed on the service matcher rather than the section name, because sections are
user-editable in `Sections.json` and renaming one must not quietly re-arm the
greying. The trade is that the matcher is name-based (`exo|exchange`,
`spo|sharepoint`), so a future tool called "SharePoint Sites Report" would also
skip greying. Checked against the current `Tools.json`: no false positives
today.

## Connect-ExSE.ps1 -- on-premises Exchange

Same three-line recipe Microsoft documents, with three changes that matter when
it runs from a button rather than being typed.

### `Import-PSSession` alone would have failed silently

`Import-PSSession` builds a temporary module and **returns** it. It has no
`-Global` parameter, and Microsoft Learn documents nothing about its scope. Run
interactively the commands land where you want them; run inside a `.ps1`
launched with `&` -- which is how every console button works -- there is no
guarantee they outlive the script, and in practice they do not.

The button would have reported success and left the window with no Exchange
cmdlets in it. Nothing would have said so.

The documented fix is to capture the module and re-import it:

```powershell
$module = Import-PSSession -Session $session -DisableNameChecking -AllowClobber
Import-Module -ModuleInfo $module -Global -DisableNameChecking
```

`-Global` is documented as "imports modules into the global session state so
they're available to all commands in the session." That is the guarantee the
bare form does not give.

### No credential prompt by default

`-Credential` is **not mandatory** on `New-PSSession` -- the documented default
is "Current user". With Kerberos that means the logged-on ticket, which is the
normal single-sign-on case on a jump box.

Which matters here beyond convenience: `Get-Credential` raises a Windows
credential dialog, and that is the same modal-loop trap that broke the cloud
sign-in prompts. Not prompting sidesteps it entirely.

`-PromptForCredential` exists for signing in as a different admin account, and
says plainly that if the dialog does not appear the script should be run from
the prompt instead of the button.

### `-AllowClobber`, and the warning that goes with it

Exchange Online may already own `Get-Mailbox`, `Set-Mailbox` and
`Get-DistributionGroup`. Without `-AllowClobber` those imports are skipped and
the connection looks partial for no visible reason.

With it, the on-premises versions win -- which is the collision ARCHITECTURE §2
is about. The script does not block it, because a hybrid admin legitimately
needs both, but it says so once, at the moment it becomes true, rather than
leaving it to be discovered at a call site.

`Set-ADServerSettings -ViewEntireForest $true` is kept: without it recipient
lookups silently return only the local domain, and an empty result reads as "no
such mailbox".

The session is stored in `$global:ExchangeOnPremSession`. Microsoft: "you must
keep the connection to the other session open. If you close it, the imported
commands are no longer available." A session nobody holds a reference to is one
nobody can deliberately close.

## `Get-ConnectButtonService` needed two fixes

Adding a third connect button exposed both.

**Order.** 'Connect Exchange SE' contains the word *exchange*, so with the
cloud pattern tested first the on-premises button would have been painted with
**Exchange Online's** status -- green while the on-prem session was down, and
nothing anywhere to contradict it. The EMS pattern is now tested first.

**Scope.** The EMS pattern includes `on-?prem`, which matched two ordinary
tools -- `1-Onprem-Export` and `2-Onprem-Decomm`. Since this function also
decides which buttons skip capability greying, those two would have quietly
stopped greying out when the on-premises session was missing.

The function now requires the name to contain the word **connect** before any
service pattern is considered. That is the honest test: it asks whether a
button's job is to establish a connection.

## SPO admin URL

`Get-SpoAdminUrl`: `Settings.json` first, then derived.

Deriving is safe because the host is not a free choice — it is always
`https://<tenant>-admin.sharepoint.com`. The only unknown is the tenant label,
and two things already know it:

1. **The session organisation** (`Get-SessionOrganization`), preferred, because
   the SharePoint tenant label follows the `*.onmicrosoft.com` prefix.
2. **The mail domain's first label**, as a fallback when nothing is connected
   yet — which is the normal state when you are about to press Connect.

They agree here. They do not always: a tenant that changed its vanity mail
domain would otherwise get a URL that never existed. Hence the order.

The setting exists for a vanity SharePoint host or a tenant whose SPO label
genuinely differs. Blank is the normal state, and the Settings dialog shows what
blank currently resolves to rather than leaving "empty" as an unknown.

## Other decisions

**Neither script opens a window.** The sign-in prompt is the window; anything of
ours competing with it is the problem being solved. Output goes to the host,
which is where the operator already is.

**Already-connected is not an error.** Both exit early with what the existing
session is, and take `-Force` to reconnect anyway. Re-authenticating for no
reason is how an operator ends up staring at an MFA prompt wondering what they
did wrong.

**The result is proved, not assumed.** `Connect-*` returning without throwing is
not the same as a usable session, so both scripts re-probe with the same
`Test-*ConnectionState` the console uses to colour its buttons. The message and
the button state cannot disagree, because they come from one answer.

**`Clear-ConnectedTenantDomainCache` after connecting to Exchange.** The tenant
is cached by whatever asked first; connecting is exactly the event that can
change it. `Test-TenantDomainMismatch` then runs, so connecting to the wrong
tenant is caught at the moment it happens rather than three tools later.

**Connect-SPO names the likely cause on failure.** The commonest error is the
ordinary SharePoint URL instead of the `-admin` host, and the MSAL error does
not mention it.

## Still open

- One SharePoint connection is supported per PowerShell session, so a second
  connect replaces the first. On a shared jump box that is worth knowing and is
  not currently enforced anywhere.
- None of the three are in `FH-Agnostic/` or `CloudOnly/` yet. Connect-ExO is
  portable as-is; Connect-SPO needs `Get-SpoAdminUrl`; Connect-ExSE is
  on-premises and belongs only in AdminConsole.
- `Connect-ExSE` does not disconnect. `Remove-PSSession
  $global:ExchangeOnPremSession` is the manual step, and there is no button for
  it.

## The first Connect ExO of a session failed, the second worked

Symptom: a forty-line MSAL stack trace ending in
`WinFormsLegacyWebUi.OnAuthenticate`, then

    ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be
    instantiated because the current thread is not in a single-threaded apartment

`8856f961-` is the IE WebBrowser control. `-DisableWAM` selects the embedded
browser rather than the WAM broker, and that control can only be created on an
STA thread with a live message pump.

AdminConsole runs the connect scripts at the very end, after `ShowDialog()`
returns -- deliberately, because a modal form blocks the sign-in window from
taking input. But `ShowDialog` installs a `WindowsFormsSynchronizationContext`
on the thread and does NOT remove it when the form closes. MSAL captures that
context and posts the browser onto a pump that died with the form.

The context is now cleared in two places: in AdminConsole before invoking any
connect script, so the problem is never handed over; and inside Connect-ExO,
which can be run directly. The console-side clear covers Connect-SPO and
PimActivation too -- both do interactive auth through the same PendingConnect
path.

Also: the thread apartment is checked BEFORE attempting, so an MTA host fails
in one readable line instead of forty; the error stream is redirected so the
MSAL plumbing is not printed, since only its last line names a cause; and there
is exactly ONE retry, gated on this specific message. Retrying anything else
would hide it.

CAVEAT: the synchronization-context explanation fits the evidence but was not
proved -- PowerShell could not be run here. The apartment state and any stale
context are now reported before sign-in, so the next run says which it was.
If the retry line appears every time, the context clear is not the cause.
