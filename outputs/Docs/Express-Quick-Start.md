# Express Admin Toolkit v3.x -- Quick Start

Two pages. Setup is one-time; after this you just connect and work.

**Stuck, or wondering why?** Open **Express Tenant Setup Guide** -- it has the
same steps with the reasoning. Section names below match it.

---

## 1. Unblock  *(2 min)*

Windows blocks files that came from another machine. Nothing runs until this is
done.

1. Right-click **`UnBlock.ps1`** -> **Properties** -> tick **Unblock** -> OK
2. Open **Windows PowerShell** in the toolkit folder and run:

       .\UnBlock.ps1

> Windows PowerShell 5.1 -- *not* PowerShell 7.

## 2. Launch

    .\AdminConsole.ps1

A prompt about a missing `Config\Tenant.json` is expected on a new copy. Say
**Yes** to open Tenant Manager.

## 3. Tenant Manager -- confirm these

Your copy ships pre-filled. **Check, do not retype.**

| | |
|---|---|
| Tenant name | `Express` |
| Primary domain | `expresspros.com` |
| onmicrosoft | `expresspros.onmicrosoft.com` |
| Capabilities | **all six ticked** |
| Directories | `HQ` and `Franchise` (Franchise = default) |
| Exchange server | `exch-vm-hybrid1.expresspersonnel.com` |
| Licence groups | four rows, E1/E5 x HQ/Franchise |

**Validate**, then **Save**.

## 4. Tool Manager -> Settings

Three things only:

- **Reports** and **Logs** -- a folder you can write to
- **Your email** -- your own `@expresspros.com` address
- **Edge profiles** -- your profile names as shown in Edge

Leave **Retry queue**, **SPO admin URL** and **Exchange server** blank.

## 5. Theme  *(optional)*

Theme Manager -> **Add shipped themes** -> pick one -> **Save**.

Worth 30 seconds: colour is how you tell two tenants apart at a glance.

## 6. Back up  *(30 sec)*

**Config Backup** -> **Back up now**.

Do it again before you ever copy a newer toolkit folder over this one.

---

## Every session

    Activate  ->  Connect ExO  ->  Connect SPO  ->  Connect OnPrem

**Activate first** if your rights are on-demand -- a session made before
activating does not carry rights granted afterwards. The console closes while
a sign-in window is open and reopens by itself.

---

## Ready check

    [ ]  Header shows  Tenant: expresspros.com
    [ ]  ExO: Connected      [ ]  SPO: Connected
    [ ]  EMS + AD: Connected            (on-premises tools only)
    [ ]  Dashboard -> Toolkit Health -> Files verified shows a count
    [ ]  Version line at the bottom is not amber
    [ ]  A mailbox lookup returns  Found: Mailbox (HQ)  or  (Franchise)

All ticked -- ready for production use.

---

## Three things that will bite you

**On-premises Exchange tools need "Own window".** On-prem and Exchange Online
share cmdlet names; in one process the last one connected wins. A tool without
Own window can act on the wrong directory **with no error at all**. If in
doubt, tick it.

**Never copy a toolkit folder over another without backing up `Config\`
first.** It silently replaces your settings, tool list, theme and tenant file.

**Never hand-edit the .json files.** Use the manager tools. One wrong character
takes out every button, not one.

---

## When something is wrong

**Greyed button** -- hover it. The tooltip says which: wrong object type (look
up something else), not connected (connect it, activating first), or not
available in this tenant (connecting will not help).

**Anything else** -- run **Tester**. It checks the whole folder and names what
is wrong.

Every tool has a **?** button explaining what it does and what it touches. Use
it before running anything unfamiliar -- several of these write to production.
