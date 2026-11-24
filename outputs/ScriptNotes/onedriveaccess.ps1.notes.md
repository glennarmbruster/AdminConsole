# onedriveaccess.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 1 -->

```text
.SYNOPSIS
    WinForms tool to grant or remove OneDrive site collection admin access for custodians.

.DESCRIPTION
    Manages SharePoint Online site collection admin rights across all combinations of one or
    more disabled users and one or more custodians (many-to-many). Every custodian in the
    Custodians box receives (or loses) IsSiteCollectionAdmin on every OneDrive site derived
    from the Users box.

    Lookup key: email address (documented exception to the sAMAccountName default).
    Target OneDrive sites are identified by the disabled user's email address, which is used
    to construct the personal-site URL. Local-part alone is accepted; the script silently
    appends @expresspros.com before constructing the URL.

    OneDrive URL format:
        https://expresspros-my.sharepoint.com/personal/<first_last>_expresspros_com

    Actions (mutually exclusive radio buttons):
      Grant Site Collection Admin � calls Set-SPOUser -IsSiteCollectionAdmin $true for each
                                    custodian on each user's OneDrive.
      Remove Site Collection Admin � calls Set-SPOUser -IsSiteCollectionAdmin $false for each
                                     custodian on each user's OneDrive.

    What this script does NOT change:
      - No other site permissions, sharing settings, or document library permissions are modified.
      - No Exchange, Active Directory, or Azure AD / Entra ID attributes are touched.
      - No OneDrive files or folders are moved, deleted, or shared.
      - IsSiteCollectionAdmin on other existing admins is never cleared.

    Pre-flight: users and custodians are validated for non-empty input before any change is made.
    The OneDrive URL is constructed from the email prefix � if the URL is incorrect, Set-SPOUser
    will return an error which is isolated per site and surfaced in the Result box.

    Cmdlet provenance: Set-SPOUser is taken directly from the working script (treated as verified
    per Section 2). Get-SPOUser is used for post-action verification (standard SPO module cmdlet).

    Environment: SharePoint Online only.
    Connection is assumed already established � no Connect-SPOService inside this script.

.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-01
    Revised         : 2026-07-01 0000
    Run from        : Standard PowerShell with Microsoft.Online.SharePoint.PowerShell module loaded
    Required modules: Microsoft.Online.SharePoint.PowerShell
```


## function Invoke-OneDriveAccess

- **L131** "Grant" | "Remove"
- **L158** Set-SPOUser (SharePoint Online) � sets or clears IsSiteCollectionAdmin
- **L168** Verification: read back current site collection admins for this site
- **L190-191** Logging: one row per user (target object), per Section 21 Grammar: <Verb> OneDriveAdmin <user> (custodians: c1+c2)

## region WinForms UI

- **L209** Shared state between Parse and Run
- **L213** --- Form ---
- **L222** --- Warning label (always first � red bold, Section 4) ---
- **L231** --- Disabled User(s) label + box ---
- **L247** --- Custodian(s) label + box ---
- **L263** --- Parsed preview label + box ---
- **L280** --- Action group ---
- **L300** --- Buttons: Parse / Run / Close ---
- **L319** No AcceptButton set � Enter in multiline boxes adds newlines, not trigger Run
- **L321** --- Result label + box ---
- **L337** --- Helper: return selected action string ---

## function Get-SelectedAction

- **L343** --- Helper: populate the parsed preview box ---

## function Update-ParsedPreview

- **L355** --- Parse button: preview without executing ---
- **L385** --- Run button: auto-parses if stale, confirms, then executes ---
- **L393** Silently re-parse if input changed since last Parse click
- **L443** --- Close button ---
