import os, re, shutil
D='/sessions/sweet-nifty-gauss/mnt/outputs/AdminConsole/Projects/Migration/'
U='/sessions/sweet-nifty-gauss/mnt/outputs/undo/2026-10-05-migration-tag/'
os.makedirs(U, exist_ok=True)
for f in os.listdir(D):
    if f.endswith('.ps1') or f=='README.md': shutil.copy(D+f, U+f)

def load(n):
    s=open(D+n,encoding='utf-8-sig',newline='').read(); assert '\r\n' in s; return s.replace('\r\n','\n')
def save(n,s): open(D+n,'w',encoding='utf-8-sig',newline='').write(s.replace('\n','\r\n'))
def rep(s,o,n,c=1):
    k=s.count(o); assert k==c,(k,o[:80]); return s.replace(o,n)

# ---------------- _Shared: the tag
s=load('_Shared.ps1')
s=rep(s,'''function Get-RecipientSmtp {''',r'''# THE DISPLAY-NAME TAG (2026-10-05). The admin center's Deleted users page
# searches on display name, so a batch tagged 'SMB-10-10-2026-' is one
# search, one select-all, one Restore -- instead of 456 of them. Step 1
# writes it on-premises AFTER exporting the clean name; Step 2 refuses to
# call the batch ready until the tag has synced; Step 4 (restamp) puts the
# clean name back from the exports. Date is the day Step 1 ran.
$script:TagPattern = '^SMB-\d{1,2}-\d{1,2}-\d{4}-'
function Get-MigrationTag { return ('SMB-' + (Get-Date -Format 'M-d-yyyy') + '-') }
function Test-MigrationTagged { param([string]$DisplayName) return ("$DisplayName" -match $script:TagPattern) }
function Remove-MigrationTag { param([string]$DisplayName) return ("$DisplayName" -replace $script:TagPattern, '') }

function Get-RecipientSmtp {''')
s=rep(s,'''        this every Get-RemoteMailbox / Get-ADObject in Steps 1, 3, 3R and 7''','''        this every Get-RemoteMailbox / Get-ADObject in Steps 1, 3, 3R and 6''')
save('_Shared.ps1',s)

# ---------------- 1: tag after export
s=load('1-ONPREM-Export-SMB.ps1')
s=rep(s,'''# STEP 1 of the Migration project. Exports every ON-PREMISES attribute of the
# shared mailboxes named in SMB-Input.csv.
#
# READ ONLY. Nothing is changed.
''','''# STEP 1 of the Migration project. Exports every ON-PREMISES attribute of the
# shared mailboxes named in SMB-Input.csv, THEN TAGS THEIR DISPLAY NAMES.
#
# THE TAG. After the export (which keeps the clean name), each mailbox's
# DisplayName is prefixed on-premises: 'SMB-<M-d-yyyy>-' + the name. Only
# DisplayName -- never alias, addresses or UPN. Why: the restore after Step 3
# is done by hand in the admin center (Users > Deleted users), and that page
# searches display names. One search for the tag, select all, Restore.
# Step 4 (restamp) puts the clean name back from this export.
#
# The tag must reach the cloud BEFORE Step 3 deletes the cloud copy, so:
# run this, wait for a sync (~30 min), then Step 2 -- which checks the tag
# is in the cloud and says NOT READY until it is. Mailboxes already tagged
# are left alone, so re-running is safe.
''')
s=rep(s,'''# Run Step 2 (the cloud copy) next.''','''# Run Step 2 (the cloud copy) after the next sync.''')
s=rep(s,'''* Nothing is changed. This only reads.
* Output: Migration-SMB-ONPREM_<date>_<time>.csv next to this script.''','''* Output: Migration-SMB-ONPREM_<date>_<time>.csv next to this script.
* THEN each mailbox's DisplayName is prefixed on-premises with
  '$(Get-MigrationTag)' so the batch can be restored from the admin
  center with one search. Alias, addresses and UPN are not touched.
* Wait for a sync (~30 min) before Step 2.''')
s=rep(s,'''$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }

# The MemberOf warning''','''$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }

# ---- the tag: written AFTER the export, so the export holds the clean name ----
$tag = Get-MigrationTag
$tagged = 0; $already = 0; $tagFail = 0
Write-Host ''
Write-Host "Tagging display names with '$tag' ..." -ForegroundColor Cyan
foreach ($r in ($rows | Where-Object { -not $_.Error })) {
    $cur = "$($r.DisplayName)"
    if (Test-MigrationTagged $cur) { $already++; Write-Host "  already  $($r.PrimarySmtpAddress)  ($cur)" -ForegroundColor DarkGray; continue }
    $p = @{ Identity = "$($r.DistinguishedName)"; DisplayName = ($tag + $cur); Confirm = $false; ErrorAction = 'Stop' }
    $srv = Get-MigrationAdServer -DistinguishedName "$($r.DistinguishedName)"
    if ($srv) { $p['DomainController'] = $srv }
    try {
        Set-RemoteMailbox @p
        $tagged++; Write-Host "  tagged   $($r.PrimarySmtpAddress)  ->  $($tag + $cur)" -ForegroundColor Green
    } catch {
        $tagFail++; Write-Host "  FAILED   $($r.PrimarySmtpAddress)  $($_.Exception.Message)" -ForegroundColor Red
        Write-FailedLog -Action 'MigrationTagDisplayName' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -ErrorRecord $_
    }
}
Write-Host "Tagged $tagged, already tagged $already, failed $tagFail."
if ($tagged) { Write-CompletedLog -Action 'MigrationTagDisplayName' -Target "$tagged mailboxes" -Result $(if ($tagFail) { 'Partial' } else { 'Success' }) -Detail "prefix $tag" }
Write-Host "NEXT: wait for a directory sync (~30 min), then run Step 2 -- it confirms the tag reached the cloud." -ForegroundColor Cyan

# The MemberOf warning''')
save('1-ONPREM-Export-SMB.ps1',s)

# ---------------- 2: tag check + strip
s=load('2-CLOUD-Export-SMB.ps1')
s=rep(s,'''# READ ONLY. Nothing is changed.
''','''# READ ONLY. Nothing is changed.
#
# THE TAG GATE (2026-10-05). Step 1 prefixed each DisplayName on-premises;
# this step checks that prefix has synced to the cloud and says NOT READY
# until every mailbox carries it. The exported DisplayName is the CLEAN one
# (prefix stripped), so Step 4 restamps the real name and Step 5 verifies
# against it.
''')
s=rep(s,'''# Needs Exchange Online connected in this window. Run Step 3 next -- but
# Step 3 is the one that starts the clock, so read its help first.''','''# Needs Exchange Online connected in this window. When it says READY, run
# Step 3 -- the one that starts the clock, so read its help first.''')
s=rep(s,'''            DisplayName                          = "$($mb.DisplayName)"''','''            DisplayName                          = (Remove-MigrationTag "$($mb.DisplayName)")
            TaggedInCloud                        = $(if (Test-MigrationTagged "$($mb.DisplayName)") { 'True' } else { 'False' })''')
s=rep(s,'''$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }
Write-CompletedLog -Action 'MigrationExportSMB' -Target 'cloud\'''','''$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }

# ---- the tag gate ----
$good = @($rows | Where-Object { -not $_.Error })
$untagged = @($good | Where-Object { "$($_.TaggedInCloud)" -ne 'True' })
Write-Host ''
if ($good.Count -and $untagged.Count -eq 0) {
    Write-Host "READY: the display-name tag from Step 1 is in the cloud on all $($good.Count) mailbox(es). Run Step 3." -ForegroundColor Green
} elseif ($good.Count) {
    Write-Host "NOT READY: $($untagged.Count) of $($good.Count) mailbox(es) do not yet show the Step 1 tag in the cloud:" -ForegroundColor Yellow
    $untagged | ForEach-Object { Write-Host "  $($_.PrimarySmtpAddress)" -ForegroundColor Yellow }
    Write-Host "Wait for the next sync and re-run this step. Do NOT run Step 3 until it says READY -- the tag is what makes the admin-center restore one search." -ForegroundColor Yellow
}
Write-CompletedLog -Action 'MigrationExportSMB' -Target 'cloud\'''')
save('2-CLOUD-Export-SMB.ps1',s)

# ---------------- 3: after-move instructions
s=load('3-ONPREM-MoveToNonSyncOU.ps1')
s=rep(s,'''# AFTER THIS: WAIT. Sync runs every 30 minutes; allow 75 minutes (two
# cycles plus slack) before Step 4, then confirm in the cloud that the
# mailboxes show as soft-deleted (Step 4 shows you). If the batch is larger
# than the Entra Connect export deletion threshold (500 by default), the
# sync will refuse the deletes and Step 4 will find nothing -- see
# Disable-ADSyncExportDeletionThreshold on the sync server.''','''# AFTER THIS: WAIT, THEN RESTORE BY HAND. Sync runs every 30 minutes; allow
# 75 minutes (two cycles plus slack). Then, in the Microsoft 365 admin
# center: Users > Deleted users > search the Step 1 tag (SMB-<date>-) >
# select all > Restore (keep password). Exchange reconnects each mailbox to
# its restored user within a few minutes. Then run Step 4. If the batch is
# larger than the Entra Connect export deletion threshold (500 by default,
# and other deletions in the same cycle count), the sync refuses the deletes
# -- see Disable-ADSyncExportDeletionThreshold on the sync server.''')
s=rep(s,'''* On the next sync the cloud copies are DELETED. Mailboxes become
  soft-deleted (30 days, restorable in Step 4).
* Then WAIT 75 MINUTES before Step 4.''','''* On the next sync the cloud copies are DELETED (soft-deleted, 30 days).
* Then WAIT 75 MINUTES, restore the users in the admin center (Deleted
  users > search the tag > select all > Restore), wait a few minutes,
  and run Step 4.''')
s=rep(s,'''Write-Host "$ok moved, $failed failed. Sync runs every 30 minutes -- WAIT 75 MINUTES, then run Step 4." -ForegroundColor Cyan''',
'''Write-Host "$ok moved, $failed failed. Sync runs every 30 minutes -- WAIT 75 MINUTES." -ForegroundColor Cyan
Write-Host "THEN: Microsoft 365 admin center > Users > Deleted users > search '$(Get-MigrationTag)' (the Step 1 tag -- use that day's date) > select all > Restore." -ForegroundColor Cyan
Write-Host "      Wait a few minutes for the mailboxes to reconnect, then run Step 4 (restamp)." -ForegroundColor Cyan''')
s=rep(s,'''#     Migration-Moved_<date>_<time>.csv   what moved, from where -- Steps 4
''','''#     Migration-Moved_<date>_<time>.csv   what moved, from where -- Steps 4, 5
''')
save('3-ONPREM-MoveToNonSyncOU.ps1',s)

# ---------------- remove 4; renumber 5->4, 6->5, 7->6
os.makedirs('/sessions/sweet-nifty-gauss/mnt/outputs/obsolete/Migration-4-CLOUD-Restore_2026-10-05', exist_ok=True)
shutil.move(D+'4-CLOUD-Restore.ps1','/sessions/sweet-nifty-gauss/mnt/outputs/obsolete/Migration-4-CLOUD-Restore_2026-10-05/4-CLOUD-Restore.ps1')
M={'4':'(gui)','5':'4','6':'5','7':'6'}
def renum(text):
    # file names / titles first
    text=re.sub(r"\b([567])(-(?:ONPREM|CLOUD)-)", lambda m: M[m.group(1)]+m.group(2), text)
    text=re.sub(r"'([567]) - (ONPREM|CLOUD)", lambda m: "'"+M[m.group(1)]+" - "+m.group(2), text)
    def r2(m):
        word=m.group(1)+m.group(2); nums=re.findall(r'\d',m.group(3)); out=[]
        for n in nums:
            if n in M: out.append(M[n])
            else: out.append(n)
        return word+' '+(out[0] if len(out)==1 else ', '.join(out[:-1])+' and '+out[-1])
    text=re.sub(r'\b(STEP|Step|step)(s?) ((?:[4-7])(?:(?:, | and | or |, and )(?:[4-7]))*)(?![\dR-])', r2, text)
    text=text.replace(' of 7 --',' of 6 --')
    return text
for old,new in [('5-CLOUD-Restamp.ps1','4-CLOUD-Restamp.ps1'),('6-CLOUD-Verify.ps1','5-CLOUD-Verify.ps1'),('7-ONPREM-Cleanup.ps1','6-ONPREM-Cleanup.ps1')]:
    os.rename(D+old,D+new)
for f in [x for x in os.listdir(D) if x.endswith('.ps1') or x=='README.md']:
    raw=open(D+f,'rb').read(); bom=raw.startswith(b'\xef\xbb\xbf'); t=raw.decode('utf-8-sig'); t2=renum(t)
    if t2!=t: open(D+f,'wb').write((b'\xef\xbb\xbf' if bom else b'')+t2.encode())

# ---------------- 4 (restamp): shared conversion, not-found wording
s=load('4-CLOUD-Restamp.ps1')
s=rep(s,'''    if (-not $mb) { Note 'SMB' $smtp 'mailbox' 'not found -- Step (gui) not run for it?' 'skipped'; continue }
    if ("$($mb.IsDirSynced)" -eq 'True') { Note 'SMB' $smtp 'mailbox' 'still directory-synced -- restamp would be overwritten' 'skipped'; continue }''',
'''    if (-not $mb) { Note 'SMB' $smtp 'mailbox' 'not active yet -- restore its user in the admin center (Deleted users > Restore), wait a few minutes, re-run' 'skipped'; continue }
    if ("$($mb.IsDirSynced)" -eq 'True') { Note 'SMB' $smtp 'mailbox' 'still directory-synced -- restamp would be overwritten' 'skipped'; continue }
    # SHARED STAYS SHARED. A restore through the user can bring the mailbox
    # back as UserMailbox; Learn (Set-Mailbox): -Type Shared.
    if ("$($mb.RecipientTypeDetails)" -eq 'UserMailbox') {
        try { exo 'Set-Mailbox' @{ Identity = $smtp; Type = 'Shared'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null; Note 'SMB' $smtp 'type' 'converted UserMailbox -> Shared' 'set' }
        catch { Note 'SMB' $smtp 'type' 'convert to Shared' 'FAILED' "$($_.Exception.Message)" }
    } else { Note 'SMB' $smtp 'type' "$($mb.RecipientTypeDetails)" 'present' }''')
s=rep(s,'''Run Step 3 (and (gui)) first.''','''Run Step 3 first, then restore the users in the admin center.''')
s=rep(s,'''# FOR EACH MAILBOX (restored intact in Step (gui), so mostly verification):''','''# FOR EACH MAILBOX (restored by hand in the admin center, so mostly verification):''')
# summary count at end
s=rep(s,'''#   * membership of cloud lists it was in.''','''#   * DisplayName -- which is how the Step 1 tag comes OFF: the exports hold
#     the clean name, so the tagged cloud name is 'different' and is set back
#   * Shared: a mailbox that came back as a user mailbox is converted
#   * membership of cloud lists it was in.''')
save('4-CLOUD-Restamp.ps1',s)

# scrub any "(gui)" leftovers
for f in [x for x in os.listdir(D) if x.endswith('.ps1') or x=='README.md']:
    raw=open(D+f,'rb').read(); bom=raw.startswith(b'\xef\xbb\xbf'); t=raw.decode('utf-8-sig')
    t2=t.replace('Step (gui)','the admin-center restore').replace('step (gui)','the admin-center restore').replace('Steps (gui)','the admin-center restore')
    if t2!=t: open(D+f,'wb').write((b'\xef\xbb\xbf' if bom else b'')+t2.encode())
print(sorted(os.listdir(D)))
