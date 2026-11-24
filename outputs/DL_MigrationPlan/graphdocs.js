const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak
} = require('docx');

const LETTER = { size: { width: 12240, height: 15840 } };
const NAVY = '1F3864', GREY = '595959', RED = '9C2A2A', GRN = '1E6B45', AMB = '8A6100';

const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 },
  children: [new TextRun({ text: t, bold: true, size: 30, color: NAVY })] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 90 },
  children: [new TextRun({ text: t, bold: true, size: 24, color: NAVY })] });
const h3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 170, after: 70 },
  children: [new TextRun({ text: t, bold: true, size: 21, color: '2E5496' })] });
const p = (t, o = {}) => new Paragraph({ spacing: { after: o.after === undefined ? 110 : o.after },
  children: [new TextRun({ text: t, size: 21, italics: !!o.italic, bold: !!o.bold,
    color: o.grey ? GREY : (o.red ? RED : undefined) })] });
const runs = (parts, after = 110) => new Paragraph({ spacing: { after },
  children: parts.map(x => new TextRun({ text: x.t, bold: !!x.b, italics: !!x.i, size: 21, color: x.c })) });
const bullet = t => new Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 55 },
  children: [new TextRun({ text: t, size: 21 })] });
const code = t => new Paragraph({ spacing: { after: 0 }, indent: { left: 200 },
  children: [new TextRun({ text: t === '' ? ' ' : t, size: 17, font: 'Consolas' })] });
const codeBlock = s => s.split('\n').map(code);
const callout = (title, text, color) => new Paragraph({ spacing: { before: 110, after: 150 },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color, space: 12 } }, indent: { left: 190 },
  children: [new TextRun({ text: title + '  ', bold: true, size: 21, color }),
             new TextRun({ text, size: 21 })] });

function cell(t, w, o = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade, color: 'auto' } : undefined,
    margins: { top: 70, bottom: 70, left: 110, right: 110 },
    children: [new Paragraph({ children: [new TextRun({ text: t,
      size: o.code ? 18 : 20, bold: !!o.bold,
      color: o.head ? 'FFFFFF' : undefined, font: o.code ? 'Consolas' : undefined })] })] });
}
function table(widths, header, rows, codeCols) {
  const cc = codeCols || [];
  const trs = [new TableRow({ tableHeader: true,
    children: header.map((t, i) => cell(t, widths[i], { bold: true, head: true, shade: NAVY })) })];
  rows.forEach((r, ri) => trs.push(new TableRow({ cantSplit: true,
    children: r.map((t, i) => cell(t, widths[i],
      { shade: ri % 2 ? 'F2F2F2' : undefined, code: cc.indexOf(i) >= 0 })) })));
  return new Table({ columnWidths: widths,
    width: { size: widths.reduce((a, x) => a + x, 0), type: WidthType.DXA }, rows: trs });
}
const gap = (h = 130) => new Paragraph({ spacing: { after: h }, children: [] });
const brk = () => new Paragraph({ children: [new PageBreak()] });

function titleBlock(title, sub) {
  return [
    new Paragraph({ spacing: { after: 50 }, children: [new TextRun({
      text: title, bold: true, size: 38, color: NAVY })] }),
    new Paragraph({ spacing: { after: 40 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 6 } },
      children: [new TextRun({ text: sub, size: 24, color: GREY })] }),
    p('Glenn Armbruster, Exchange Engineering  |  18 August 2026', { grey: true, after: 220 })
  ];
}
function build(name, children) {
  const doc = new Document({
    numbering: { config: [
      { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 340, hanging: 190 } } } }] },
      { reference: 'stp', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 380, hanging: 250 } } } }] }
    ]},
    styles: { default: { document: { run: { font: 'Calibri', size: 21 } } } },
    sections: [{ properties: { page: { ...LETTER,
      margin: { top: 1000, bottom: 1000, left: 1330, right: 1330 } } }, children }]
  });
  return Packer.toBuffer(doc).then(x => { fs.writeFileSync(name, x);
    console.log('written', name, x.length, 'bytes'); });
}

/* ============================ DOC 1 ============================ */
const d1 = [];
d1.push(...titleBlock('Graph Permission Request',
  'Distribution group migration to Microsoft 365'));

d1.push(h1('1.  Prerequisites — Permissions, Roles and Version Information'));
d1.push(p('Four items are required, and they are granted in three different places by potentially three different people. All four must be in place before the conversion will run. Granting the Graph permissions alone produces an authorization failure that looks like the consent did not take.'));
d1.push(table([700, 1500, 3200, 4000], ['', 'Type', 'What', 'Granted by'], [
  ['1', 'Graph permission', 'Group-OnPremisesSyncBehavior.ReadWrite.All', 'Admin consent, from an Application Administrator or Cloud Application Administrator'],
  ['2', 'Graph permission', 'Group.Read.All  (read-only)', 'The same admin consent, at the same time'],
  ['3', 'Directory role', 'Hybrid Identity Administrator', 'A Privileged Role Administrator, assigned to the operator account (PIM-eligible is acceptable)'],
  ['4', 'Version', 'Entra Connect Sync 2.5.76.0 or later', 'The team that owns the Entra Connect servers — a version check, not a permission']
], [2]));
d1.push(gap());

d1.push(h2('1.1  Graph permissions'));
d1.push(p('This is not a request for general Graph access. The primary scope below governs a single property on a single resource type and grants nothing else.'));
d1.push(table([2500, 6900], ['', 'Value'], [
  ['Permission', 'Group-OnPremisesSyncBehavior.ReadWrite.All'],
  ['Note on the name', 'A HYPHEN after Group, not a period. This is Entra’s resource-specific naming convention and is the most common transcription error.'],
  ['What it permits', 'Reading and setting the on-premises synchronization behavior of groups. Nothing else. It cannot read group membership, mail, or user data.'],
  ['Type needed', 'Delegated (work or school account). An operator signs in and runs the script; no unattended application identity is required.'],
  ['Delegated permission ID', '37e00479-5776-4659-aecf-4841ec5d590a'],
  ['Application permission ID', '2d9bd318-b883-40be-9df7-63ec4fcdc424  (not requested)'],
  ['Admin consent', 'Required.'],
  ['Licensing', 'Microsoft Entra Free. No additional cost.']
], [1]));
d1.push(gap());

d1.push(h3('Supporting scope'));
d1.push(p('The conversion call itself needs nothing further. The script also reads each group object to confirm it is a valid candidate before acting on it, which needs one read-only scope:'));
d1.push(table([2500, 6900], ['Permission', 'Why'], [
  ['Group.Read.All', 'Read-only. Resolves the group by its mail address and confirms it is currently directory-synchronized, so a bad row in the input file cannot send the conversion at the wrong object. Also identifies groups already converted, which makes a partly-completed batch safe to re-run.']
], [0]));
d1.push(gap());

d1.push(h2('1.2  Directory role'));
d1.push(p('The Graph consent is necessary but not sufficient. Microsoft additionally requires the signed-in account to hold a supported directory role for this operation.'));
d1.push(table([2500, 6900], ['', 'Value'], [
  ['Role', 'Hybrid Identity Administrator'],
  ['Role template ID', '8ac3fc64-6eca-42ea-9e69-59f4c7b60eb2'],
  ['Why this one', 'Microsoft documents it as the least-privileged role able to perform the Source of Authority transfer. Global Administrator also works and is not being requested.'],
  ['Assignment model', 'PIM-eligible rather than permanently active is acceptable and preferred — the role is activated for the duration of a change window and expires on its own.']
], [1]));
d1.push(callout('This is a role, not a scope.',
  'It is assigned in Entra role management, not in the app consent screen. It is the single most commonly missed prerequisite for this operation, and its absence presents as a permission error rather than as a missing role.',
  AMB));
d1.push(gap());

d1.push(h2('1.3  Version prerequisite — Entra Connect Sync'));
d1.push(p('Source of Authority transfer is only honored by current sync clients. On an older client the conversion appears to succeed — the Graph call returns 204 — and the next synchronization cycle silently overwrites the object from Active Directory again, reverting it.'));
d1.push(table([2500, 6900], ['', 'Value'], [
  ['Minimum version', 'Microsoft Entra Connect Sync 2.5.76.0'],
  ['Applies to', 'Every Entra Connect server synchronizing the affected directories, including any staging-mode server.'],
  ['Why it matters', 'Older clients ignore the isCloudManaged flag and continue to export the object, so the group reverts to directory-managed without raising an error.'],
  ['How to check', 'The version is shown in the Entra Connect wizard, in Programs and Features, or via Get-ADSyncGlobalSettings on the sync server.'],
  ['Action needed', 'Confirmation of the installed version from the Entra Connect owners before the first production conversion. An upgrade, if required, is a standard in-place upgrade.']
], [1]));
d1.push(p('No conversion should be run in production until the installed version is confirmed. This is a five-minute check that prevents a silent, tenant-wide revert.'));
d1.push(gap());


d1.push(h1('2.  Mail-enabled security groups — no additional Graph permission'));
d1.push(p('Roughly a third of the estate consists of mail-enabled security groups, which both deliver mail and grant access. The plan splits them: the on-premises security group stays exactly as it is, and a separate cloud distribution group is created to carry the mail.'));
d1.push(callout('No further permission is needed for this.',
  'Creating a cloud distribution group is an Exchange Online operation, performed with Exchange PowerShell under existing Exchange role-based access. It does not touch Microsoft Graph. The Graph permission in section 1 is used only on the groups being converted in place.',
  GRN));
d1.push(p('These groups are deliberately NOT converted with the Graph permission. Converting one would move its access-granting function to the cloud along with its mail, and any on-premises permission depending on that group would be affected. Leaving them in place removes that risk entirely.'));

/* ---- justification ---- */
d1.push(brk());
d1.push(h1('3.  Justification'));

d1.push(h2('3.1  Scale'));
d1.push(table([3200, 2100, 4100], ['Directory', 'Distribution groups', 'Note'], [
  ['Headquarters', '~250', 'Internal population'],
  ['Franchise', '~4,800', 'Approximately 800 offices'],
  ['Total', '~5,050', 'Each currently read-only in Exchange Online']
]));
d1.push(gap());
d1.push(p('Every one of these groups is mastered in on-premises Active Directory. That is what makes them read-only in the cloud, and it is why on-premises Exchange management is still required to administer them.'));

d1.push(h2('3.2  The two available methods'));
d1.push(table([2400, 3400, 3600], ['', 'With the permission', 'Without it'], [
  ['Method', 'Convert in place — one API call per group', 'Delete on-premises, rebuild in the cloud'],
  ['Object identity', 'Preserved', 'New object'],
  ['Addresses, membership, ownership, restrictions, moderation', 'Preserved automatically', 'Must be exported and re-applied'],
  ['Reversible', 'Yes — one call back', 'No'],
  ['Scripts to build and test', 'One conversion loop', 'Capture, cutover and verification']
]));

d1.push(h2('3.3  Risk reduction'));
d1.push(runs([{ t: 'Distribution groups are excluded from Microsoft Entra’s 30-day restore. ', b: true, c: RED },
  { t: 'Microsoft states this explicitly: the recycle bin covers Microsoft 365 groups and cloud security groups, and not distribution groups. A deleted distribution group cannot be recovered. It can only be rebuilt from whatever was exported beforehand.' }]));
d1.push(p('Without the permission, the project performs roughly 5,050 irreversible deletions. With it, it performs roughly 5,050 reversible conversions.'));
d1.push(h3('The attribute surface'));
d1.push(p('Rebuilding a group means re-applying, correctly, every property that is not automatically recreated:'));
d1.push(bullet('The hidden historical address that Outlook auto-complete resolves against. If missed, replies to recent mail fail with a non-delivery report.'));
d1.push(bullet('Every proxy address, including any existing X500 entries.'));
d1.push(bullet('Membership, ownership, and send-on-behalf delegates.'));
d1.push(bullet('Delivery restrictions — a restricted group silently becomes open to everyone if missed.'));
d1.push(bullet('Moderation settings and moderators.'));
d1.push(bullet('The hidden-from-address-lists flag, and Send-As permissions held separately from the group.'));
d1.push(runs([{ t: 'That is roughly ten values per group across about 5,050 groups — on the order of 50,000 values that must survive an export and re-import intact. ', b: true },
  { t: 'Every one of those is silent when it goes wrong. Nothing errors; a restriction is simply gone. The conversion path moves zero of them.' }]));
d1.push(h3('Two further risks that disappear'));
d1.push(bullet('Address reuse. Recreating a group requires its address to be released by directory synchronization first. Where that stalls, the documented remediation tooling runs on the Entra Connect server — the very component this programme works toward retiring.'));
d1.push(bullet('Synchronization deletion thresholds. Entra Connect halts an export that exceeds its deletion threshold, which defaults to 500 objects. Deleting 5,050 groups requires raising that guard; converting them requires no deletions at all.'));

d1.push(h2('3.4  Time saved'));
d1.push(p('Three separate savings. The first is the largest and the easiest to state.'));
d1.push(table([2900, 3200, 3300], ['', 'With the permission', 'Without it'], [
  ['Tooling', 'One conversion loop, roughly a day to write and test', 'Capture, cutover and verification scripts — on the order of a week, plus proving them safe before touching production'],
  ['Change windows', '4 to 5 sessions', '6 to 7 weekends'],
  ['Per batch', 'Convert, wait one sync cycle, verify', 'Export, delete, wait for address release, rebuild, re-stamp, verify']
]));
d1.push(gap());
d1.push(p('The conversion path also removes the wait that makes this weekend work. Recreating a group cannot begin until synchronization has released its address, and that wait is paid on every batch. A conversion has no such gap, which is why the sessions are both fewer and shorter.'));
d1.push(runs([{ t: 'These are estimates and are stated as such. ', b: true },
  { t: 'A pilot exists specifically to replace them with measurements before the bulk of the work is scheduled.' }]));

d1.push(h2('3.5  Microsoft best practice'));
d1.push(p('This is not a workaround or a novel technique. It is the method Microsoft introduced for this purpose and documents as the path for this exact scenario.'));
d1.push(bullet('Group Source of Authority transfer was released by Microsoft in November 2025 for converting synchronized groups to cloud management.'));
d1.push(bullet('Microsoft’s guidance on decommissioning the last Exchange server names Group SOA as the route for distribution lists and mail-enabled security groups specifically, noting they have no other cloud-management flag.'));
d1.push(bullet('Microsoft publishes a sample bulk PowerShell script for organizations converting at scale.'));
d1.push(bullet('Microsoft names Hybrid Identity Administrator as the least-privileged role for the operation, and documents Administrative Unit scoping to narrow it further.'));
d1.push(p('By contrast, there is no Microsoft guidance, script, or documented procedure for deleting on-premises distribution groups and recreating them in the cloud. It is achievable, and this team has done comparable work before, but it is unsupported by omission.'));

d1.push(h2('3.6  What is being asked'));
d1.push(callout('Summary.',
  'Consent for Group-OnPremisesSyncBehavior.ReadWrite.All and the Hybrid Identity Administrator role, together with Group.Read.All for the pre-flight check. This allows about 5,050 distribution groups to be converted to cloud management in place, reversibly, using Microsoft’s own documented method — instead of being deleted and rebuilt with no possibility of recovery.',
  NAVY));

/* ---- script ---- */
// No break here -- it left page 5 holding one callout and half a page of
// white. The break before section 3 is worth a page; a second one is not.
d1.push(h1('4.  Sample script'));
d1.push(p('Preview by default. Nothing is changed unless -Apply is passed. Reads a CSV of candidate groups, converts each one, and writes a result file.'));
d1.push(p('Requires only the Microsoft.Graph.Authentication module — not the full Graph SDK. The conversion has no dedicated cmdlet, so the call is made with Invoke-MgGraphRequest.', { grey: true }));

d1.push(h3('Ordering — read this before running it at scale'));
d1.push(p('The conversion is not recursive. A group that contains other groups must be converted AFTER its children, starting from the lowest level of the hierarchy. The input CSV must already be in that order; this script does not sort it.'));

// Read the real script from disk. Embedding it in a template literal is not
// possible: PowerShell uses the backtick as its line-continuation character,
// which terminates a JS template literal.
d1.push(...codeBlock(
  fs.readFileSync('Convert-DLSourceOfAuthority.ps1', 'utf8')
    .replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\s+$/, '')
));

d1.push(gap(160));
d1.push(h3('Rolling a group back'));
d1.push(p('The same call with the value inverted. The rollback completes on the next synchronization cycle.'));
d1.push(...codeBlock([
  '$uri = "https://graph.microsoft.com/v1.0/groups/$groupId/onPremisesSyncBehavior"',
  '',
  'Invoke-MgGraphRequest -Method PATCH -Uri $uri -ContentType \'application/json\' `',
  '    -Body \'{"isCloudManaged": false}\''
].join('\n')));

d1.push(gap(120));
d1.push(p('Before rolling back, remove any members added in the cloud after conversion and any access-package references created against the group, or the Active Directory takeover misbehaves.', { grey: true }));

d1.push(new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 8 } },
  spacing: { before: 220 },
  children: [new TextRun({ text: 'Permission names, IDs, role and constraints in section 1 were verified against Microsoft Learn on 18 August 2026. Group counts are approximate and from the directory inventory of 13 August 2026.',
    size: 18, italics: true, color: GREY })] }));

/* ============================ DOC 2 ============================ */
const d2 = [];
d2.push(...titleBlock('Microsoft Graph Permissions',
  'Additional scopes for Microsoft 365 engineering'));

d2.push(h1('Purpose'));
d2.push(p('A separate request from the distribution group migration. These scopes are not needed for that project; they address work this team does routinely and currently does slowly, through administration portals, or not at all.'));
d2.push(p('Eighteen scopes. Thirteen are READ-ONLY. The five write scopes are each deliberately narrow, and each replaces a much broader permission that would otherwise be required to do the same job.'));

d2.push(h1('1.  Read-only — directory and identity'));
d2.push(table([3100, 2000, 4300], ['Permission', 'Consent', 'What it enables'], [
  ['User.Read.All', 'Admin', 'Read full user profiles: sign-in name, licenses, manager, department, account state. Removes the need to open a portal for routine lookups, and lets tooling resolve a person from any of the several identifiers a ticket might contain. Microsoft also names this as half of the least-privileged combination for the account-disable scope in section 4.'],
  ['Group.Read.All', 'Admin', 'Read group objects and membership. Supports reporting on group sprawl, ownerless groups, and membership questions raised by audit.'],
  ['Organization.Read.All', 'Admin', 'Read tenant information including subscribed license SKUs and consumed counts. Answers "how many licenses of X are left, and how many are assigned" without a portal visit. Microsoft documents this as the scope required to read the licenses available in a tenant, and it is a prerequisite for any license reporting at all — there is no other route to a purchased total.'],
  ['Application.Read.All', 'Admin', 'Read app registrations, service principals and the permissions they hold. Answers "what application has access to this mailbox or site, and who consented to it" — a question currently answerable only by escalation.'],
  ['Device.Read.All', 'Admin', 'Read registered and joined devices. Supports termination handling (what did this person have) and multi-factor triage (is this the device they are actually signing in from).'],
  ['Directory.Read.All', 'Admin', 'Broad read across directory objects — administrative units, directory settings and relationships that the narrower scopes above do not cover. Read-only.']
], [0]));
d2.push(gap());

d2.push(h1('2.  Read-only — diagnostics and reporting'));
d2.push(p('This section carries the largest time saving and the lowest risk. Everything here is read-only and covers questions this team is asked constantly.'));
d2.push(table([3100, 2000, 4300], ['Permission', 'Consent', 'What it enables'], [
  ['AuditLog.Read.All', 'Admin', 'Read sign-in logs and directory audit logs. Answers "who deleted this group", "when did this account last sign in", and "what changed on this object" from a script rather than a portal search. Directly relevant to the two malformed directory objects found during the migration inventory — the audit log is how you learn who created them and when.'],
  ['Policy.Read.All', 'Admin', 'Read Conditional Access and other tenant policies. This is the answer to "why can’t this user sign in", which currently requires either portal access or escalation to the identity team. Read-only — it cannot alter a policy.'],
  ['UserAuthenticationMethod.Read.All', 'Admin', 'Read which authentication methods a user has registered. Service desk triage for multi-factor problems without escalation. It reads the METHOD types, not credentials.'],
  ['Reports.Read.All', 'Admin', 'Read Microsoft 365 usage reports — mailbox activity, OneDrive usage, Teams activity. Identifies inactive mailboxes and unused licenses, which is a direct cost saving and is also the honest way to answer "is this distribution group still used" during the migration.'],
  ['RoleManagement.Read.Directory', 'Admin', 'Read which accounts hold which administrative roles. Supports periodic privilege review and answers audit questions about who can do what.']
], [0]));

d2.push(brk());
d2.push(h1('3.  Read-only — workload objects'));
d2.push(p('Exchange and SharePoint objects this team administers daily. Read-only: the corresponding changes are still made through Exchange Online and SharePoint PowerShell under existing role-based access, which is unaffected by this request.'));
d2.push(table([3100, 2000, 4300], ['Permission', 'Consent', 'What it enables'], [
  ['Place.Read.All', 'Admin', 'Read rooms, workspaces and room lists. This team maintains several hundred resource objects across the estate, and reconciling capacity, building and regional settings currently means one Exchange call per room.'],
  ['Sites.Read.All', 'Admin', 'Read SharePoint site metadata and sharing configuration. Supports the external-sharing and orphaned-OneDrive reporting this team already produces, without the per-site enumeration those reports do today. This reads site PROPERTIES, not file content.'],
  ['MailboxSettings.Read', 'Admin', 'Read automatic replies, time zone, language and working hours. This is the mailbox’s CONFIGURATION and explicitly not its contents — it cannot read a message.']
], [0]));
d2.push(gap());

d2.push(brk());
d2.push(h1('4.  Write scopes'));
d2.push(p('Five, each replacing a broader permission that would otherwise be required to do the same job.'));

d2.push(h2('4.1  Termination handling'));
d2.push(p('The termination pipeline this team runs currently stops short of the identity steps, because the only permission that would cover them is broad. Two granular scopes cover them precisely instead.'));
d2.push(table([3100, 2000, 4300], ['Permission', 'Consent', 'What it enables'], [
  ['User.EnableDisableAccount.All', 'Admin', 'Disable (and re-enable) a user account — the accountEnabled property, and nothing else. It cannot change a password, an address, group membership, or any other attribute. Microsoft names this plus User.Read.All as the least-privileged combination for the operation.'],
  ['User.RevokeSessions.All', 'Admin', 'Invalidate a user’s existing sessions and refresh tokens, so access stops immediately rather than when a token happens to expire. Microsoft names this as the least-privileged permission for the operation.']
], [0]));
d2.push(gap());

d2.push(h2('4.2  Licensing'));
d2.push(table([3100, 2000, 4300], ['Permission', 'Consent', 'What it enables'], [
  ['LicenseAssignment.ReadWrite.All', 'Admin', 'Assign and remove license assignments for users and groups. Nothing else — it cannot read or alter any other user property. Reclaiming licenses from leavers is a direct and measurable cost saving, and it is currently done by editing group membership by hand.']
], [0]));
d2.push(p('Until July 2024 the only way to assign a license through Graph was User.ReadWrite.All — full write access to every property of every user in the tenant, in order to change one. Microsoft introduced this scope specifically so that is no longer necessary. Requesting the broad permission when the narrow one now exists would be the harder position to defend.'));
d2.push(gap());

d2.push(h2('4.3  Group and mailbox configuration'));
d2.push(table([3100, 2000, 4300], ['Permission', 'Consent', 'What it enables'], [
  ['Group.ReadWrite.All', 'Admin', 'Create and modify groups, membership and ownership. The immediate driver is ownerless-group remediation: this team reports on groups with no valid owner and cannot currently fix them, so the report is produced and then handed to somebody else. Supersedes Group.Read.All in section 1 if granted.'],
  ['MailboxSettings.ReadWrite', 'Admin', 'Set automatic replies on a mailbox. Requested for the leaver case, where an out-of-office redirecting correspondents to a named colleague has to be set after the user can no longer sign in. Still cannot read mail.']
], [0]));
d2.push(gap());

d2.push(h2('Why granular rather than one broad scope'));
d2.push(runs([{ t: 'The alternative to most of the above is User.ReadWrite.All, ', b: true },
  { t: 'which permits changing any attribute on any user in the tenant — names, addresses, licenses, and more. It would be granted to accomplish a handful of specific actions and would carry the authority to perform hundreds of others.' }]));
d2.push(p('Each scope above does one job. Where Microsoft has published a narrow permission for an operation, the narrow one is what is requested here.'));

d2.push(h2('Risk reduction'));
d2.push(bullet('A terminated account currently remains signed in on any device holding a valid refresh token until that token expires. Session revocation closes that window at the moment of termination rather than hours later. This is the single largest security gap in the current process.'));
d2.push(bullet('The disable step is performed today through a portal, by hand, as a separate action from the rest of the termination. Manual steps in a sequence get missed — and a termination that is 90 per cent complete looks finished.'));
d2.push(bullet('Bringing both steps into the automated pipeline means they are logged in the same audit trail as everything else the tool does, rather than existing only as a portal action nobody can later reconstruct.'));

d2.push(h1('5.  What is not being requested'));
d2.push(p('Stated explicitly, because a permission request is easier to assess when the boundaries are visible.'));
d2.push(bullet('NO ACCESS TO MAIL CONTENT. Mail.Read and its relatives are not requested and are not needed. MailboxSettings covers a mailbox’s configuration — out-of-office, time zone, working hours — and confers no ability to read a message.'));
d2.push(bullet('No access to file or document content. Sites.Read.All reads site properties and sharing configuration; Files.Read.All is not requested.'));
d2.push(bullet('No ability to create or delete users.'));
d2.push(bullet('No password reset capability.'));
d2.push(bullet('No write access to Conditional Access or other policies. Policy.Read.All is read-only and is requested so this team can EXPLAIN a sign-in failure, not change the rule that caused it.'));
d2.push(bullet('No write access to role assignments. RoleManagement.Read.Directory cannot grant anybody anything.'));
d2.push(bullet('No broad user write. User.ReadWrite.All is deliberately not requested; the granular scopes in section 4 exist precisely so it does not have to be.'));
d2.push(bullet('No application permissions. Every scope above is requested as DELEGATED — an engineer signs in and the action is attributable to them. Nothing runs unattended, and nothing works when nobody is signed in.'));

d2.push(h1('6.  Best practice alignment'));
d2.push(bullet('Least privilege by scope. Where Microsoft offers a granular permission, the granular one is requested rather than the broad parent — which is Microsoft’s own stated guidance on choosing permissions. Three of the five write scopes exist only because Microsoft split them out of User.ReadWrite.All for exactly this reason.'));
d2.push(bullet('Delegated rather than application. Actions remain attributable to a named engineer, which is both the safer default and the more auditable one.'));
d2.push(bullet('Read-only wherever read-only will do. Thirteen of the eighteen scopes requested cannot change anything.'));
d2.push(bullet('Every action taken with these scopes is written to the toolkit’s own audit log as well as to the tenant audit log, so there are two independent records of anything done.'));

d2.push(new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 8 } },
  spacing: { before: 220 },
  children: [new TextRun({ text: 'Permission names and behaviour verified against Microsoft Learn, 18 August 2026.',
    size: 18, italics: true, color: GREY })] }));

build('Graph-DLMigration.docx', d1)
  .then(() => build('Graph-M365Engineer.docx', d2));
