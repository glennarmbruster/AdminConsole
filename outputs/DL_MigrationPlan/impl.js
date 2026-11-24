const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak
} = require('docx');

const LETTER = { size: { width: 12240, height: 15840 } };
const NAVY = '1F3864';
const GREY = '595959';
const RED  = '9C2A2A';
const GRN  = '1E6B45';

const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 },
  children: [new TextRun({ text: t, bold: true, size: 30, color: NAVY })] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 },
  children: [new TextRun({ text: t, bold: true, size: 24, color: NAVY })] });
const h3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 180, after: 80 },
  children: [new TextRun({ text: t, bold: true, size: 21, color: '2E5496' })] });

function p(t, o = {}) {
  return new Paragraph({ spacing: { after: o.after === undefined ? 120 : o.after },
    children: [new TextRun({ text: t, size: 21, italics: !!o.italic, bold: !!o.bold,
      color: o.grey ? GREY : (o.red ? RED : undefined) })] });
}
function runs(parts, after = 120) {
  return new Paragraph({ spacing: { after },
    children: parts.map(x => new TextRun({ text: x.t, bold: !!x.b, italics: !!x.i, size: 21, color: x.c })) });
}
const bullet = t => new Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 60 },
  children: [new TextRun({ text: t, size: 21 })] });
const step = t => new Paragraph({ numbering: { reference: 'steps', level: 0 }, spacing: { after: 80 },
  children: [new TextRun({ text: t, size: 21 })] });
const mono = t => new Paragraph({ spacing: { after: 100 }, indent: { left: 360 },
  children: [new TextRun({ text: t, size: 19, font: 'Consolas' })] });
const callout = (title, text, color) => new Paragraph({ spacing: { before: 120, after: 160 },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color, space: 12 } }, indent: { left: 200 },
  children: [new TextRun({ text: title + '  ', bold: true, size: 21, color }),
             new TextRun({ text, size: 21 })] });

function cell(t, w, o = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: t, size: 20, bold: !!o.bold,
      color: o.head ? 'FFFFFF' : undefined, font: o.code ? 'Consolas' : undefined })] })] });
}
function table(widths, header, rows) {
  const trs = [new TableRow({ tableHeader: true,
    children: header.map((t, i) => cell(t, widths[i], { bold: true, head: true, shade: NAVY })) })];
  rows.forEach((r, ri) => trs.push(new TableRow({ cantSplit: true,
    children: r.map((t, i) => cell(t, widths[i], { shade: ri % 2 ? 'F2F2F2' : undefined })) })));
  return new Table({ columnWidths: widths,
    width: { size: widths.reduce((a, x) => a + x, 0), type: WidthType.DXA }, rows: trs });
}
const gap = (h = 140) => new Paragraph({ spacing: { after: h }, children: [] });
const brk = () => new Paragraph({ children: [new PageBreak()] });

const b = [];

// ---------------- title ----------------
b.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({
  text: 'Distribution Group Migration', bold: true, size: 40, color: NAVY })] }));
b.push(new Paragraph({ spacing: { after: 40 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 6 } },
  children: [new TextRun({ text: 'Implementation plan', size: 24, color: GREY })] }));
b.push(p('Glenn Armbruster, Exchange Engineering  |  14 August 2026  |  For project meeting',
         { grey: true, after: 240 }));

b.push(h1('1.  How to read this'));
b.push(p('There are two ways to do this work. This document plans for both, but they are not equal and the document is not neutral between them.'));
b.push(bullet('Route A — Group Source of Authority transfer. Microsoft’s documented method. Converts each group in place, preserves everything, and is reversible. Requires one Microsoft Graph permission. This is the recommended route and sections 3 and 5 to 8 assume it.'));
b.push(bullet('Route B — delete and recreate, using Exchange PowerShell only. Requires no new permissions. It works, and this team has done it before, but it is irreversible at scale. Section 4 covers it in full as the fallback.'));
b.push(p('The decision between them is section 2. Everything from section 5 onward — batching, ordering, phasing, exceptions — applies to whichever route is chosen.', { after: 200 }));

// ---------------- 2. the two routes ----------------
b.push(h1('2.  The two routes'));
b.push(p('Both routes end in the same place: distribution groups mastered in Microsoft 365, administered without on-premises Exchange. They differ in how they get there and in what happens when something goes wrong.'));
b.push(table([2300, 3530, 3530],
  ['', 'Route A — Group SOA', 'Route B — delete and recreate'],
  [
    ['Method', 'Convert in place, one API call per group', 'Delete on-premises, rebuild in the cloud'],
    ['Object identity', 'Preserved', 'New object'],
    ['Hidden historical address', 'Preserved automatically', 'Captured and re-applied by hand'],
    ['Membership and ownership', 'Preserved automatically', 'Captured and rebuilt'],
    ['Restrictions and moderation', 'Preserved automatically', 'Captured and rebuilt'],
    ['Reversible', 'Yes — one call back', 'No. Distribution groups are excluded from Entra’s 30-day restore'],
    ['If a batch goes wrong', 'Roll the batch back', 'Rebuild from the capture file, or the groups are gone'],
    ['Tooling to build', 'Minimal. Microsoft publishes a bulk script', 'Capture, cutover and verification scripts'],
    ['New permission required', 'One Graph scope, admin-consented', 'None'],
    ['Change windows', 'Fewer — no synchronization wait between delete and create', 'More — each batch waits on directory synchronization']
  ]));
b.push(gap());
b.push(callout('The recommendation.',
  'Route A. The deciding factor is not effort, it is that Route B has no undo. At 5,000 objects, an irreversible procedure is a poor trade for avoiding a single scoped permission — particularly when the permission is being requested from the same function that is asking for the Exchange decommission.',
  GRN));
b.push(runs([
  { t: 'Route B remains genuinely viable. ', b: true },
  { t: 'It is not a straw man and it is not being presented as impossible. If the permission is declined, the project proceeds on Route B with smaller batches and more verification. Nothing about the phasing or the business case changes.' }
], 200));

// ---------------- 3. route A ----------------
b.push(brk());
b.push(h1('3.  Route A — Group Source of Authority transfer'));

b.push(h2('3.1  What it does'));
b.push(p('Group SOA transfer, released by Microsoft in November 2025, converts a directory-synchronized group to cloud-managed in place. One call per group:'));
b.push(mono('PATCH /groups/{id}/onPremisesSyncBehavior     { "isCloudManaged": true }'));
b.push(p('Entra Connect then stops synchronizing that object, stamps it accordingly, and logs event ID 6956 each time it skips it. The group keeps its identity, its hidden historical address, every proxy address, its membership, its ownership, its restrictions and its moderation settings. Nothing is captured, nothing is rebuilt, and nothing is lost.'));
b.push(p('Microsoft names this as the path for distribution lists and mail-enabled security groups in their own guidance on decommissioning the last Exchange server, and publishes a bulk PowerShell script for organizations at our scale.'));

b.push(h2('3.2  Exactly what would be requested'));
b.push(p('One Microsoft Graph permission. Not Graph access generally — this single scope, which governs one property on one resource type.'));
b.push(table([2600, 6760], ['', 'Value'],
  [
    ['Permission', 'Group-OnPremisesSyncBehavior.ReadWrite.All'],
    ['Note on the name', 'A HYPHEN after Group, not a period. This is the newer Entra resource-specific naming convention and is the most common transcription error.'],
    ['What it governs', 'Reading and setting the on-premises sync behavior of groups. Nothing else.'],
    ['Type', 'Delegated (work or school account) or Application. Delegated is sufficient for an operator running a script interactively.'],
    ['Delegated permission ID', '37e00479-5776-4659-aecf-4841ec5d590a'],
    ['Application permission ID', '2d9bd318-b883-40be-9df7-63ec4fcdc424'],
    ['Admin consent', 'Required. Granted by an Application Administrator or Cloud Application Administrator.'],
    ['Directory role to execute', 'Hybrid Identity Administrator — role template ID 8ac3fc64-6eca-42ea-9e69-59f4c7b60eb2. Microsoft names this as the least-privileged role for the operation. Global Administrator also works but is not required and should not be requested.'],
    ['Other Graph scopes', 'None for the conversion itself. Group.Read.All or Group.ReadWrite.All are needed only by tooling that also reads or edits the group objects.'],
    ['Sync client', 'Entra Connect Sync 2.5.76.0 or later. Older clients ignore the setting and keep overwriting the object.'],
    ['Licensing', 'Microsoft Entra Free. No additional cost.']
  ]));
b.push(gap());

b.push(h3('Narrowing the request'));
b.push(bullet('Administrative Unit scoping. The Hybrid Identity Administrator role can be scoped to an Administrative Unit containing only the groups in scope for this project, rather than granted tenant-wide. Microsoft documents this specifically for source-of-authority operations.'));
b.push(bullet('The scope alone, without Group.ReadWrite.All. If the tool only performs the conversion and does not read or edit group objects, Group-OnPremisesSyncBehavior.ReadWrite.All is sufficient by itself. Microsoft’s sample script requests more because it also enumerates groups.'));

b.push(h3('Two things to disclose up front'));
b.push(runs([{ t: 'There is no read-only variant. ', b: true },
  { t: 'Group-OnPremisesSyncBehavior.Read.All does not exist — the read and the write use the same scope, so a staged "read first, write later" approval is not available.' }]));
b.push(runs([{ t: 'The API reference is published as beta while the operational guidance says v1.0. ', b: true },
  { t: 'Microsoft’s Entra how-to and its bulk script both call the v1.0 endpoint; the Graph reference pages for this resource exist only under /beta. Most likely the reference lags the release. The permission requirement is identical either way.' }]));

b.push(h2('3.3  Which groups SOA suits'));
b.push(callout('This is the most important design point in the document.',
  'SOA converts the WHOLE object. For a pure distribution group that is exactly what is wanted. For a mail-enabled security group it would move the access-granting function to the cloud as well — and any on-premises permission that depends on that group would be affected. The two group types therefore take different treatment even under Route A.',
  RED));
b.push(table([2600, 2100, 4660], ['Group type', 'Count', 'Treatment under Route A'],
  [
    ['Distribution groups', '~3,435', 'Convert in place with SOA. Nothing else required.'],
    ['Mail-enabled security groups', '1,641', 'Do NOT convert with SOA. Leave the on-premises group exactly as it is and create a separate cloud distribution group to carry the mail, as described in section 6.'],
    ['Room lists', 'small', 'Out of scope — moves with the resource mailboxes.'],
    ['Legacy non-universal groups', '2', 'Convert to universal scope first (section 6), then treat as distribution groups.']
  ]));
b.push(gap());
b.push(p('So Route A applies cleanly to roughly 3,435 groups — the large majority — and the mail-enabled security groups follow the same split approach under either route.', { after: 160 }));

b.push(h2('3.4  Prerequisites'));
b.push(table([3900, 5460], ['Item', 'Why'],
  [
    ['Permission granted and consented', 'Section 3.2.'],
    ['Entra Connect Sync at 2.5.76.0 or later', 'Older clients ignore the conversion and keep overwriting the object from Active Directory. Verify the version before converting anything.'],
    ['Nesting order established', 'The conversion is not recursive. See 3.6.'],
    ['Baseline record of every group', 'Not needed to perform the conversion, but needed to prove afterward that nothing changed. Reuses the inventory tooling already built.'],
    ['Validator sweep', 'Malformed objects fail in ways that do not name them. Two already found.'],
    ['Confirm no on-premises application reads these groups by LDAP', 'See 3.8.']
  ]));

b.push(h2('3.5  Per-batch procedure'));
b.push(step('Select the batch, respecting nesting order (3.6) and office grouping (section 5).'));
b.push(step('Record the current state of each group in the batch — membership count, addresses, ownership. This is the comparison baseline, not a recovery file.'));
b.push(step('Convert each group: PATCH onPremisesSyncBehavior with isCloudManaged true.'));
b.push(step('Wait for the next Entra Connect synchronization cycle to complete.'));
b.push(step('Confirm the conversion took: the object reports as cloud-managed and no longer as directory-synchronized.'));
b.push(step('Confirm Entra Connect is skipping the object — event ID 6956 is logged for it.'));
b.push(step('Compare against the baseline: membership, addresses and ownership unchanged.'));
b.push(step('Test that a converted group can now be edited in Exchange Online. This is the actual objective and is worth proving on each batch, not assuming.'));
b.push(step('Record the batch complete. Anything that failed goes to the retry queue.'));

b.push(h2('3.6  Ordering — nested groups'));
b.push(p('The conversion is not recursive. A group containing other groups must be converted after its children, starting from the lowest level of the hierarchy and working upward.'));
b.push(p('Across 4,781 franchise groups this is a dependency sort, not a judgement call, and it is built once before Phase 2 rather than improvised per batch. The nested-group detection needed to produce that ordering already exists in earlier project tooling.'));

b.push(h2('3.7  Rollback'));
b.push(p('Setting isCloudManaged back to false returns the group to Active Directory authority. The rollback completes on the next synchronization cycle.'));
b.push(p('Two conditions before rolling back a group, or the takeover misbehaves:'));
b.push(bullet('Remove any members added in the cloud after conversion — Active Directory does not know about them.'));
b.push(bullet('Remove any access-package or entitlement references created against the group while it was cloud-managed.'));
b.push(p('For mail-enabled groups, Microsoft’s helper script can identify cloud-added members but cannot remove them; removal is done in Exchange.', { after: 160 }));

b.push(h2('3.8  Constraints and side effects'));
b.push(table([3400, 5960], ['Constraint', 'Consequence'],
  [
    ['Not recursive', 'Nested groups convert innermost first. See 3.6.'],
    ['No dual write', 'After conversion, cloud membership changes do not flow back into still-on-premises parent groups.'],
    ['On-premises applications reading these groups by LDAP', 'Lose visibility of the group once it is no longer in Active Directory. Microsoft’s remedy is to de-mail-enable and convert to a plain security group first. Worth confirming before Phase 2 that no franchise application depends on reading these groups directly.'],
    ['Converted groups are not visible in the Entra admin center', 'They are managed through the Exchange admin center and Exchange PowerShell only. This is a change in where administrators look, and worth telling the service desk before it surprises them.'],
    ['Custom LDAP connector in Entra Connect', 'Unsupported with Group SOA. Confirm the connector configuration before starting.']
  ]));

// ---------------- 4. route B ----------------
b.push(brk());
b.push(h1('4.  Route B — the PowerShell fallback'));
b.push(p('Used only if the permission in 3.2 is declined. Requires no new access. Everything here is achievable with Exchange PowerShell and on-premises Active Directory alone.'));
b.push(callout('The governing constraint.',
  'There is no recycle bin for distribution groups and no supported way to undo a deletion. The capture file is the only recovery path. Everything in this section follows from that.',
  RED));

b.push(h2('4.1  Additional prerequisites'));
b.push(bullet('Raise the Entra Connect export deletion threshold. It defaults to 500 deletions and will halt the export mid-run at our batch sizes.'));
b.push(bullet('Capture script built and proven. Nothing is deleted before its capture file is verified readable.'));
b.push(bullet('Verification script built. At 4,781 groups, failures cannot be found by eye.'));

b.push(h2('4.2  Capture — what must be banked'));
b.push(p('Captured before anything is touched. None of it survives deletion, and none of it can be retrieved afterward.'));
b.push(table([3400, 5960], ['Attribute', 'Consequence if missed'],
  [
    ['LegacyExchangeDN', 'Replies from Outlook auto-complete and from older messages fail with a 5.1.1 non-delivery report. Must be re-applied to the new group as an X500 address.'],
    ['EmailAddresses (all)', 'Every proxy and any existing X500 entries. Mail to secondary addresses stops.'],
    ['Members', 'Distribution stops working for everyone on the list.'],
    ['ManagedBy', 'Ownership lost; nobody can be asked about the group afterward.'],
    ['GrantSendOnBehalfTo', 'Delegates silently lose the ability to send on behalf of the group.'],
    ['AcceptMessagesOnlyFrom / RejectMessagesFrom', 'Delivery restrictions disappear — a restricted group becomes open to everyone.'],
    ['Moderation settings and moderators', 'A moderated group becomes unmoderated without notice.'],
    ['HiddenFromAddressListsEnabled', 'A deliberately hidden group reappears in the address book.'],
    ['Send-As permissions', 'Held separately from the group object; lost with it.'],
    ['Office code and full OU path', 'The batch key. Needed to group franchise work by office.']
  ]));
b.push(gap());

b.push(h2('4.3  Per-batch runbook'));
b.push(p('Steps 4 and 8 are stop gates. Do not continue past a failure at either.'));
b.push(step('Select the batch, respecting nesting order and office grouping.'));
b.push(step('Run the capture script. One row per group, everything in 4.2.'));
b.push(step('Verify the capture: row count matches, no blank LegacyExchangeDN, no blank membership on a group that has members.'));
b.push(step('STOP GATE — do not proceed unless the capture is complete and readable. This is the only recovery path.'));
b.push(step('Mail-disable the groups on-premises.'));
b.push(step('Wait for directory synchronization to remove the cloud objects and release the addresses. Confirm rather than assume — query for the address before reusing it.'));
b.push(step('Create the cloud groups from the capture file: primary address, all proxies, membership, ownership, restrictions, moderation, hidden flag.'));
b.push(step('Stamp each old LegacyExchangeDN onto its new group as an X500 address.'));
b.push(step('STOP GATE — run verification. Every captured group exists in the cloud with matching address set and member count.'));
b.push(step('Send a test message to two or three groups, including one with a restriction or moderation set.'));
b.push(step('Record the batch complete. Failures go to the retry queue and are resolved before the next batch.'));

b.push(h2('4.4  Address reuse'));
b.push(p('An address cannot be assigned to two objects at once. If creation fails on a conflict, the old object has not fully released it — wait and retry.'));
b.push(runs([{ t: 'Do not force this. ', b: true, c: RED },
  { t: 'The documented remediation for a stuck address requires tooling that runs on the Entra Connect server, which is the very thing this project works toward retiring. A stuck address resolved by waiting costs an hour; one resolved badly can cost a support case.' }]));

// ---------------- 5. common ----------------
b.push(h1('5.  Batching and ordering'));
b.push(p('Applies to both routes.'));
b.push(h3('By office, never alphabetically'));
b.push(p('An alphabetical batch that fails affects a stripe across all 800 franchise offices. The same failure batched by office affects a handful that can be named and telephoned. The office code is captured for exactly this reason.'));
b.push(h3('Nested groups innermost first'));
b.push(p('Required under Route A because the conversion is not recursive, and under Route B because membership must be rebuilt pointing at objects that still exist. The same ordering serves both.'));
b.push(h3('Batch size'));
b.push(p('Set by the pilot, not assumed. Small enough that a failure is contained and verification is quick.'));

b.push(h1('6.  Exceptions'));
b.push(h2('Mail-enabled security groups'));
b.push(p('Under either route, the on-premises security group stays exactly as it is and only the mail function moves — a new cloud distribution group is created to carry it. Nothing is removed on-premises, so no access can be lost.'));
b.push(p('Under Route A this is a deliberate exclusion from SOA rather than an oversight: converting the group would move the access-granting function to the cloud along with the mail, which is the one outcome to avoid. See 3.3.'));
b.push(h2('The two legacy non-universal groups'));
b.push(p('Both must be converted to universal scope before they can be mail-enabled. Active Directory refuses the conversion in two cases, and the Exchange cmdlet is subject to the same rules — it is not a way around them:'));
b.push(bullet('A group scoped Global cannot convert while it is a member of another Global group.'));
b.push(bullet('A group scoped Domain Local cannot convert while it contains another Domain Local group.'));
b.push(p('Remove the nesting, convert, re-nest. Afterward confirm the change registered on both the directory object and the Exchange recipient — they are read from different places and can disagree.'));

b.push(h1('7.  Phasing'));
b.push(table([1500, 3600, 1900, 2360], ['Phase', 'Content', 'Route A', 'Route B'],
  [
    ['Pilot', 'A few low-risk headquarters groups', '1 session', '1 weekend'],
    ['Phase 1', 'All headquarters groups (298)', '1 session', '1 weekend'],
    ['Phase 2', 'All franchise groups (4,781), by office', '2 to 3 sessions', '4 to 5 weekends'],
    ['Close', 'Verification and reconciliation', '1 week', '1 week']
  ]));
b.push(gap());
b.push(p('Route A durations are shorter because there is no synchronization wait between deleting and recreating, and no rebuild step. They are estimates and the pilot exists to replace them with measurements.', { grey: true }));

b.push(h1('8.  What is needed to proceed'));
b.push(table([3400, 5960], ['From', 'What'],
  [
    ['Information Security', 'Consent for Group-OnPremisesSyncBehavior.ReadWrite.All, and the Hybrid Identity Administrator role — scoped to an Administrative Unit if preferred. This is the decision that selects Route A over Route B.'],
    ['Identity / AD team', 'Confirmation of the Entra Connect Sync version, and that they retain ownership of the on-premises security groups this project leaves in place.'],
    ['The business', 'Change windows, and an owner for communications to franchise offices during Phase 2.'],
    ['This team', 'Conversion and verification tooling, execution, and reconciliation.']
  ]));

b.push(new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 8 } },
  spacing: { before: 240 },
  children: [new TextRun({
    text: 'Object counts from the directory inventory of 13 August 2026. Batch sizes and durations are set by the pilot, which exists to measure them.',
    size: 18, italics: true, color: GREY })] }));

const doc = new Document({
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 200 } } } }] },
    { reference: 'steps', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] }
  ]},
  styles: { default: { document: { run: { font: 'Calibri', size: 21 } } } },
  sections: [{ properties: { page: { ...LETTER,
    margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } }, children: b }]
});

Packer.toBuffer(doc).then(x => {
  fs.writeFileSync('DL-Migration-Implementation-Plan.docx', x);
  console.log('written', x.length, 'bytes');
});
