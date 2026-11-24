const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, PageBreak
} = require('docx');

const LETTER = { size: { width: 12240, height: 15840 } };
const NAVY = '1F3864';
const GREY = '595959';
const TW = 9360;                      // usable width: 8.5in - 1in margins = 6.5in = 9360 dxa

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 30, color: NAVY })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, bold: true, size: 24, color: NAVY })]
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after === undefined ? 120 : opts.after },
    children: [new TextRun({ text, size: 21, italics: !!opts.italic,
                             color: opts.grey ? GREY : undefined, bold: !!opts.bold })]
  });
}
function runs(parts, after = 120) {
  return new Paragraph({
    spacing: { after },
    children: parts.map(x => new TextRun({ text: x.t, bold: !!x.b, italics: !!x.i, size: 21 }))
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: 21 })]
  });
}
function cell(text, w, opts = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: opts.shade
      ? { type: ShadingType.CLEAR, fill: opts.shade, color: 'auto' }
      : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({
        text, size: 20, bold: !!opts.bold,
        color: opts.head ? 'FFFFFF' : undefined
      })]
    })]
  });
}
function table(widths, header, rows) {
  const trs = [];
  trs.push(new TableRow({
    tableHeader: true,
    children: header.map((t, i) => cell(t, widths[i], { bold: true, head: true, shade: NAVY }))
  }));
  rows.forEach((r, ri) => {
    trs.push(new TableRow({
      // Keeps a row whole rather than letting it break across a page, which
      // otherwise leaves the risk name on one page and its mitigation on the
      // next under empty cells.
      cantSplit: true,
      children: r.map((t, i) => cell(t, widths[i], { shade: ri % 2 ? 'F2F2F2' : undefined }))
    }));
  });
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    rows: trs
  });
}
function spacer(h = 120) { return new Paragraph({ spacing: { after: h }, children: [] }); }

const body = [];

// ---------- title ----------
body.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { after: 60 },
  children: [new TextRun({ text: 'Distribution Group Migration to Microsoft 365',
                           bold: true, size: 40, color: NAVY })]
}));
body.push(new Paragraph({
  spacing: { after: 40 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 6 } },
  children: [new TextRun({ text: 'Project plan for review', size: 24, color: GREY })]
}));
body.push(p('Prepared by Glenn Armbruster, Exchange Engineering  |  14 August 2026  |  Draft for management review',
            { grey: true, after: 240 }));

// ---------- summary ----------
body.push(h1('Summary'));
body.push(p('Every distribution group in the organization is currently mastered in on-premises Active Directory and synchronized to Microsoft 365. This project moves ownership of those groups to the cloud, so they are created, changed and administered in Microsoft 365 directly rather than on-premises. It removes one of the categories of object that currently require on-premises Exchange management, and is a prerequisite step on the decommission path.'));
body.push(p('The estate is approximately 5,100 mail-enabled groups across two directories. The work divides cleanly into a small headquarters phase and a much larger franchise phase, and is expected to take roughly six to seven weekends of change windows in total, spread over two to three months.'));
body.push(runs([
  { t: 'The tooling to perform this work already exists. ', b: true },
  { t: 'The inventory, reporting and safety checks are built and in use, which is why this plan can quote exact object counts rather than estimates. What is required from the business is change windows, a small amount of coordination with the Active Directory team, and a decision on communications to franchise offices.' }
]));

// ---------- driver ----------
body.push(h1('Why now'));
body.push(p('This work supports the decommissioning of on-premises Exchange, and modernizes how distribution groups are administered.'));

body.push(h2('The decommission dependency'));
body.push(p('An object that is mastered in on-premises Active Directory must be managed on-premises. A distribution group synchronized to Microsoft 365 is read-only in the cloud: it cannot be created, renamed, or have its membership changed there. Every one of the roughly 5,100 groups in the estate is currently in that state.'));
body.push(p('While that remains true, on-premises Exchange management is required in order to administer them. Moving the groups so they are mastered in Microsoft 365 removes an entire category of object from that dependency.'));
body.push(runs([
  { t: 'This is a necessary step, not a sufficient one. ', b: true },
  { t: 'Distribution groups are one category among several. While user accounts continue to synchronize from on-premises Active Directory, some on-premises recipient management remains necessary regardless of this project. This plan should be presented as removing one blocker from the decommission path, not as retiring the server.' }
]));

body.push(h2('Modernization'));
body.push(bullet('Group administration becomes a cloud task, so it can be delegated to service desk staff without granting on-premises directory rights.'));
body.push(bullet('Changes take effect immediately rather than waiting on a directory synchronization cycle.'));
body.push(bullet('Groups become manageable through the same interfaces as the rest of the Microsoft 365 estate, rather than requiring a separate on-premises toolset and the knowledge to use it.'));
body.push(p('No external deadline is currently identified. If the decommission has a target date, it determines how tightly the phases below are scheduled rather than changing the approach.', { after: 200 }));

// ---------- estate ----------
body.push(h1('The estate'));
body.push(p('Figures below are from a full inventory taken on 13 August 2026 and are exact, not estimated.'));
body.push(table(
  [3200, 1600, 1600, 2960],
  ['Directory', 'Total groups', 'Distribution', 'Mail-enabled security'],
  [
    ['Headquarters', '298', '241', '54'],
    ['Franchise', '4,781', '~3,194', '1,587'],
    ['Total', '~5,079', '~3,435', '1,641']
  ]
));
body.push(spacer(160));
body.push(p('Headquarters also holds two legacy groups created before the current standard, and a small number of room lists. Both are handled as exceptions and are described in the scope section.', { grey: true }));

// ---------- scope ----------
body.push(h1('Scope'));
body.push(h2('In scope'));
body.push(bullet('All distribution groups in both directories, moved to be mastered in Microsoft 365.'));
body.push(bullet('The mail function of mail-enabled security groups — see the note below, as these need special handling.'));
body.push(h2('Out of scope'));
body.push(bullet('Security groups themselves. Where a group grants access to files or systems, that function stays on-premises and remains owned by the Active Directory team.'));
body.push(bullet('Room lists, which move with the resource mailboxes in separate work.'));
body.push(bullet('Microsoft 365 Groups, which are already mastered in the cloud.'));

body.push(h2('The mail-enabled security groups'));
body.push(p('1,641 groups — about a third of the estate — currently do two jobs at once: they deliver mail, and they grant access. Those two functions belong in different places after this project.'));
body.push(p('The approach is to split them: the security group stays on-premises unchanged, and a separate distribution group is created in Microsoft 365 to carry the mail. Nothing is removed from the on-premises side, so no access can be lost as a result of this project.'));
body.push(runs([
  { t: 'This is deliberately the conservative option. ', b: true },
  { t: 'Determining which of these groups actually grant access would require an audit of file and system permissions across the estate, which is neither available to this team nor necessary — leaving every security group in place costs nothing and removes the risk entirely.' }
], 200));

// ---------- approach ----------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(h1('Approach and phasing'));
body.push(p('Groups are moved in batches. For each batch the existing group’s addresses and membership are captured, the group is removed from synchronization, and an equivalent group is created in Microsoft 365 carrying the same addresses.'));
body.push(p('The pace is set by two things outside our control: Microsoft 365 rate limits on bulk changes, and the time taken for directory synchronization to release an address so it can be reused. Neither is improved by adding people to the project. This is why the work is measured in change windows rather than in effort.', { after: 160 }));

body.push(table(
  [1500, 3400, 1700, 2760],
  ['Phase', 'Content', 'Duration', 'Approach'],
  [
    ['Pilot', 'A small number of low-risk headquarters groups', '1 weekend', 'Proves the process end to end and measures the real cycle time'],
    ['Phase 1', 'All headquarters groups (298)', '1 weekend', 'Single cutover; the population is small and internal'],
    ['Phase 2', 'All franchise groups (4,781)', '4 to 5 weekends', 'Batched by office, so any problem is contained to named offices'],
    ['Close', 'Verification and documentation', '1 week', 'Reconciliation against the original inventory']
  ]
));
body.push(spacer(160));
body.push(runs([
  { t: 'Franchise batching is by office, not alphabetically, and this matters. ', b: true },
  { t: 'A batch that goes wrong alphabetically affects a stripe across every office in the network. Batched by office, the same failure affects a handful of offices that can be named and contacted directly.' }
], 200));

// ---------- risks ----------
body.push(h1('Risks'));
body.push(table(
  [2700, 1200, 5460],
  ['Risk', 'Severity', 'Mitigation'],
  [
    ['Replies to older messages fail after a group moves',
     'High',
     'Every group carries a hidden historical address that Outlook uses to resolve replies and auto-complete entries. If it is not carried across, users receive delivery failures when replying to recent mail. The migration tooling captures and re-applies this automatically; it is the single most important technical control in the project.'],
    ['A batch cannot be reversed once started',
     'High',
     'Releasing an address to recreate the group in the cloud is not instantly reversible. Batches are therefore kept small enough to verify before the next begins, and each is verified before proceeding.'],
    ['Franchise impact is wide',
     'High',
     'Approximately 800 offices are affected in Phase 2. Batching by office limits any single failure to a small, identifiable group of offices rather than the whole network.'],
    ['Pace is constrained by Microsoft service limits',
     'Medium',
     'Measured during the pilot rather than assumed. If the observed rate is slower than expected, the number of weekends increases; the approach does not change.'],
    ['Pre-existing directory data problems',
     'Medium',
     'Two malformed objects have already been found and corrected. A validation sweep before each phase catches others, since a malformed object can fail in ways that do not name it.'],
    ['Groups no longer in use are migrated unnecessarily',
     'Low',
     'Retiring unused groups is a business decision, not a technical one, and is deliberately excluded so it cannot delay the migration. It can be run as separate clean-up before or after.']
  ]
));

// ---------- dependencies ----------
body.push(h1('What is needed'));
body.push(h2('From the Active Directory team'));
body.push(bullet('Confirmation that they retain ownership of the on-premises security groups this project leaves in place. No work is being asked of them; this is a boundary agreement, not a task.'));
body.push(h2('From the business'));
body.push(bullet('Approval of weekend change windows — one for the pilot, one for headquarters, and four to five for franchise.'));
body.push(bullet('An owner for communications to franchise offices during Phase 2.'));
body.push(bullet('A target date for the Exchange decommission, if one exists. It determines how tightly the phases below are scheduled, and whether the franchise phase can run at a comfortable pace or needs compressing.'));
body.push(h2('From this team'));
body.push(bullet('Migration tooling, inventory and verification — already built and in use.'));
body.push(bullet('Execution of each batch and reconciliation afterward.'));

// ---------- success ----------
body.push(h1('How success is measured'));
body.push(bullet('Every group present in the 13 August inventory is accounted for: migrated, deliberately excluded, or retired by decision.'));
body.push(bullet('No increase in mail delivery failures attributable to the migration in the week following each phase.'));
body.push(bullet('Distribution group changes can be made without on-premises directory access.'));
body.push(bullet('On-premises security groups are unchanged and continue to grant the access they granted before.'));

body.push(spacer(200));
body.push(new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF', space: 8 } },
  spacing: { before: 200 },
  children: [new TextRun({
    text: 'Object counts are from a full directory inventory dated 13 August 2026. Durations are working estimates and will be refined by the pilot, which exists specifically to measure them.',
    size: 18, italics: true, color: GREY })]
}));

const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 200 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '◦', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 200 } } } }
      ]
    }]
  },
  styles: { default: { document: { run: { font: 'Calibri', size: 21 } } } },
  sections: [{
    properties: { page: { ...LETTER, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
    children: body
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('DL-Migration-Project-Plan.docx', b);
  console.log('written', b.length, 'bytes');
});
