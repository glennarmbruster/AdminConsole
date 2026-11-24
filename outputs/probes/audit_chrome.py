#!/usr/bin/env python3
"""
Static consistency audit of the AdminConsole WinForms tools.

READ ONLY. Reports; changes nothing.

WHAT IT CAN AND CANNOT SEE. This reads source, not a running form. Controls
positioned with literal numbers are checked; controls positioned from variables,
built in loops, or laid out by a container are reported as unverifiable rather
than silently passed. "Not checked" is an honest answer; "no findings" from a
scan that skipped half the file is not.

Comments are stripped, STRING LITERALS ARE NOT. Cmdlet and control names live
inside strings all over this toolkit, and a scanner that strips strings goes
blind exactly where it matters.
"""
import glob, os, re, sys
from collections import defaultdict

DIR = sys.argv[1] if len(sys.argv) > 1 else '.'
SKIP = {'Common.ps1', 'Publish-Toolkit.ps1'}

def strip_comments(s):
    s = re.sub(r'(?s)<#.*?#>', lambda m: '\n' * m.group(0).count('\n'), s)
    out = []
    for line in s.split('\n'):
        # Blank the comment, keep the line, so line numbers stay honest.
        q = None; cut = None
        for i, ch in enumerate(line):
            if q:
                if ch == q: q = None
                continue
            if ch in '"\'': q = ch; continue
            if ch == '#': cut = i; break
        out.append(line if cut is None else line[:cut])
    return '\n'.join(out)

findings = defaultdict(list)          # file -> [(rule, detail)]
stats = {}

for path in sorted(glob.glob(os.path.join(DIR, '*.ps1'))):
    f = os.path.basename(path)
    if f in SKIP: continue
    raw = open(path, encoding='utf-8-sig', errors='replace').read().replace('\r\n', '\n')
    code = strip_comments(raw)
    if 'System.Windows.Forms.Form' not in code: continue
    lines = code.split('\n')

    def add(rule, detail): findings[f].append((rule, detail))

    # ---- which variables are Forms -------------------------------------
    forms = re.findall(r'\$(\w+)\s*=\s*New-Object\s+System\.Windows\.Forms\.Form\b', code)
    # the form actually shown at the end of the script
    shown = re.findall(r'\$(\w+)\.ShowDialog', code)
    main = shown[-1] if shown else (forms[0] if forms else None)
    if not main: continue
    stats[f] = main

    # ---- CHR001  help button -------------------------------------------
    if 'Add-ToolHelpButton' not in code:
        add('CHR001', 'no Add-ToolHelpButton -- this window has no "?"')

    # ---- CHR002  theme --------------------------------------------------
    if 'Set-ConsoleTheme' not in code:
        add('CHR002', 'no Set-ConsoleTheme -- window will not match the tenant palette')

    # ---- CHR003  screen corner -----------------------------------------
    if 'Set-FormScreenCorner' not in code and not re.search(r'\$'+main+r'\.Location\s*=', code):
        add('CHR003', 'no Set-FormScreenCorner and no manual Location -- opens wherever Windows decides')

    # ---- CHR004  Size vs ClientSize ------------------------------------
    has_size = re.search(r'\$'+main+r'\.Size\s*=', code)
    has_client = re.search(r'\$'+main+r'\.ClientSize\s*=', code)
    if has_size and has_client:
        a = has_size.start(); b = has_client.start()
        add('CHR004', 'sets BOTH Size and ClientSize; ClientSize wins if it is later '
                      f'({"ClientSize is later -- Size is dead code" if b > a else "Size is later -- overrides ClientSize"})')

    # ---- CHR005  font ---------------------------------------------------
    m = re.search(r'\$'+main+r'\.Font\s*=\s*New-Object\s+System\.Drawing\.Font\(([^)]*)\)', code)
    if not m:
        add('CHR005', 'no form Font set -- inherits the Windows default, which differs from the other tools')
    else:
        stats[f + '::font'] = m.group(1).strip()

    # ---- CHR006  a way out ----------------------------------------------
    if not re.search(r'(?i)\.Text\s*=\s*[\'"]\s*(close|cancel|exit|done)\s*[\'"]', code):
        add('CHR006', 'no Close/Cancel/Exit button found -- title-bar X only')

    # ---- collect control geometry, grouped by PARENT ---------------------
    kinds = dict(re.findall(r'\$(\w+)\s*=\s*New-Object\s+System\.Windows\.Forms\.(\w+)', code))
    loc, size, line_of = {}, {}, {}
    for i, ln in enumerate(lines):
        m = re.search(r'\$(\w+)\.Location\s*=\s*New-Object\s+System\.Drawing\.Point\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', ln)
        if m: loc[m.group(1)] = (int(m.group(2)), int(m.group(3))); line_of[m.group(1)] = i + 1
        m = re.search(r'\$(\w+)\.Size\s*=\s*New-Object\s+System\.Drawing\.Size\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)', ln)
        if m: size[m.group(1)] = (int(m.group(2)), int(m.group(3)))

    parent = {}
    add_line = {}
    for m in re.finditer(r'\$(\w+)\.Controls\.Add\(\s*\$(\w+)\s*\)', code):
        parent[m.group(2)] = m.group(1)
        add_line[m.group(2)] = code[:m.start()].count('\n') + 1
    for m in re.finditer(r'\$(\w+)\.Controls\.AddRange\(\s*@\(([^)]*)\)', code):
        for v in re.findall(r'\$(\w+)', m.group(2)):
            parent[v] = m.group(1)
            add_line[v] = code[:m.start()].count('\n') + 1

    # A CONTAINER VARIABLE CAN BE REUSED FOR SEVERAL DIFFERENT WINDOWS.
    # ToolManager declares $dlg twice, 670 lines apart, for two unrelated
    # dialogs. Treating both as one parent made every control of the first
    # "overlap" every control of the second -- eleven confident findings, all
    # imaginary. So each container gets bucketed by WHICH declaration of that
    # variable precedes the control being added to it.
    decl_lines = defaultdict(list)
    for m in re.finditer(r'\$(\w+)\s*=\s*New-Object\s+System\.Windows\.Forms\.(?:Form|GroupBox|Panel|TabPage|TableLayoutPanel|FlowLayoutPanel|SplitContainer)\b', code):
        decl_lines[m.group(1)].append(code[:m.start()].count('\n') + 1)

    def instance_of(container, at):
        ds = decl_lines.get(container)
        if not ds or len(ds) == 1: return container
        prior = [d for d in ds if d <= at]
        return f'{container}#{max(prior) if prior else ds[0]}'

    # Controls hidden at build time are not on screen at the same moment as
    # the control they sit on top of. Deliberately stacking a Cancel button
    # exactly over a Run button and toggling Visible is a normal, good WinForms
    # pattern -- five tools here do it -- and reporting it as a collision
    # trains the reader to ignore the check.
    hidden = set(re.findall(r'\$(\w+)\.Visible\s*=\s*\$false', code))

    # ---- CHR007  overlap, within one parent only ------------------------
    # Coordinates are PARENT-RELATIVE. Comparing a GroupBox child against a
    # form-level control is how a flat scan invents collisions that do not
    # exist -- the bug in my own first attempt at this.
    groups = defaultdict(list)
    for v, p in parent.items():
        if v in loc and v in size:
            groups[instance_of(p, add_line.get(v, 0))].append(v)

    for p, members in groups.items():
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                a, b = members[i], members[j]
                if a in hidden or b in hidden: continue
                ax, ay = loc[a]; aw, ah = size[a]
                bx, by = loc[b]; bw, bh = size[b]
                ox = min(ax + aw, bx + bw) - max(ax, bx)
                oy = min(ay + ah, by + bh) - max(ay, by)
                if ox > 0 and oy > 0:
                    # A Label overlapped by a few px is usually harmless; a
                    # Button under something is a dead click. Report area so
                    # the reader can judge rather than the scanner guessing.
                    ka, kb = kinds.get(a, '?'), kinds.get(b, '?')
                    add('CHR007', f'${a} ({ka}) and ${b} ({kb}) overlap by {ox}x{oy}px '
                                  f'inside ${p}  [lines {line_of.get(a,"?")}, {line_of.get(b,"?")}]')

    # ---- CHR008  control past the form's right/bottom edge --------------
    dim = has_client or has_size
    if dim:
        mm = re.search(r'Size\(\s*(\d+)\s*,\s*(\d+)\s*\)', code[dim.start():dim.start() + 200])
        if mm:
            fw, fh = int(mm.group(1)), int(mm.group(2))
            approx = False
            if has_size and not has_client:
                # Form.Size includes the chrome, and how much chrome depends on
                # FormBorderStyle, the Windows theme and DPI. 16x39 is right for
                # a Sizable form at 100%, but it is an ESTIMATE, so a small
                # overhang is reported as "verify" rather than asserted. A
                # scanner that states a 5px clip as fact will be wrong often
                # enough that its real findings stop being believed.
                fw -= 16; fh -= 39; approx = True
            tol = 12 if approx else 2
            for v, p in parent.items():
                if p != main or v not in loc or v not in size: continue
                x, y = loc[v]; w, h = size[v]
                over_x = (x + w) - fw
                over_y = (y + h) - fh
                if over_x > tol or over_y > tol:
                    add('CHR008', f'${v} ({kinds.get(v,"?")}) ends at {x+w},{y+h}, '
                                  f'client area about {fw}x{fh} -- over by {max(over_x,0)}x{max(over_y,0)}px'
                                  f'  [line {line_of.get(v,"?")}]')
                elif over_x > 0 or over_y > 0:
                    add('CHR009', f'${v} ({kinds.get(v,"?")}) sits {max(over_x,0)}x{max(over_y,0)}px past the '
                                  f'estimated client edge -- within the margin of error for border and title bar, '
                                  f'so eyeball it  [line {line_of.get(v,"?")}]')

    # ---- how much of this file was actually checkable --------------------
    placed = len([v for v in parent if v in loc and v in size])
    total = len([v for v in parent])
    if total and placed / total < 0.6:
        add('INFO', f'only {placed} of {total} controls have literal Location+Size; '
                    'the rest are positioned dynamically and were NOT geometry-checked')

# ---------------------------------------------------------------------------
print('=' * 78)
print('CHROME CONSISTENCY AUDIT -- read only')
print('=' * 78)

fonts = defaultdict(list)
for k, v in stats.items():
    if k.endswith('::font'): fonts[v].append(k[:-6])
print('\nFORM FONTS IN USE')
for fnt, files in sorted(fonts.items(), key=lambda kv: -len(kv[1])):
    print(f'  {len(files):3}x  {fnt}')
    if len(files) <= 6:
        for x in files: print(f'          {x}')

order = ['CHR001','CHR002','CHR003','CHR004','CHR005','CHR006','CHR007','CHR008','CHR009','INFO']
names = {
 'CHR001':'no help button','CHR002':'no theme','CHR003':'no window placement',
 'CHR004':'Size and ClientSize both set','CHR005':'no form font','CHR006':'no close button',
 'CHR007':'controls overlap (same parent)','CHR008':'control past the form edge','CHR009':'near the form edge -- verify visually','INFO':'coverage note'}
by_rule = defaultdict(list)
for f, items in findings.items():
    for rule, detail in items: by_rule[rule].append((f, detail))

for rule in order:
    hits = by_rule.get(rule, [])
    if not hits: continue
    print(f'\n{rule}  {names[rule]}  ({len(hits)})')
    print('-' * 78)
    for f, d in sorted(hits): print(f'  {f:30} {d}')

print(f'\nforms scanned: {len([k for k in stats if not k.endswith("::font")])}, '
      f'files with findings: {len(findings)}')
