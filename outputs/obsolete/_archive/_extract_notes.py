#!/usr/bin/env python3
"""
PowerShell comment extractor.

Splits a .ps1/.psm1/.psd1 into (a) code with commentary removed and (b) a
markdown notes file holding everything that was removed, with the function or
region each comment belonged to recorded alongside it.

This is a state machine rather than a regex pass because '#' is only a comment
in one of several possible lexical states. The cases that actually bite:

  * Here-strings.  @' ... '@ and @" ... "@ hold the embedded HTML/CSS/JS for
    the M365_View_* files.  Those contain CSS hex colours (#1f2937, #fff) and
    JS fragments.  A regex that treats '#' as a comment marker deletes them
    and produces a broken page with no syntax error to warn you.
  * Quoted strings.  '#' inside '...' or "..." is data.
  * Nested block comments.  PowerShell permits <# <# #> #>, so block comments
    need a depth counter, not a non-greedy match.

Preserved deliberately:
  * #Requires   -- a parser directive, not a comment.  Removing it changes
                   behaviour.
  * #region / #endregion -- editor folding structure; navigational, and they
                   carry no reasoning worth withholding.
  * Anything matching KEEP_PATTERNS -- terse guardrail warnings.
"""

import os
import re
import sys
import json

KEEP_PATTERNS = [
    re.compile(r'^\s*#\s*(GUARD|SAFETY|DO NOT REMOVE)', re.I),
]

NORMAL, SQ, DQ, HS_SQ, HS_DQ, BLOCK = range(6)


def scan(src):
    """Return (code_chars, comments) where comments is a list of dicts."""
    out = []
    comments = []
    state = NORMAL
    depth = 0
    i = 0
    line = 1
    n = len(src)
    buf = []          # current comment text
    buf_line = 0

    def at_line_start(pos):
        j = pos - 1
        while j >= 0 and src[j] in ' \t':
            j -= 1
        return j < 0 or src[j] == '\n'

    while i < n:
        c = src[i]
        nxt = src[i + 1] if i + 1 < n else ''

        if c == '\n':
            line += 1

        if state == NORMAL:
            # here-string openers must be the last thing on their line
            if c == '@' and nxt in ("'", '"'):
                j = i + 2
                while j < n and src[j] in ' \t\r':
                    j += 1
                if j < n and src[j] == '\n':
                    state = HS_SQ if nxt == "'" else HS_DQ
                    out.append(c)
                    out.append(nxt)
                    i += 2
                    continue
            if c == '<' and nxt == '#':
                state = BLOCK
                depth = 1
                buf = []
                buf_line = line
                i += 2
                continue
            if c == '#':
                # only a comment when it opens a token, otherwise it is data
                # such as a CSS colour reached through string interpolation
                prev = src[i - 1] if i else '\n'
                if prev in ' \t\n\r;{(|,=' or i == 0:
                    j = src.find('\n', i)
                    if j == -1:
                        j = n
                    text = src[i:j]
                    if (re.match(r'\s*#\s*requires', text, re.I)
                            or re.match(r'\s*#\s*(end)?region', text, re.I)
                            or any(p.match(text) for p in KEEP_PATTERNS)):
                        out.append(text)
                    else:
                        comments.append({'line': line, 'kind': 'line',
                                         'text': text.lstrip().lstrip('#').strip()})
                        # drop the whole physical line if nothing precedes it
                        if at_line_start(i):
                            k = len(out) - 1
                            while k >= 0 and out[k] in ' \t':
                                out.pop()
                                k -= 1
                            if j < n:
                                j += 1        # eat the newline too
                                line += 1
                    i = j
                    continue
                out.append(c)
                i += 1
                continue
            if c == "'":
                state = SQ
            elif c == '"':
                state = DQ
            out.append(c)
            i += 1
            continue

        if state == BLOCK:
            if c == '<' and nxt == '#':
                depth += 1
                buf.append('<#')
                i += 2
                continue
            if c == '#' and nxt == '>':
                depth -= 1
                i += 2
                if depth == 0:
                    comments.append({'line': buf_line, 'kind': 'block',
                                     'text': ''.join(buf).strip()})
                    state = NORMAL
                    # swallow trailing newline left behind by the block
                    if i < n and src[i] == '\n':
                        i += 1
                        line += 1
                    while out and out[-1] in ' \t':
                        out.pop()
                    continue
                buf.append('#>')
                continue
            buf.append(c)
            i += 1
            continue

        # inside some kind of string: copy verbatim
        out.append(c)
        if state == SQ and c == "'":
            state = NORMAL if nxt != "'" else SQ
            if nxt == "'":
                out.append(nxt)
                i += 1
        elif state == DQ:
            if c == '`':
                if i + 1 < n:
                    out.append(nxt)
                    i += 1
            elif c == '"':
                state = NORMAL
        elif state in (HS_SQ, HS_DQ):
            term = "'@" if state == HS_SQ else '"@'
            if c == '\n' and src[i + 1:i + 3] == term:
                out.append(term)
                state = NORMAL
                i += 2
        i += 1

    return ''.join(out), comments


def tidy(code):
    """Collapse the blank-line runs left behind by removed comments."""
    lines = code.split('\n')
    res = []
    blanks = 0
    for L in lines:
        if L.strip() == '':
            blanks += 1
            if blanks > 1:
                continue
        else:
            blanks = 0
        res.append(L.rstrip())
    while res and res[0].strip() == '':
        res.pop(0)
    while res and res[-1].strip() == '':
        res.pop()
    return '\n'.join(res) + '\n'


def context_map(src):
    """line number -> nearest preceding function/region name."""
    m = {}
    cur = '(top of file)'
    for idx, L in enumerate(src.split('\n'), start=1):
        fn = re.match(r'\s*function\s+([\w\-]+)', L, re.I)
        rg = re.match(r'\s*#region\s+(.*)', L, re.I)
        if fn:
            cur = 'function ' + fn.group(1)
        elif rg:
            cur = 'region ' + rg.group(1).strip(' -')
        m[idx] = cur
    return m


HEADER = """#Requires -Version {ver}
<#
    {name}
    {purpose}

    Author   : Glenn Armbruster
    Revised  : {rev}
#>
"""


def process(path, notes_dir, purpose_map, write=True):
    raw = open(path, encoding='utf-8-sig', errors='replace').read()
    name = os.path.basename(path)
    ctx = context_map(raw)
    code, comments = scan(raw)
    code = tidy(code)

    # notes file
    lines = ['# {} -- design notes'.format(name), '',
             '> Private. Not distributed with the scripts.', '']
    if not comments:
        lines.append('_No commentary was present in this file._')
    # Consecutive line-comments are almost always one wrapped sentence, so
    # merge runs of adjacent lines into a single entry rather than emitting a
    # bullet per physical line and cutting sentences in half.
    merged = []
    for c in comments:
        if (merged and c['kind'] == 'line' and merged[-1]['kind'] == 'line'
                and c['line'] == merged[-1]['end'] + 1):
            merged[-1]['text'] = (merged[-1]['text'].rstrip() + ' '
                                  + c['text'].lstrip())
            merged[-1]['end'] = c['line']
        else:
            c = dict(c)
            c['end'] = c['line']
            merged.append(c)

    last_ctx = None
    for c in merged:
        where = ctx.get(c['line'], '(top of file)')
        if where != last_ctx:
            lines += ['', '## ' + where, '']
            last_ctx = where
        body = re.sub(r'\s+', ' ', c['text']).strip() if c['kind'] == 'line' \
            else c['text']
        if c['kind'] == 'block':
            lines += ['<!-- line {} -->'.format(c['line']), '', '```text',
                      body, '```', '']
        else:
            ref = ('L{}'.format(c['line']) if c['end'] == c['line']
                   else 'L{}-{}'.format(c['line'], c['end']))
            lines.append('- **{}** {}'.format(ref, body))
    notes = '\n'.join(lines).rstrip() + '\n'

    if write:
        open(os.path.join(notes_dir, name.rsplit('.', 1)[0] + '.notes.md'),
             'w', encoding='utf-8').write(notes)
    return code, notes, len(comments)


if __name__ == '__main__':
    print('module -- import and call process()')
