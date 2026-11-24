# HeaderAnalyzer.ps1

Paste raw internet message headers, get the delivery path with per-hop delays,
the authentication verdicts, and the Microsoft antispam headers decoded.

Built 2026-08-06 as a trial — a handful of header analyses a month is thin
justification, so this is deliberately narrow and gets a button on the
understanding that it either earns it or gets removed.

**Offline. Nothing leaves the machine**, which is the main reason it exists.

---

## Why not just use MXToolbox

Two reasons, one of which is the whole argument.

**Headers get pasted into a third-party website.** A full header set carries
internal hostnames, private IPs, the recipient's address, the subject line and
internal message IDs. Across 800 franchise offices that is a data-handling
question somebody could reasonably ask, and "it's only a diagnostic tool" is a
weak answer. Offline makes the question disappear.

**The generic tools treat the Microsoft headers as opaque.** MXToolbox renders
the Received chain well and shows `X-Forefront-Antispam-Report` as a wall of
text. For an Exchange admin that wall *is* the answer — SCL, SFV, CAT and
compauth are what decide whether a message was junked and why. Decoding them is
the only part worth writing ourselves.

Deliberately **not** included: DNSBL aggregation. Checking a hundred blocklists
is precisely what MXToolbox is for, and reimplementing it is a maintenance
treadmill with no end.

## Every decoded value was read, not recalled

Transcribed from Microsoft Learn,
[Anti-spam message headers](https://learn.microsoft.com/en-us/defender-office-365/message-headers-eop-mdo),
page updated 2025-10-08, read 2026-08-06. Checking it changed the code twice:

- **`SFV:SFE`** — filtering skipped because the sender is on the recipient's
  Safe Senders list. Would have been missed, and it is one of the more common
  "why did this get through" answers.
- **`PCL`** — phishing confidence level — **is not in the current header set.**
  Writing it from memory would have produced a decoder for a field that never
  appears. That is the worst kind of wrong: it looks like working code forever,
  because nothing ever exercises it.

When a value renders as `unknown ... re-read the Microsoft Learn page`, that is
the tool saying Microsoft changed something — not that the parser broke.

## Unfolding is the whole parsing job

A header continues onto the next line whenever that line begins with a space or
tab (RFC 5322 §2.2.3). `Received:` headers are folded across three or four lines
essentially always, **and the date is on the last fold**.

Parse line-by-line without unfolding and you get a Received chain with no dates
in it — so no hop timings, which is half the point of the tool. It fails
plausibly rather than loudly.

Duplicates are kept in a list, not a hashtable. `Received` appears once per hop
and `Authentication-Results` can appear more than once (an on-premises hop and
Microsoft 365 both stamp one); a hashtable keeps exactly one of each and the one
it keeps is arbitrary.

The blank line ends the header block. Anything after it is body, however much it
looks like a header — the test file has a fake `Received:` line down there to
prove it is ignored.

## Hop order is backwards on purpose

Each hop **prepends** its `Received` header, so the first one in the file is the
**last** hop. The chain is reversed before display so it reads in the direction
the message travelled.

Delays are the difference between consecutive servers' **own clocks**, and the
report says so. A server with a bad clock produces a negative or absurd delay,
and that is the clock, not the message. Total transit is labelled a floor rather
than a total when any hop's date failed to parse.

## Two bugs the test found

Written, then exercised against a realistic Microsoft 365 header set before
delivery. Both of these were live:

**IPv6 in parentheses, not brackets.** Microsoft writes `(2603:10b6:5:1f5::8)`.
The original pattern only matched `[...]`, so every hop inside Exchange Online —
most of them — silently reported no IP.

**A version string that looks like an address.** `id 15.20.771.19` in the `id`
clause matches an IPv4 shape exactly. On a header with no connecting IP at all,
the tool would have reported the Exchange build number as the sender's IP: a
plausible, confident, completely wrong answer.

Fixed by searching **only the text before `by`** — the connecting IP is in the
`from` clause by definition — and range-checking the octets. Both fixes are
regression-tested in the notes below.

## What it flags

Only things that are worth a second look, never a verdict:

- From / Return-Path domain mismatch (normal for mailing lists, also what a
  spoof looks like)
- Reply-To pointing somewhere other than From
- SPF fail or softfail, DKIM fail, DMARC fail, compauth fail
- More than 12 hops — forwarding or a list, and every extra hop is another place
  DKIM can break
- `bestguesspass` is called out explicitly, because it reads like a pass and
  means *there is no DMARC record*

When nothing trips, it says "nothing in the headers contradicts itself" rather
than "clean". The tool cannot see the body, the attachments or the links.

## Adding the button

| Field | Value |
|---|---|
| Name | `Header Analyzer` |
| Script | `.\HeaderAnalyzer.ps1` |
| Section | anywhere with room |
| Own window | unticked |
| ObjectTypes | empty |

It touches no directory, so it is never greyed out and needs no
`TOOLKIT-REQUIRES` declaration.

## Still open

- **Received parsing is best-effort.** RFC 5321 defines the tokens but not the
  order, and every MTA differs. Each clause is matched independently so a
  missing one costs a field rather than the whole hop, but appliances that
  format creatively will still produce gaps.
- **No DNS lookups.** SPF/DMARC records could be retrieved with `Resolve-DnsName`
  to show what the *current* policy is versus what was evaluated at send time.
  Genuinely useful for "why did this fail" — and it is the first thing to add if
  the tool earns its button.
- **Adjacent encoded-words** should have the whitespace between them removed per
  RFC 2047. Not implemented; the effect is an occasional extra space in a
  subject line.
- **No report file.** Output goes to the window and the clipboard. Header
  analysis is a look-at-it-now activity, and writing customer headers to disk
  would partly undo the reason for building this.
