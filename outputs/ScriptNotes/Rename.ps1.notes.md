# Rename.ps1

Packages an agnostic toolkit for email and unpacks it at the far end.
`FH-Agnostic/` and `CloudOnly/` only — the hybrid console is not shipped.

```
1  Package    append .txt to every blocked extension, write a manifest
2  Restore    strip .txt, verify every file against the manifest
3  Unblock    clear the mark-of-the-web (Windows only)
```

Runs on Windows PowerShell 5.1 and on PowerShell 7 on macOS. Modes 1 and 2 are
cross-platform; mode 3 is a no-op anywhere but Windows, because the zone marker
it clears is an NTFS alternate data stream.

**It renames everything except itself and the manifest** — rename `Rename.ps1`
to `Rename.txt` by hand going out, and back on arrival.

## What travels

`reports`, `Logs` and `captures` are **output**. They are regenerated wherever
the toolkit lands, they can be large, and one of them may hold real mailbox
data — so none of it belongs in an email. `Config\` **does** travel: it carries
the tool registry and layout, which is the toolkit's shape rather than its
results (§19).

Only extensions a gateway is likely to strip get renamed. Everything else is
still hashed into the manifest; renaming a `.png` buys nothing and costs a round
trip.

Appending `.txt` rather than replacing the extension makes reversal
deterministic with no state, and stops `Common.ps1` and `Common.md` colliding
into one name. Re-running is idempotent — a file already carrying `.txt` over a
blocked extension is recorded under its original name and left alone, so it
never produces `Common.ps1.txt.txt`.

## Why there is a manifest

The rename is the easy part. **The real risk is a gateway silently rewriting a
`.txt` in transit** — a line-ending conversion or a re-encode strips the CRLF or
the BOM, and you find out weeks later when `#Requires` breaks in the ISE (§12).

So every file is hashed (SHA-256) on the way out and verified on the way back,
and every restored `.ps1` is additionally checked for a single UTF-8 BOM and
CRLF endings. Corruption becomes an error instead of a mystery.

A missing manifest is **not fatal**: stripping `.txt` still recovers a working
toolkit. The operator just loses the integrity check, and is told so rather than
left to assume it passed.

## Details worth remembering

`$IsWindows` only exists in PowerShell 6+. In 5.1 it is undefined, and 5.1 only
runs on Windows, so undefined means Windows.

Manifest paths are normalised to backslash on write and translated back on read,
because a package built on macOS has to verify on Windows.

`.DS_Store` is removed rather than merely skipped. It is invisible on a Mac,
visible and confusing on Windows, and **it has already shipped once in this
project** inside a folder that was supposed to be clean.

## Not the first choice

Transport preference is: share link → an approved path from security → zip →
this. The manifest verifies a SharePoint download or an unzipped folder just as
well, since it checks bytes rather than transport.
