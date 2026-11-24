#!/usr/bin/env python3
"""Rebuild DEVIN-PACK.md (and its .txt twin) from the numbered source documents.

Run this after editing ANY numbered source document. The flat file and the
sources disagreeing has happened during this build more than once.

    python3 build-pack.py
"""
import hashlib
import pathlib

HERE = pathlib.Path(__file__).resolve().parent

# (filename, short description, cross-reference id)
DOCS = [
    ("00-BUILD-PROMPT.md",
     "The instruction. Role, environment reality, phases, gates, failure modes.", "00"),
    ("01-HARD-CONSTRAINTS.md",
     "The traps. Authoritative and non-negotiable.", "01"),
    ("02-ARCHITECTURE-SPEC.md",
     "Structure, config schemas, conventions, console layout.", "02"),
    ("03-SELF-TEST-SPEC.md",
     "The verifier. Build it first.", "03"),
    ("08-COMMON-CONTRACT.md",
     "Fixed function names every tool calls, plus a size budget.", "08"),
    ("04-REFERENCE-FILES.md",
     "How to use the reference files.", "04"),
    ("05-TOOL-SPECS.md",
     "Per-tool specs for the first pass, in build order, each with an exemplar "
     "and a done-when.", "05"),
    ("09-FULL-TOOL-CATALOGUE.md",
     "The destination: every remaining tool, the finished console layout, and "
     "the stages that get there.", "09"),
    ("06-TOOL-SKELETON.ps1",
     "A complete annotated tool in the required shape.", "06"),
    ("07-ACCEPTANCE-CHECKLIST.md",
     "For the OPERATOR, not the agent. Windows validation.", "07"),
]

RULE = "=" * 78


def separator(index: int, total: int, name: str, desc: str, ref: str) -> str:
    return (
        f"\n\n<br>\n\n{RULE}\n\n"
        f"# ◆ DOCUMENT {index} OF {total} — `{name}`\n\n"
        f"> {desc}\n>\n"
        f"> *Standalone document. Referred to elsewhere as `{ref}`.*\n\n"
        f"{RULE}\n\n"
    )


def main() -> None:
    total = len(DOCS)
    parts = [(HERE / "_PREAMBLE.md").read_text(encoding="utf-8").rstrip("\n")]

    for i, (name, desc, ref) in enumerate(DOCS, start=1):
        src = HERE / name
        if not src.exists():
            raise SystemExit(f"missing source document: {name}")
        body = src.read_text(encoding="utf-8").rstrip("\n")
        parts.append(separator(i, total, name, desc, ref))
        if name.endswith(".ps1"):
            parts.append(
                "*This document is a PowerShell file. Reproduce it as "
                f"`{name}`, not as markdown.*\n\n"
                "```powershell\n" + body + "\n```"
            )
        else:
            parts.append(body)

    parts.append(
        f"\n\n<br>\n\n{RULE}\n\n# END OF BUILD PACK\n\n"
        f"{total} documents. Reference code is in `DEVIN-REFERENCE-CODE.md` "
        "and `DEVIN-REFERENCE-CODE-2.md`.\n"
    )

    text = "".join(parts)
    for out in ("DEVIN-PACK.md", "DEVIN-PACK.txt"):
        (HERE / out).write_text(text, encoding="utf-8", newline="\n")

    digest = hashlib.md5(text.encode("utf-8")).hexdigest()
    print(f"wrote DEVIN-PACK.md and DEVIN-PACK.txt")
    print(f"  {total} documents, {len(text.splitlines()):,} lines, "
          f"{len(text.encode('utf-8')):,} bytes")
    print(f"  md5 {digest}")


if __name__ == "__main__":
    main()
