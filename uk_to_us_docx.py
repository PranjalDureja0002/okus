#!/usr/bin/env python3
"""
uk_to_us_docx.py

Finds UK English spellings in a Word (.docx) document using Claude Opus 5.5
(extra-high effort), replaces them with their US English equivalents, and saves
the document back in place, keeping all formatting.

Setup:
    pip install anthropic python-docx
    export ANTHROPIC_API_KEY="sk-ant-..."        # Windows: set ANTHROPIC_API_KEY=...

Usage:
    python uk_to_us_docx.py "Scoping Template.docx"            # fix in place (+ .bak backup)
    python uk_to_us_docx.py "Scoping Template.docx" --dry-run  # only report, don't save
    python uk_to_us_docx.py "Scoping Template.docx" --no-backup
"""

import argparse
import json
import re
import shutil
import sys

import anthropic
from docx import Document
from docx.oxml.ns import qn

MODEL = "claude-opus-5-5"
EFFORT = "xhigh"          # "extra high" reasoning
MAX_TOKENS = 64000        # room for thinking + answer
CHUNK_CHARS = 30000       # document text sent per request

W_P, W_T, W_LANG, W_VAL = qn("w:p"), qn("w:t"), qn("w:lang"), qn("w:val")
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

SYSTEM_PROMPT = """You are an expert copy editor converting British (UK) English to American (US) English.
The client requires ZERO UK English in the document.

You receive a JSON list of paragraphs, each with an "id" and "text".
Find every UK English spelling or UK-only word form and give its US English replacement. Cover, for example:
- -our -> -or (colour, behaviour, favour, labour, organisation's "honour")
- -ise/-isation/-yse -> -ize/-ization/-yze (organise, prioritise, utilisation, analyse, catalyse)
- -re -> -er (centre, metre, theatre, fibre)
- -ence -> -ense (licence, defence, offence, pretence)
- doubled consonants (travelled, cancelled, modelling, labelled, fulfil -> fulfill, enrol -> enroll)
- ae/oe forms (paediatric, oestrogen, manoeuvre)
- -ogue -> -og (catalogue, dialogue, analogue)
- other UK forms: programme -> program, cheque -> check, tyre -> tire, grey -> gray, judgement -> judgment,
  ageing -> aging, acknowledgement -> acknowledgment, whilst -> while, amongst -> among, learnt -> learned,
  spelt -> spelled, plough -> plow, storey -> story, kerb -> curb, aluminium -> aluminum, sceptical -> skeptical
- context-dependent ones: UK "practise" (verb) -> "practice"; UK "licence" (noun) -> "license"

Rules:
- Use context. Only flag a word if it is genuinely UK English in that sentence.
- Do NOT change proper nouns or official names (e.g. "Ministry of Defence", "Labour Party", "Centre Parcs"),
  titles of laws/standards/publications, URLs, email addresses, file names, or code.
- "uk" must be the word exactly as it appears in the text (same letters and case).
- "us" must keep the same capitalization pattern as the original.
- List each distinct UK word once per paragraph; every whole-word occurrence in that paragraph will be replaced.
- If there are none, return an empty list.

Reply with ONLY this JSON, no other text:
{"changes": [{"id": <paragraph id>, "uk": "<UK word>", "us": "<US word>"}]}"""


# ---------------------------------------------------------------- reading the docx

def _owning_paragraph(node):
    el = node.getparent()
    while el is not None and el.tag != W_P:
        el = el.getparent()
    return el


def _xml_roots(doc):
    """Body plus every header/footer that has its own content."""
    roots, seen = [doc.element.body], set()
    for section in doc.sections:
        for hf in (section.header, section.footer,
                   section.first_page_header, section.first_page_footer,
                   section.even_page_header, section.even_page_footer):
            if not hf.is_linked_to_previous:
                el = hf._element
                if id(el) not in seen:
                    seen.add(id(el))
                    roots.append(el)
    return roots


def collect_paragraphs(doc):
    """Every paragraph (body, tables, nested tables, text boxes, headers, footers)
    as (list_of_w:t_nodes, text). Works on raw text nodes so hyperlinks are included
    and words split across formatting runs are still found."""
    paragraphs = []
    for root in _xml_roots(doc):
        for p in root.iter(W_P):
            nodes = [t for t in p.iter(W_T) if _owning_paragraph(t) is p]
            text = "".join(t.text or "" for t in nodes)
            if text.strip():
                paragraphs.append((nodes, text))
    return paragraphs


# ---------------------------------------------------------------- asking Claude

def make_chunks(paragraphs):
    chunk, size = [], 0
    for i, (_, text) in enumerate(paragraphs):
        if chunk and size + len(text) > CHUNK_CHARS:
            yield chunk
            chunk, size = [], 0
        chunk.append({"id": i, "text": text})
        size += len(text)
    if chunk:
        yield chunk


def ask_claude(client, chunk):
    # Streaming is required for large max_tokens; adaptive thinking is always on for Opus 5.5.
    with client.messages.stream(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        output_config={"effort": EFFORT},
        messages=[{"role": "user", "content": json.dumps(chunk, ensure_ascii=False)}],
    ) as stream:
        message = stream.get_final_message()

    reply = "".join(b.text for b in message.content if b.type == "text")
    start, end = reply.find("{"), reply.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"Model did not return JSON:\n{reply[:500]}")
    return json.loads(reply[start:end + 1]).get("changes", [])


# ---------------------------------------------------------------- replacing text

def _match_case(original, replacement):
    if original.isupper():
        return replacement.upper()
    if original.islower():
        return replacement.lower()
    if original[:1].isupper():
        return replacement[:1].upper() + replacement[1:].lower()
    return replacement


def _replace_span(nodes, start, end, new_text):
    """Replace characters [start, end) of the paragraph's combined text,
    even if they are spread over several runs. Formatting of the first run is kept."""
    pos, first, last = 0, None, None
    bounds = []
    for i, t in enumerate(nodes):
        length = len(t.text or "")
        bounds.append(pos)
        if first is None and start < pos + length:
            first = i
        if end <= pos + length:
            last = i
            break
        pos += length

    a, b = nodes[first], nodes[last]
    a_text, b_text = a.text or "", b.text or ""
    if first == last:
        a.text = a_text[:start - bounds[first]] + new_text + a_text[end - bounds[first]:]
    else:
        a.text = a_text[:start - bounds[first]] + new_text
        for t in nodes[first + 1:last]:
            t.text = ""
        b.text = b_text[end - bounds[last]:]
        b.set(XML_SPACE, "preserve")
    a.set(XML_SPACE, "preserve")


def apply_change(nodes, uk, us):
    text = "".join(t.text or "" for t in nodes)
    pattern = re.compile(r"(?<![A-Za-z])" + re.escape(uk) + r"(?![A-Za-z])", re.IGNORECASE)
    matches = list(pattern.finditer(text))
    for m in reversed(matches):  # right to left keeps earlier offsets valid
        _replace_span(nodes, m.start(), m.end(), _match_case(m.group(0), us))
    return len(matches)


def set_language_us(doc):
    """Switch Word's proofing language from en-GB to en-US so spell-check agrees."""
    roots = [doc.styles.element] + _xml_roots(doc)
    for root in roots:
        for lang in root.iter(W_LANG):
            if (lang.get(W_VAL) or "").lower().startswith("en-"):
                lang.set(W_VAL, "en-US")


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description="Replace UK English with US English in a .docx using Claude.")
    ap.add_argument("docx", help="Path to the .docx file (edited in place)")
    ap.add_argument("--dry-run", action="store_true", help="Report changes without saving")
    ap.add_argument("--no-backup", action="store_true", help="Don't write a .bak copy first")
    ap.add_argument("--keep-language", action="store_true", help="Don't switch proofing language to en-US")
    args = ap.parse_args()

    if not args.docx.lower().endswith(".docx"):
        sys.exit("Only .docx files are supported (save .doc files as .docx first).")

    doc = Document(args.docx)
    paragraphs = collect_paragraphs(doc)
    print(f"Read {len(paragraphs)} paragraphs from {args.docx}")

    client = anthropic.Anthropic()
    changes = []
    chunks = list(make_chunks(paragraphs))
    for n, chunk in enumerate(chunks, 1):
        print(f"Checking part {n}/{len(chunks)} with {MODEL} (effort={EFFORT})...")
        changes.extend(ask_claude(client, chunk))

    total, log = 0, []
    for c in changes:
        try:
            pid, uk, us = int(c["id"]), str(c["uk"]).strip(), str(c["us"]).strip()
        except (KeyError, TypeError, ValueError):
            continue
        if not (0 <= pid < len(paragraphs)) or not uk or not us or uk.lower() == us.lower():
            continue
        nodes, _ = paragraphs[pid]
        if args.dry_run:
            count = len(re.findall(r"(?<![A-Za-z])" + re.escape(uk) + r"(?![A-Za-z])",
                                   "".join(t.text or "" for t in nodes), re.IGNORECASE))
        else:
            count = apply_change(nodes, uk, us)
        if count:
            total += count
            log.append((uk, us, count))

    if not log:
        print("No UK English found. Document unchanged.")
        return

    print(f"\n{'UK':<25}{'US':<25}Count")
    summary = {}
    for uk, us, count in log:
        key = (uk.lower(), us.lower())
        summary[key] = summary.get(key, 0) + count
    for (uk, us), count in sorted(summary.items()):
        print(f"{uk:<25}{us:<25}{count}")
    print(f"\nTotal replacements: {total}")

    if args.dry_run:
        print("Dry run: document not saved.")
        return

    if not args.keep_language:
        set_language_us(doc)
    if not args.no_backup:
        shutil.copy2(args.docx, args.docx + ".bak")
        print(f"Backup saved to {args.docx}.bak")
    doc.save(args.docx)
    print(f"Saved changes in place to {args.docx}")


if __name__ == "__main__":
    main()
