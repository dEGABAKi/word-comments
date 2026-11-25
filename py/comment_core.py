"""Word native comment (OOXML) core: read / add / delete, with copy-on-lock.

Used by the DSH word-comments companion plugin; also runnable by hand:

    python comment_core.py read  <docx>
    python comment_core.py add   <docx>  <json-items-file>
    python comment_core.py delete <docx> <ids-comma|all>

Prints a JSON object on stdout. Never touches a document that is open in
Word/WPS: writes go to a sibling copy named "<name>-批注版.docx".
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import copy
import datetime as dt

from docx import Document
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"

COMMENT_SUFFIX = "-批注版"


class CoreError(Exception):
    pass


# ---------------------------------------------------------------- helpers

def _is_locked(path: str) -> bool:
    """True when another application holds the file open for writing."""
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except FileNotFoundError:
        raise CoreError(f"file not found: {path}")
    except OSError:
        return True


def _copy_target(path: str) -> str:
    base, ext = os.path.splitext(path)
    candidate = f"{base}{COMMENT_SUFFIX}{ext}"
    n = 2
    while os.path.exists(candidate):
        candidate = f"{base}{COMMENT_SUFFIX}({n}){ext}"
        n += 1
    shutil.copyfile(path, candidate)
    return candidate


def _write_target(path: str, copy: bool | None) -> tuple[str, bool]:
    """Return (target_path, copied). copy=None -> auto (copy only if locked)."""
    if not os.path.exists(path):
        raise CoreError(f"file not found: {path}")
    if copy is None:
        copy = _is_locked(path)
    if copy:
        return _copy_target(path), True
    if _is_locked(path):
        return _copy_target(path), True
    return path, False


def _iter_body_paragraphs(doc):
    """Every w:p in the main document body, in document order (incl. tables)."""
    body = doc.element.body
    for p in body.iter(qn("w:p")):
        yield p


def _paragraph_at(doc, para_no: int):
    """Wrap the para_no-th body paragraph (1-based, tables included) as Paragraph."""
    from docx.text.paragraph import Paragraph

    if para_no < 1:
        raise CoreError("para must be >= 1")
    current = 0
    for p in _iter_body_paragraphs(doc):
        current += 1
        if current == para_no:
            # Paragraph's parent only needs to expose the part; the body works.
            return Paragraph(p, doc._body)
    raise CoreError(f"para {para_no} out of range (document body has {current} paragraphs)")


def _p_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _anchor_map(doc) -> dict[int, dict]:
    """comment id -> {"para": 1-based body paragraph index, "text": anchored text}."""
    anchors: dict[int, dict] = {}
    active: dict[int, list] = {}
    for idx, p in enumerate(_iter_body_paragraphs(doc), start=1):
        for node in p.iter():
            tag = node.tag
            if tag == qn("w:commentRangeStart"):
                cid = int(node.get(qn("w:id")))
                active[cid] = [idx, []]
            elif tag == qn("w:commentRangeEnd"):
                cid = int(node.get(qn("w:id")))
                if cid in active:
                    start_para, chunks = active.pop(cid)
                    anchors[cid] = {"para": start_para, "text": "".join(chunks)}
            elif tag == qn("w:t") and active:
                for cid, buf in active.items():
                    buf[1].append(node.text or "")
    return anchors


def _comment_meta(doc) -> dict[int, dict]:
    out: dict[int, dict] = {}
    for c in doc.comments:
        ts = getattr(c, "timestamp", None)
        out[c.comment_id] = {
            "id": c.comment_id,
            "author": c.author or "",
            "initials": c.initials or "",
            "date": ts.isoformat() if isinstance(ts, dt.datetime) else (str(ts) if ts else ""),
            "text": c.text,
        }
    return out


# ---------------------------------------------------------------- read

def read_comments(path: str) -> dict:
    doc = Document(path)
    meta = _comment_meta(doc)
    anchors = _anchor_map(doc)
    items = []
    for cid in sorted(meta):
        m = meta[cid]
        a = anchors.get(cid, {})
        m = dict(m)
        m["anchor_para"] = a.get("para")
        m["anchor_text"] = a.get("text", "")
        items.append(m)
    return {
        "ok": True,
        "op": "read",
        "path": path,
        "count": len(items),
        "comments": items,
    }


# ---------------------------------------------------------------- add

def _split_run_for_find(p, find: str, occurrence: int):
    """Return the list of Run objects exactly covering `find` inside paragraph p.

    Falls back to the whole paragraph's runs when the match spans several runs.
    """
    runs = p.runs
    joined = "".join(r.text for r in runs)
    if not find:
        return runs
    if not joined:
        raise CoreError("paragraph has no text to anchor a comment on")
    start = -1
    pos = -1
    for _ in range(max(1, occurrence)):
        start = joined.find(find, pos + 1)
        if start < 0:
            raise CoreError(f"text not found in paragraph: {find!r}")
        pos = start
    end = start + len(find)

    # map char offsets to runs
    spans = []
    cursor = 0
    for r in runs:
        t = r.text
        spans.append((cursor, cursor + len(t), r))
        cursor += len(t)

    hit = [r for (s, e, r) in spans if e > start and s < end]
    if not hit:
        raise CoreError(f"text not found in paragraph: {find!r}")
    first, last = hit[0], hit[-1]

    # exact single-run match -> split so the anchor is precisely the match
    if len(hit) == 1:
        r = hit[0]
        s, e, _ = next(x for x in spans if x[2] is r)
        local_start = start - s
        local_end = end - s
        if local_start == 0 and local_end == len(r.text):
            return [r]
        rpr = None
        if r._r.rPr is not None:
            rpr = copy.deepcopy(r._r.rPr)
        text = r.text
        r.text = text[:local_start]
        mid = copy.deepcopy(r._r)
        # replace <w:t> content of the copy
        for t in mid.findall(qn("w:t")):
            mid.remove(t)
        from docx.oxml import OxmlElement
        newt = OxmlElement("w:t")
        newt.text = text[local_start:local_end]
        newt.set(qn("xml:space"), "preserve")
        mid.append(newt)
        tail = copy.deepcopy(r._r)
        for t in tail.findall(qn("w:t")):
            tail.remove(t)
        newt2 = OxmlElement("w:t")
        newt2.text = text[local_end:]
        newt2.set(qn("xml:space"), "preserve")
        tail.append(newt2)
        r._r.addnext(mid)
        mid.addnext(tail)
        from docx.text.run import Run
        return [Run(mid, p)]
    return hit


def add_comments(path: str, items: list[dict], author: str = "", copy: bool | None = None) -> dict:
    if not items:
        raise CoreError("no comment items given")
    target, copied = _write_target(path, copy)
    doc = Document(target)
    added = []
    for n, item in enumerate(items, start=1):
        try:
            para_no = int(item.get("para", 0))
        except (TypeError, ValueError):
            raise CoreError(f"item {n}: para must be an integer")
        if para_no < 1:
            raise CoreError(f"item {n}: para must be >= 1")
        p = _paragraph_at(doc, para_no)
        expect = item.get("expect")
        if expect and not _p_text(p._p).startswith(expect):
            raise CoreError(
                f"item {n}: paragraph {para_no} does not start with {expect!r}; it starts with "
                f"{_p_text(p._p)[:40]!r}"
            )
        find = item.get("find") or ""
        occurrence = int(item.get("occurrence", 1) or 1)
        try:
            runs = _split_run_for_find(p, find, occurrence)
        except CoreError as e:
            raise CoreError(f"item {n} (para {para_no}): {e}") from None
        text = item.get("comment") or item.get("text") or ""
        if not text:
            raise CoreError(f"item {n}: comment text is empty")
        c = doc.add_comment(runs, text=text, author=item.get("author", author) or "", initials=item.get("initials"))
        added.append({"id": c.comment_id, "para": para_no, "anchor_text": find or _p_text(p._p), "author": c.author})
    doc.save(target)
    return {"ok": True, "op": "add", "path": target, "source": path, "copied": copied, "added": added}


# ---------------------------------------------------------------- delete

def _purge_extended(doc, live_ids: set[int]) -> None:
    """Drop commentsExtended / commentsIds entries that no longer have a comment."""
    try:
        part = doc.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentsExtended"
        )
    except KeyError:
        return
    root = part.element
    # collect paraIds still used by remaining comments
    keep = set()
    for c in doc.comments:
        if c.comment_id not in live_ids:
            continue
        for el in c._comment_elm.iter():
            for attr in ("paraId",):
                v = el.get(f"{{{W14_NS}}}{attr}")
                if v:
                    keep.add(v)
    for ex in list(root):
        pid = ex.get(f"{{{W15_NS}}}paraId") or ex.get(f"{{{W14_NS}}}paraId")
        if pid and pid not in keep:
            root.remove(ex)


def delete_comments(path: str, ids, copy: bool | None = None) -> dict:
    target, copied = _write_target(path, copy)
    doc = Document(target)
    existing = {c.comment_id for c in doc.comments}
    if ids in (None, "all", ["all"]):
        wanted = set(existing)
    else:
        wanted = {int(i) for i in ids}
    missing = sorted(wanted - existing)
    if missing:
        raise CoreError(f"no such comment id(s): {missing}; document has {sorted(existing)}")
    if not wanted:
        raise CoreError("nothing to delete")

    body = doc.element.body
    for el in list(body.iter()):
        if el.tag in (qn("w:commentRangeStart"), qn("w:commentRangeEnd")):
            try:
                cid = int(el.get(qn("w:id")))
            except (TypeError, ValueError):
                continue
            if cid in wanted:
                el.getparent().remove(el)

    # runs that only carry the commentReference
    for el in list(body.iter()):
        if el.tag == qn("w:commentReference"):
            try:
                cid = int(el.get(qn("w:id")))
            except (TypeError, ValueError):
                continue
            if cid in wanted:
                run = el.getparent()
                if run is not None and run.tag == qn("w:r"):
                    # Drop the whole run only when the reference is all it holds;
                    # otherwise keep the text and remove just the reference.
                    outer = run.getparent()
                    if outer is not None:
                        if len(run) <= 1:
                            outer.remove(run)
                        else:
                            run.remove(el)

    comments_el = doc.comments._comments_elm
    for celm in list(comments_el):
        try:
            cid = int(celm.get(qn("w:id")))
        except (TypeError, ValueError):
            continue
        if cid in wanted:
            comments_el.remove(celm)

    remaining = {c.comment_id for c in doc.comments}
    _purge_extended(doc, remaining)
    doc.save(target)
    return {
        "ok": True,
        "op": "delete",
        "path": target,
        "source": path,
        "copied": copied,
        "deleted": sorted(wanted),
        "remaining": len(remaining),
    }


# ---------------------------------------------------------------- CLI

def _parse_flags(args: list[str]) -> tuple[list[str], bool | None]:
    """Split --copy / --no-copy out of positional arguments."""
    rest: list[str] = []
    copy: bool | None = None
    for a in args:
        if a == "--copy":
            copy = True
        elif a == "--no-copy":
            copy = False
        else:
            rest.append(a)
    return rest, copy


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001 - older runtimes
        pass
    if len(argv) < 3:
        print(json.dumps({"ok": False, "error": "usage: comment_core.py read|add|delete <docx> [...]"}))
        return 2
    op, path = argv[1], argv[2]
    rest, copy = _parse_flags(argv[3:])
    try:
        if op == "read":
            result = read_comments(path)
        elif op == "add":
            raw = rest[0] if rest else "{}"
            if raw == "-":
                raw = sys.stdin.read()
            elif raw.startswith("@"):
                with open(raw[1:], "r", encoding="utf-8") as fh:
                    raw = fh.read()
            payload = json.loads(raw) if raw.strip() else {}
            if copy is None:
                copy = payload.get("copy")
            result = add_comments(path, payload.get("items", []), payload.get("author", ""), copy)
        elif op == "delete":
            raw = rest[0] if rest else "all"
            ids = raw if raw == "all" else [int(x) for x in raw.split(",") if x.strip()]
            result = delete_comments(path, ids, copy)
        else:
            raise CoreError(f"unknown op: {op}")
    except CoreError as e:
        result = {"ok": False, "op": op, "error": str(e)}
    except Exception as e:  # noqa: BLE001 - surface anything else as JSON
        result = {"ok": False, "op": op, "error": f"{type(e).__name__}: {e}"}
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
