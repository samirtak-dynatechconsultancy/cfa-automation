"""Repair the comment layer of an openpyxl-saved workbook so Excel (esp. Excel Online) accepts it.

openpyxl re-serialises cell comments into a shape Excel dislikes:
  * the comment VML uses generic namespace prefixes (``ns0:``/``ns1:``/``ns2:``) instead of the
    ``o:``/``v:``/``x:`` prefixes Excel's VML parser expects,
  * relationship targets are absolute (``/xl/comments/comment1.xml``) instead of relative,
  * modern *threaded* comments (``xl/threadedComments`` + ``xl/persons``) are dropped entirely.

The workbook opens, but any operation that re-validates the comment layer — "Create a copy",
"Download a copy", or co-authoring merging a freshly added note — fails with Excel's generic
"Sorry, something went wrong." That, in turn, is what corrupts the file and loses user notes.

The app never *creates* comments itself: every comment originates from the master / existing year
file. So this module transplants the ORIGINAL Excel-native comment parts back in, mapped by sheet
name — byte-faithful, and preserving threaded comments. Any output sheet that has no source
counterpart (e.g. a freshly cloned period sheet) instead gets its openpyxl VML normalised
(prefixes fixed, targets made relative) so it is still valid Excel.

Everything is wrapped so a failure never makes the file worse than plain openpyxl output.
"""

from __future__ import annotations

import io
import logging
import posixpath
import re
import uuid
import zipfile

log = logging.getLogger("cfa.comment_repair")

_NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_T_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_T_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
_T_THREADED = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
_T_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"

_CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
_CT_THREADED = "application/vnd.ms-excel.threadedcomments+xml"
_CT_PERSON = "application/vnd.ms-excel.person+xml"
_CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"

_COMMENT_REL_TYPES = {_T_COMMENTS, _T_VML, _T_THREADED}


def _attr(tag: str, name: str) -> str | None:
    m = re.search(rf'\b{name}\s*=\s*"([^"]*)"', tag)
    return m.group(1) if m else None


def _rels(xml: str) -> list[dict]:
    out = []
    for m in re.finditer(r"<Relationship\b[^>]*/>", xml):
        t = m.group(0)
        out.append({"Id": _attr(t, "Id"), "Type": _attr(t, "Type"),
                    "Target": _attr(t, "Target"), "Mode": _attr(t, "TargetMode")})
    return out


def _resolve(base_part: str, target: str) -> str:
    """Resolve a relationship Target (relative to the part's folder) to a package path."""
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def _sheet_name_to_part(names: dict[str, bytes]) -> dict[str, str]:
    """Map worksheet display name -> package part path via workbook.xml + its rels."""
    wb = names["xl/workbook.xml"].decode("utf-8", "replace")
    rel_xml = names.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "replace")
    id2tgt = {r["Id"]: r["Target"] for r in _rels(rel_xml)}
    out: dict[str, str] = {}
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        tag = m.group(0)
        nm, rid = _attr(tag, "name"), _attr(tag, "r:id")
        if nm and rid and rid in id2tgt:
            out[nm] = _resolve("xl/workbook.xml", id2tgt[rid])
    return out


def _sheet_comment_parts(part: str, names: dict[str, bytes]) -> dict[str, str]:
    """For a worksheet part, return {kind: package_path} for its comment/vml/threaded parts."""
    rp = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if rp not in names:
        return {}
    found: dict[str, str] = {}
    for r in _rels(names[rp].decode("utf-8", "replace")):
        if r["Type"] == _T_COMMENTS:
            found["comments"] = _resolve(part, r["Target"])
        elif r["Type"] == _T_VML:
            found["vml"] = _resolve(part, r["Target"])
        elif r["Type"] == _T_THREADED:
            found["threaded"] = _resolve(part, r["Target"])
    return found


def _normalise_vml(data: bytes) -> bytes:
    """Rewrite openpyxl's ns0/ns1/ns2 comment-VML prefixes to Excel's o/v/x."""
    txt = data.decode("utf-8", "replace")
    # xmlns declarations -> canonical prefixes
    txt = txt.replace('xmlns:ns0="urn:schemas-microsoft-com:office:office"',
                      'xmlns:o="urn:schemas-microsoft-com:office:office"')
    txt = txt.replace('xmlns:ns1="urn:schemas-microsoft-com:vml"',
                      'xmlns:v="urn:schemas-microsoft-com:vml"')
    txt = txt.replace('xmlns:ns2="urn:schemas-microsoft-com:office:excel"',
                      'xmlns:x="urn:schemas-microsoft-com:office:excel"')
    # element/attribute prefixes
    txt = re.sub(r"([<\s/])ns0:", r"\1o:", txt)
    txt = re.sub(r"([<\s/])ns1:", r"\1v:", txt)
    txt = re.sub(r"([<\s/])ns2:", r"\1x:", txt)
    return txt.encode("utf-8")


def _fresh_id(existing: set[str], prefix: str = "rId") -> str:
    n = 1
    while f"{prefix}{n}" in existing:
        n += 1
    rid = f"{prefix}{n}"
    existing.add(rid)
    return rid


def _unique_part(existing: set[str], path: str) -> str:
    if path not in existing:
        existing.add(path)
        return path
    stem, dot, ext = path.rpartition(".")
    n = 2
    while f"{stem}_{n}{dot}{ext}" in existing:
        n += 1
    p = f"{stem}_{n}{dot}{ext}"
    existing.add(p)
    return p


def repair_comment_layer(out_bytes: bytes, source_bytes: bytes,
                         master_bytes: bytes | None = None,
                         new_comment_sheets: set[str] | None = None,
                         changed_sheets: set[str] | None = None) -> bytes:
    """Return `out_bytes` with its comment layer replaced by Excel-native parts from `source_bytes`.

    The name-matched transplant from `source_bytes` is UNCHANGED — same-period reruns keep their notes
    exactly as before. Additionally, when `master_bytes` (the template) and `new_comment_sheets` (the
    period sheets created THIS run) are given, the template sheet's comment is grafted onto those new
    sheets — the same "carry the note over" idea, only the source is the template excel. This runs as a
    separate, isolated pass so it can never affect the existing behaviour.

    `changed_sheets` are period sheets whose rows/columns were inserted THIS run (by
    sync_template_rows, or by adding an entity column). For those, the source's comment coordinates
    are stale — the cells (and openpyxl's own correctly-positioned comments) have shifted — so we do
    NOT transplant; openpyxl's comments are kept and merely normalised (VML prefixes + relative
    targets) so they stay Excel-valid AND correctly positioned.

    Falls back to the (normalised, else original) openpyxl bytes on any error so the result is never
    worse than what openpyxl produced.
    """
    try:
        result = _repair(out_bytes, source_bytes, changed_sheets or frozenset())
        if master_bytes and new_comment_sheets:
            try:
                result = _graft_template_comments(result, master_bytes, new_comment_sheets)
            except Exception:               # graft is optional — never worsen the repaired output
                log.exception("template-comment graft failed; keeping repaired output")
        return result
    except Exception:                       # never regress below plain openpyxl output
        log.exception("comment-layer repair failed; falling back to normalise-only")
        try:
            return _normalise_only(out_bytes)
        except Exception:
            log.exception("comment-VML normalisation also failed; using raw openpyxl output")
            return out_bytes


def _read_zip(data: bytes) -> tuple[dict[str, bytes], list[str]]:
    names: dict[str, bytes] = {}
    order: list[str] = []
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for n in z.namelist():
            names[n] = z.read(n)
            order.append(n)
    return names, order


def _write_zip(names: dict[str, bytes], order: list[str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        seen = set()
        for n in order:
            if n in names and n not in seen:
                z.writestr(n, names[n]); seen.add(n)
        for n, data in names.items():        # any newly added parts
            if n not in seen:
                z.writestr(n, data); seen.add(n)
    return buf.getvalue()


def _normalise_only(out_bytes: bytes) -> bytes:
    names, order = _read_zip(out_bytes)
    changed = False
    for n in list(names):
        if n.endswith(".vml") and b"ns0:" in names[n]:
            names[n] = _normalise_vml(names[n]); changed = True
        if re.match(r"xl/worksheets/_rels/sheet\d+\.xml\.rels$", n):
            fixed = _make_targets_relative(names[n].decode("utf-8", "replace"), n)
            if fixed is not None:
                names[n] = fixed.encode("utf-8"); changed = True
    return _write_zip(names, order) if changed else out_bytes


def _make_targets_relative(rels_xml: str, rels_part: str) -> str | None:
    """Rewrite absolute comment/vml targets (``/xl/...``) to be relative to the sheet folder."""
    sheet_dir = posixpath.dirname(posixpath.dirname(rels_part))   # drop /_rels/<name>.rels
    changed = False

    def repl(m: re.Match) -> str:
        nonlocal changed
        tag, tgt = m.group(0), m.group(1)
        if tgt.startswith("/"):
            rel = posixpath.relpath(tgt.lstrip("/"), sheet_dir)
            changed = True
            return tag.replace(f'Target="{tgt}"', f'Target="{rel}"')
        return tag

    fixed = re.sub(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*/>', repl, rels_xml)
    return fixed if changed else None


def _repair(out_bytes: bytes, source_bytes: bytes, changed_sheets: set[str] = frozenset()) -> bytes:
    out, order = _read_zip(out_bytes)
    src, _ = _read_zip(source_bytes)

    src_name2part = _sheet_name_to_part(src)
    out_name2part = _sheet_name_to_part(out)
    existing_parts = set(out)

    ct_add: dict[str, str] = {}        # PartName -> ContentType overrides to add
    need_vml_default = False
    transplanted_sheets: set[str] = set()
    persons_added = False
    add_person_rel = False

    for name, out_part in out_name2part.items():
        if name in changed_sheets:
            continue      # rows/cols shifted this run -> keep openpyxl's own comments (normalised)
        src_part = src_name2part.get(name)
        if not src_part:
            continue                                   # cloned sheet -> handled by normalise pass
        parts = _sheet_comment_parts(src_part, src)
        if not parts:
            continue

        out_rels_path = posixpath.join(posixpath.dirname(out_part), "_rels",
                                       posixpath.basename(out_part) + ".rels")
        rels_xml = out.get(out_rels_path, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           b'<Relationships xmlns="%s"></Relationships>' % _NS_REL.encode()
                           ).decode("utf-8", "replace")
        keep = [r for r in _rels(rels_xml) if r["Type"] not in _COMMENT_REL_TYPES]
        # Remove the openpyxl comment/vml parts this sheet pointed at.
        for r in _rels(rels_xml):
            if r["Type"] in _COMMENT_REL_TYPES:
                out.pop(_resolve(out_part, r["Target"]), None)

        used_ids = {r["Id"] for r in keep}
        new_rels = list(keep)
        vml_rel_id = None

        def _graft(kind: str, folder: str, base: str, rel_type: str, ct_type: str) -> str | None:
            if kind not in parts or parts[kind] not in src:
                return None
            dst = _unique_part(existing_parts, f"{folder}/{base}")
            out[dst] = src[parts[kind]]
            rid = _fresh_id(used_ids)
            tgt = posixpath.relpath(dst, posixpath.dirname(out_part))
            new_rels.append({"Id": rid, "Type": rel_type, "Target": tgt, "Mode": None})
            if ct_type:
                ct_add["/" + dst] = ct_type
            return rid

        _graft("comments", "xl", "comments1.xml", _T_COMMENTS, _CT_COMMENTS)
        vml_rel_id = _graft("vml", "xl/drawings", "vmlDrawing1.vml", _T_VML, "")
        # VML uses a Default extension (.vml) content-type, not an Override.
        if vml_rel_id:
            need_vml_default = True
        tc_id = _graft("threaded", "xl/threadedComments", "threadedComment1.xml",
                       _T_THREADED, _CT_THREADED)
        if tc_id:
            add_person_rel = True

        # Rewrite the sheet's <legacyDrawing r:id="..."> to point at the transplanted VML.
        # NB: the r: prefix is declared inline on the element, so keep the xmlns:r here.
        if vml_rel_id and out_part in out:
            ld = (f'<legacyDrawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
                  f'2006/relationships" r:id="{vml_rel_id}"/>')
            sxml = out[out_part].decode("utf-8", "replace")
            if "<legacyDrawing" in sxml:
                sxml = re.sub(r"<legacyDrawing\b[^>]*?/>", ld, sxml, count=1)
            else:
                sxml = sxml.replace("</worksheet>", ld + "</worksheet>")
            out[out_part] = sxml.encode("utf-8")

        out[out_rels_path] = _render_rels(new_rels).encode("utf-8")
        transplanted_sheets.add(name)

    if not transplanted_sheets:
        # Nothing matched by name; just normalise whatever openpyxl produced.
        return _normalise_only(out_bytes)

    # Bring persons.xml over once (shared by all threaded comments) and wire it to the workbook.
    if add_person_rel:
        src_person = next((p for p in src if re.match(r"xl/persons/person.*\.xml$", p)), None)
        if src_person:
            if "xl/persons/person.xml" not in out:
                out["xl/persons/person.xml"] = src[src_person]
            ct_add["/xl/persons/person.xml"] = _CT_PERSON
            persons_added = True

    if persons_added:
        wb_rels_path = "xl/_rels/workbook.xml.rels"
        wb_rels = out.get(wb_rels_path, b"").decode("utf-8", "replace")
        rlist = _rels(wb_rels)
        if not any(r["Type"] == _T_PERSON for r in rlist):
            used = {r["Id"] for r in rlist}
            rlist.append({"Id": _fresh_id(used), "Type": _T_PERSON,
                          "Target": "persons/person.xml", "Mode": None})
            out[wb_rels_path] = _render_rels(rlist).encode("utf-8")

    # (Transplanted sheets already dropped their own openpyxl comment/vml parts above; cloned
    # sheets keep theirs — normalised below — so there is nothing to sweep here.)

    # Normalise any remaining (cloned-sheet) openpyxl VML + relative-ise its rels.
    for n in list(out):
        if n.endswith(".vml") and b"ns0:" in out[n]:
            out[n] = _normalise_vml(out[n])
        if re.match(r"xl/worksheets/_rels/sheet\d+\.xml\.rels$", n):
            fixed = _make_targets_relative(out[n].decode("utf-8", "replace"), n)
            if fixed is not None:
                out[n] = fixed.encode("utf-8")

    out["[Content_Types].xml"] = _edit_content_types(
        out["[Content_Types].xml"].decode("utf-8", "replace"),
        add=ct_add, need_vml_default=need_vml_default, present_parts=set(out)).encode("utf-8")

    # Refresh the order list so newly added parts are included on write.
    order = [n for n in order if n in out] + [n for n in out if n not in order]
    return _write_zip(out, order)


def _render_rels(rels: list[dict]) -> str:
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             f'<Relationships xmlns="{_NS_REL}">']
    for r in rels:
        extra = f' TargetMode="{r["Mode"]}"' if r.get("Mode") else ""
        parts.append(f'<Relationship Id="{r["Id"]}" Type="{r["Type"]}" '
                     f'Target="{r["Target"]}"{extra}/>')
    parts.append("</Relationships>")
    return "".join(parts)


def _edit_content_types(ct: str, *, add: dict[str, str], need_vml_default: bool,
                        present_parts: set[str]) -> str:
    """Minimally edit [Content_Types].xml: add new overrides, ensure the vml Default, and prune
    any Override whose part no longer exists (never rebuild — preserves all unrelated entries)."""
    # Prune dangling overrides (parts we deleted, e.g. openpyxl's xl/comments/comment1.xml).
    def _keep(m: re.Match) -> str:
        pn = _attr(m.group(0), "PartName") or ""
        return "" if pn.lstrip("/") not in present_parts else m.group(0)
    ct = re.sub(r"<Override\b[^>]*/>", _keep, ct)

    inject = []
    if need_vml_default and 'Extension="vml"' not in ct:
        inject.append(f'<Default Extension="vml" ContentType="{_CT_VML}"/>')
    existing_parts = {_attr(m.group(0), "PartName") for m in re.finditer(r"<Override\b[^>]*/>", ct)}
    for pn, cty in add.items():
        if pn not in existing_parts:
            inject.append(f'<Override PartName="{pn}" ContentType="{cty}"/>')
    if inject:
        ct = ct.replace("</Types>", "".join(inject) + "</Types>", 1)
    return ct


# ---------------------------------------------------------------------------
# Template-comment graft (additive; leaves the name-matched repair above untouched)
# ---------------------------------------------------------------------------

def _comment_source_sheet(names: dict[str, bytes]) -> tuple[str | None, dict[str, str]]:
    """Find the first sheet in `names` that carries comment parts; return (part_path, parts)."""
    for _nm, part in _sheet_name_to_part(names).items():
        parts = _sheet_comment_parts(part, names)
        if parts:
            return part, parts
    return None, {}


def _new_guid() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def _graft_template_comments(out_bytes: bytes, master_bytes: bytes,
                             new_sheets: set[str]) -> bytes:
    """Copy the master TEMPLATE sheet's comment onto the given NEW period sheets.

    Purely additive and byte-faithful (preserves threaded comments): only touches sheets named in
    `new_sheets` that don't already carry a comment, so nothing the name-matched repair produced is
    changed. Each target gets a FRESH comment GUID so Excel never sees the same threaded-comment id
    on two sheets. Returns the input bytes unchanged if there is nothing to do.
    """
    mnames, _ = _read_zip(master_bytes)
    _tmpl_part, tmpl_parts = _comment_source_sheet(mnames)
    if not tmpl_parts:
        return out_bytes                       # template has no comment -> nothing to carry

    out, order = _read_zip(out_bytes)
    out_name2part = _sheet_name_to_part(out)
    existing_parts = set(out)

    old_guid = None
    if tmpl_parts.get("threaded") and tmpl_parts["threaded"] in mnames:
        mm = re.search(r'\bid="(\{[^"]+\})"',
                       mnames[tmpl_parts["threaded"]].decode("utf-8", "replace"))
        old_guid = mm.group(1) if mm else None

    ct_add: dict[str, str] = {}
    need_vml_default = False
    added_threaded = False
    grafted = False

    for name in new_sheets:
        out_part = out_name2part.get(name)
        if not out_part:
            continue
        out_rels_path = posixpath.join(posixpath.dirname(out_part), "_rels",
                                       posixpath.basename(out_part) + ".rels")
        rels_xml = out.get(out_rels_path, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           b'<Relationships xmlns="%s"></Relationships>' % _NS_REL.encode()
                           ).decode("utf-8", "replace")
        rels = _rels(rels_xml)
        if any(r["Type"] in _COMMENT_REL_TYPES for r in rels):
            continue                           # already has a comment -> leave it alone

        new_guid = _new_guid()
        used_ids = {r["Id"] for r in rels}
        new_rels = list(rels)

        def _graft(kind, folder, base, rel_type, ct_type, rewrite_guid=False):
            nonlocal need_vml_default
            if kind not in tmpl_parts or tmpl_parts[kind] not in mnames:
                return None
            data = mnames[tmpl_parts[kind]]
            if rewrite_guid and old_guid:
                data = data.replace(old_guid.encode(), new_guid.encode())
            dst = _unique_part(existing_parts, f"{folder}/{base}")
            out[dst] = data
            rid = _fresh_id(used_ids)
            tgt = posixpath.relpath(dst, posixpath.dirname(out_part))
            new_rels.append({"Id": rid, "Type": rel_type, "Target": tgt, "Mode": None})
            if ct_type:
                ct_add["/" + dst] = ct_type
            return rid

        _graft("comments", "xl", "comments1.xml", _T_COMMENTS, _CT_COMMENTS, rewrite_guid=True)
        v_id = _graft("vml", "xl/drawings", "vmlDrawing1.vml", _T_VML, "")
        if v_id:
            need_vml_default = True
        t_id = _graft("threaded", "xl/threadedComments", "threadedComment1.xml",
                      _T_THREADED, _CT_THREADED, rewrite_guid=True)
        if t_id:
            added_threaded = True

        # Wire the sheet's <legacyDrawing> to the grafted VML.
        if v_id and out_part in out:
            ld = (f'<legacyDrawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
                  f'2006/relationships" r:id="{v_id}"/>')
            sxml = out[out_part].decode("utf-8", "replace")
            if "<legacyDrawing" in sxml:
                sxml = re.sub(r"<legacyDrawing\b[^>]*?/>", ld, sxml, count=1)
            else:
                sxml = sxml.replace("</worksheet>", ld + "</worksheet>")
            out[out_part] = sxml.encode("utf-8")

        out[out_rels_path] = _render_rels(new_rels).encode("utf-8")
        grafted = True

    if not grafted:
        return out_bytes

    # persons.xml (shared) from the master, wired to the workbook — needed for threaded comments.
    if added_threaded and "xl/persons/person.xml" not in out:
        mp = next((p for p in mnames if re.match(r"xl/persons/person.*\.xml$", p)), None)
        if mp:
            out["xl/persons/person.xml"] = mnames[mp]
            ct_add["/xl/persons/person.xml"] = _CT_PERSON
    if added_threaded and "xl/persons/person.xml" in out:
        wb_rels_path = "xl/_rels/workbook.xml.rels"
        wb_rels = out.get(wb_rels_path, b"").decode("utf-8", "replace")
        rlist = _rels(wb_rels)
        if not any(r["Type"] == _T_PERSON for r in rlist):
            used = {r["Id"] for r in rlist}
            rlist.append({"Id": _fresh_id(used), "Type": _T_PERSON,
                          "Target": "persons/person.xml", "Mode": None})
            out[wb_rels_path] = _render_rels(rlist).encode("utf-8")

    out["[Content_Types].xml"] = _edit_content_types(
        out["[Content_Types].xml"].decode("utf-8", "replace"),
        add=ct_add, need_vml_default=need_vml_default, present_parts=set(out)).encode("utf-8")

    order = [n for n in order if n in out] + [n for n in out if n not in order]
    return _write_zip(out, order)
