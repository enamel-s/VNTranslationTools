#!/usr/bin/env python3
"""
vntextpatch.py - a Python take on VNTextPatch for Ren'Py games.

It does what VNTextPatch's `extractlocal` / `insertlocal` do for .rpy scripts,
but can pull the scripts straight out of the game's .rpa archives:

  extract    archives/folder -> script.xlsx   (same layout as template.xlsx)
  translate  script.xlsx     -> fills the "TL" column using the Gemini API
                                (several API keys in rotation + context caching)
  insert     archives/folder + script.xlsx -> patched .rpy files / new .rpa

Spreadsheet columns (identical to VNTextPatch):
  A original character | B original line | C translated character |
  D TL | E TLC | F Edit | G Notes
When inserting, the text used for a row is Edit, else TLC, else TL, else the
original line ("." in Edit/TLC means "skip this column", "(empty)" means "").

Requires: Python 3.8+, openpyxl.  The Gemini part only uses the standard library.
"""

import argparse
import copy
import hashlib
import io
import json
import os
import pickle
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import zlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from xml.etree import ElementTree
from xml.sax.saxutils import escape as xml_escape

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("This script needs openpyxl:  pip install openpyxl")


# ----------------------------------------------------------------------------
# RPA archives
# ----------------------------------------------------------------------------

class _SafeUnpickler(pickle.Unpickler):
    """An RPA index only holds dicts/lists/tuples/ints/strings. Refuse anything else."""

    def find_class(self, module, name):
        raise pickle.UnpicklingError(f"unexpected object {module}.{name} in RPA index")


class RpaArchive:
    """Reader for Ren'Py RPA-2.0 / RPA-3.0 archives."""

    def __init__(self, path):
        self.path = path
        self.index = {}  # member name -> list of (offset, length, prefix bytes)
        with open(path, "rb") as f:
            header = f.readline().decode("latin-1").split()
            if not header or header[0] not in ("RPA-3.0", "RPA-2.0"):
                raise ValueError(f"{path}: not an RPA-2.0/3.0 archive")
            index_offset = int(header[1], 16)
            key = 0
            if header[0] == "RPA-3.0":
                for part in header[2:]:
                    key ^= int(part, 16)
            f.seek(index_offset)
            raw = _SafeUnpickler(io.BytesIO(zlib.decompress(f.read())), encoding="latin-1").load()

        for name, chunks in raw.items():
            if isinstance(name, bytes):
                name = name.decode("utf-8", "replace")
            parsed = []
            for chunk in chunks:
                offset, length = chunk[0] ^ key, chunk[1] ^ key
                prefix = chunk[2] if len(chunk) > 2 else b""
                if isinstance(prefix, str):
                    prefix = prefix.encode("latin-1")
                parsed.append((offset, length, prefix or b""))
            self.index[name.replace("\\", "/")] = parsed

    def names(self):
        return sorted(self.index)

    def read(self, name):
        out = []
        with open(self.path, "rb") as f:
            for offset, length, prefix in self.index[name]:
                f.seek(offset)
                out.append(prefix + f.read(length - len(prefix)))
        return b"".join(out)


def write_rpa(path, files, key=0x42424242):
    """Write an RPA-3.0 archive. `files` is an iterable of (name, bytes)."""
    index = {}
    with open(path, "wb") as f:
        f.write(b"RPA-3.0 " + b"0" * 16 + b" " + b"0" * 8 + b"\n")
        for name, data in files:
            f.write(b"Made with Ren'Py.")
            # 2-tuples are understood by every Ren'Py version.
            index[name] = [(f.tell() ^ key, len(data) ^ key)]
            f.write(data)
        index_offset = f.tell()
        f.write(zlib.compress(pickle.dumps(index, 2)))
        f.seek(0)
        f.write(("RPA-3.0 %016x %08x\n" % (index_offset, key)).encode("ascii"))


# ----------------------------------------------------------------------------
# Collecting scripts from folders and archives
# ----------------------------------------------------------------------------

@dataclass
class ScriptSource:
    name: str            # path relative to the game folder / inside the archive
    data: bytes
    archive: object = None   # RpaArchive the script came from (None = loose file)


def _decompile_rpyc(name, data, unrpyc):
    """Decompile one .rpyc with an external unrpyc. Returns .rpy bytes or None."""
    if unrpyc:
        cmd = [sys.executable, unrpyc] if unrpyc.endswith(".py") else [unrpyc]
    elif shutil.which("unrpyc"):
        cmd = ["unrpyc"]
    else:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        rpyc = os.path.join(tmp, os.path.basename(name))
        with open(rpyc, "wb") as f:
            f.write(data)
        subprocess.run(cmd + [rpyc], cwd=tmp, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        rpy = rpyc[:-1]
        if os.path.exists(rpy):
            with open(rpy, "rb") as f:
                return f.read()
    return None


def collect_sources(input_path, unrpyc=None):
    """
    Returns (scripts, archives).
    `input_path` may be a single .rpa file, or a folder that contains loose .rpy
    files and/or .rpa archives (e.g. the game's "game" folder).
    Loose .rpy files win over archived ones with the same name.
    """
    scripts = {}
    archives = []
    loose = []

    if os.path.isfile(input_path):
        if input_path.lower().endswith(".rpa"):
            archives.append(RpaArchive(input_path))
        elif input_path.lower().endswith(".rpy"):
            loose.append((os.path.basename(input_path), input_path))
        else:
            raise ValueError(f"{input_path}: expected a folder, an .rpa archive or an .rpy file")
    elif os.path.isdir(input_path):
        for root, dirs, files in os.walk(input_path):
            dirs.sort()
            for fn in sorted(files):
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, input_path).replace(os.sep, "/")
                if fn.lower().endswith(".rpa"):
                    archives.append(RpaArchive(full))
                elif fn.lower().endswith(".rpy"):
                    loose.append((rel, full))
    else:
        raise FileNotFoundError(f"{input_path} does not exist")

    missing_decompiler = []
    for arc in archives:
        names = set(arc.names())
        for name in arc.names():
            low = name.lower()
            if low.endswith(".rpy"):
                scripts[name] = ScriptSource(name, arc.read(name), arc)
            elif low.endswith(".rpyc") and name[:-1] not in names:
                data = _decompile_rpyc(name, arc.read(name), unrpyc)
                if data is None:
                    missing_decompiler.append(f"{os.path.basename(arc.path)}:{name}")
                else:
                    scripts[name[:-1]] = ScriptSource(name[:-1], data, arc)
    for rel, full in loose:
        with open(full, "rb") as f:
            scripts[rel] = ScriptSource(rel, f.read(), None)

    if missing_decompiler:
        print(f"WARNING: {len(missing_decompiler)} compiled script(s) only exist as .rpyc and were skipped "
              f"(e.g. {missing_decompiler[0]}).\n"
              "         Install unrpyc (https://github.com/CensoredUsername/unrpyc) and pass --unrpyc path/to/unrpyc.py",
              file=sys.stderr)
    return [scripts[n] for n in sorted(scripts)], archives


# ----------------------------------------------------------------------------
# Ren'Py script parser (port of RenpyScript.cs / PlainTextScript.cs)
# ----------------------------------------------------------------------------

@dataclass
class StringRange:
    offset: int          # offset of the string *content* (inside the quotes)
    length: int
    quote: str           # ' or " or ''' or \"\"\"
    kind: str = "all"    # say / choice / name / string / all
    who: str = None      # speaker identifier (say statements)

    @property
    def end(self):
        return self.offset + self.length


def scan_strings(script):
    """
    Same scanner as RenpyScript.GetRanges(): skips # comments, honours backslash
    escapes, yields the content of every quoted string. Unlike the C# version it
    also understands triple-quoted strings.
    """
    ranges = []
    pos, length = 0, len(script)
    while pos < length:
        c = script[pos]
        if c == "#":
            while pos < length and script[pos] not in "\r\n":
                pos += 1
        elif c in "'\"":
            quote = c * 3 if script.startswith(c * 3, pos) else c
            start = pos + len(quote)
            pos = start
            while pos < length and not script.startswith(quote, pos):
                pos += 2 if script[pos] == "\\" else 1
            pos = min(pos, length)
            ranges.append(StringRange(start, pos - start, quote))
            pos += len(quote)
        else:
            pos += 1
    return ranges


RENPY_KEYWORDS = {
    "at", "call", "camera", "default", "define", "elif", "else", "for", "hide", "if", "image", "init",
    "jump", "label", "layeredimage", "menu", "new", "nvl", "old", "onlayer", "pass", "pause", "play",
    "python", "queue", "return", "scene", "screen", "show", "stop", "style", "transform", "translate",
    "voice", "while", "window", "with", "zorder", "behind", "as", "expression", "rpy", "testcase",
    "import", "from", "print", "renpy", "action", "add", "use", "key", "text", "textbutton", "label",
    "imagebutton", "bar", "vbar", "input", "timer", "on", "has", "frame", "vbox", "hbox", "null",
    "linear", "ease", "easein", "easeout", "xpos", "ypos", "alpha", "zoom", "repeat", "parallel", "block",
    "contains", "choice", "time", "function", "event", "animation", "music", "sound", "audio",
}
NO_NAME_SPEAKERS = {"extend", "narrator", "centered", "vcentered", "nvl_narrator", "name_only"}

_RE_PY_BLOCK = re.compile(r"^(init\s+(-?\d+\s+)?)?python\b[^#\"']*:\s*(#.*)?$")
_RE_SCREEN_BLOCK = re.compile(r"^(init\s+(-?\d+\s+)?)?screen\s+\w+.*:\s*(#.*)?$")
_RE_INERT_BLOCK = re.compile(
    r"^(?:init\s+(?:-?\d+\s+)?)?(image|transform|style|show|scene|hide|layeredimage|translate|camera|window|with|on|parallel|block|"
    r"choice|contains|testcase)\b.*:\s*(#.*)?$")
_RE_FILE_NAME = re.compile(r"^[^\n]*\.(png|jpe?g|webp|gif|ogg|mp3|wav|opus|ttf|otf|rpyc?|txt|chr|webm|mp4|ogv|json)$", re.I)
_RE_SAY_PREFIX = re.compile(r"^\s*([A-Za-z_]\w*)((?:\s+[^\s\"'=(:,]+)*)\s+$")
_RE_LITERAL_NAME_PREFIX = re.compile(r"^\s*([\"'])((?:\\.|(?!\1).)*)\1\s+$")
_RE_SAY_SUFFIX = re.compile(r"^\s*(nointeract\b.*|with\b.*|\(.*\)\s*)?\s*(#.*)?$")
_RE_CHOICE_SUFFIX = re.compile(r"^\s*(\(.*\)\s*)?(if\b.*)?:\s*(#.*)?$")
_RE_TRANSLATE_FUNC = re.compile(r"(?<![\w.])(_|__|_p)\(\s*$")
_RE_CHARACTER_CALL = re.compile(r"(?<![A-Za-z_])Character\(\s*$")
_RE_SCREEN_TEXT = re.compile(r"^\s*(text|textbutton|label|tooltip)\s+$")
_RE_DYNAMIC_CHAR = re.compile(
    r"^\s*define\s+([\w.]+)\s*=\s*DynamicCharacter\(\s*[\"']([\w.]+)[\"']", re.M)
_RE_STATIC_CHAR = re.compile(
    r"^\s*define\s+([\w.]+)\s*=\s*Character\(\s*([\"'])((?:\\.|(?!\2).)*)\2", re.M)
_RE_MARKUP = re.compile(r"\{[^{}]*\}|\[[^\[\]]*\]|\\.|%\(\w+\)s")


def has_translatable_text(text):
    if _RE_FILE_NAME.match(text.strip()):
        return False
    return any(ch.isalpha() for ch in _RE_MARKUP.sub("", text))


class RenpyProject:
    """Knowledge shared by all scripts of one game: who the characters are."""

    def __init__(self, texts, extra_patterns=()):
        self.speakers = {}        # speaker identifier -> display name
        self.name_vars = set()    # variables holding a DynamicCharacter's name
        self.extra = [re.compile(p) for p in extra_patterns]
        dynamic = {}
        for text in texts:
            for ident, var in _RE_DYNAMIC_CHAR.findall(text):
                dynamic[ident] = var
                self.name_vars.add(var)
            for ident, _q, name in _RE_STATIC_CHAR.findall(text):
                self.speakers[ident] = name
        for ident, var in dynamic.items():
            pattern = re.compile(
                r"^\s*(?:default|define|\$)\s+" + re.escape(var) + r"\s*=\s*([\"'])((?:\\.|(?!\1).)*)\1", re.M)
            for text in texts:
                m = pattern.search(text)
                if m:
                    self.speakers[ident] = m.group(2)
                    break
        self._re_name_assign = None
        if self.name_vars:
            alternatives = "|".join(re.escape(v) for v in sorted(self.name_vars))
            self._re_name_assign = re.compile(r"^\s*(?:default|define|\$)?\s*(?:" + alternatives + r")\s*=\s*$")

    def speaker_name(self, who):
        if not who or who in NO_NAME_SPEAKERS:
            return None
        return self.speakers.get(who, who)


def _line_contexts(script, ranges):
    """Maps each line start offset to 'python', 'screen', 'inert' (ATL/style blocks) or 'renpy'."""
    inside_string = set()
    for r in ranges:
        if "\n" in script[r.offset:r.end]:
            pos = script.find("\n", r.offset, r.end)
            while pos != -1:
                inside_string.add(pos + 1)
                pos = script.find("\n", pos + 1, r.end)

    contexts = {}
    stack = []  # (indent, kind)
    pos = 0
    for line in script.splitlines(True):
        start = pos
        pos += len(line)
        if start in inside_string:
            continue
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        indent = len(line) - len(line.lstrip(" \t"))
        while stack and stack[-1][0] >= indent:
            stack.pop()
        context = stack[-1][1] if stack else "renpy"
        if context != "python" and re.match(r"(\$|define\b|default\b)", stripped):
            contexts[start] = "python"
        else:
            contexts[start] = context
        if context in ("renpy", "screen"):
            if _RE_PY_BLOCK.match(stripped):
                stack.append((indent, "python"))
            elif _RE_SCREEN_BLOCK.match(stripped):
                stack.append((indent, "screen"))
            elif context == "renpy" and _RE_INERT_BLOCK.match(stripped):
                stack.append((indent, "inert"))
    return contexts


def get_ranges(script, project, mode="dialogue"):
    """
    Returns the translatable string ranges of a script, in file order.

    mode "all":      every string literal, exactly like VNTextPatch's RenpyScript
                     (this includes file names, image names, ... be careful).
    mode "dialogue": only text shown to the player - say statements, menu
                     choices, character names, _("...") strings, text in screens,
                     plus anything matched by the project's extra patterns.
    """
    ranges = scan_strings(script)
    if mode == "all":
        return ranges

    contexts = _line_contexts(script, ranges)
    result = []
    for r in ranges:
        text = script[r.offset:r.end]
        if not has_translatable_text(text):
            continue
        quote_start = r.offset - len(r.quote)
        line_start = script.rfind("\n", 0, quote_start) + 1
        line_end = script.find("\n", r.end)
        line_end = len(script) if line_end == -1 else line_end
        prefix = script[line_start:quote_start]
        suffix = script[r.end + len(r.quote):line_end].rstrip("\r")
        context = contexts.get(line_start, "renpy")

        kind, who = None, None
        if _RE_TRANSLATE_FUNC.search(prefix) or any(p.search(prefix) for p in project.extra):
            kind = "string"
        elif _RE_CHARACTER_CALL.search(prefix):
            kind = "name"
        elif project._re_name_assign and project._re_name_assign.match(prefix):
            kind = "name"
        elif context == "screen":
            if _RE_SCREEN_TEXT.match(prefix):
                kind = "string"
        elif context == "renpy":
            if not prefix.strip():
                if _RE_CHOICE_SUFFIX.match(suffix):
                    kind = "choice"
                elif _RE_SAY_SUFFIX.match(suffix):
                    kind = "say"
            else:
                m = _RE_SAY_PREFIX.match(prefix)
                if m and m.group(1) not in RENPY_KEYWORDS and _RE_SAY_SUFFIX.match(suffix):
                    kind, who = "say", m.group(1)
                else:
                    m = _RE_LITERAL_NAME_PREFIX.match(prefix)
                    if m and _RE_SAY_SUFFIX.match(suffix):
                        kind, who = "say", "\0" + m.group(2)
        if kind:
            r.kind, r.who = kind, who
            result.append(r)
    return result


def range_speaker(r, project):
    if r.who and r.who.startswith("\0"):
        return r.who[1:]
    return project.speaker_name(r.who)


def text_for_write(r, text):
    """Makes translated text safe to put between the original quotes."""
    if len(r.quote) == 3:
        return text.replace(r.quote, "\\" + r.quote[0] * 3)
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
    out, i = [], 0
    while i < len(text):
        c = text[i]
        if c == "\\" and i + 1 < len(text):
            out.append(text[i:i + 2])
            i += 2
            continue
        if c == r.quote or (c == "\\" and i + 1 == len(text)):
            out.append("\\")
        out.append(c)
        i += 1
    return "".join(out)


def patch_script(script, ranges, texts):
    """PlainTextScript.WritePatched(): copy the script, replacing each range."""
    out, copy_start = [], 0
    texts = list(texts)
    for idx, r in enumerate(ranges):
        if idx >= len(texts):
            raise ValueError(f"Translation file is missing line {idx + 1} ({script[r.offset:r.end]})")
        out.append(script[copy_start:r.offset])
        out.append(text_for_write(r, texts[idx]))
        copy_start = r.end
    if len(texts) > len(ranges):
        raise ValueError("Translation file has too many lines")
    out.append(script[copy_start:])
    return "".join(out)


def decode_script(data):
    bom = data.startswith(b"\xef\xbb\xbf")
    return data[3 if bom else 0:].decode("utf-8", "replace"), bom


def encode_script(text, bom):
    return (b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8")


# ----------------------------------------------------------------------------
# names.xml (same format as VNTextPatch's CharacterNames)
# ----------------------------------------------------------------------------

class CharacterNames:
    def __init__(self, path):
        self.path = path
        self.translations = {}
        if os.path.exists(path):
            root = ElementTree.parse(path).getroot()
            for n in root.findall("n"):
                original = n.findtext("o") or ""
                if original:
                    self.translations[original] = n.findtext("tl") or original

    def get(self, name):
        return self.translations.setdefault(name, name)

    def set(self, name, translation):
        self.translations[name] = translation

    def save(self):
        lines = ['<?xml version="1.0"?>', "<names>"]
        for original, translation in self.translations.items():
            lines.append(f"  <n>\n    <o>{xml_escape(original)}</o>\n    <tl>{xml_escape(translation)}</tl>\n  </n>")
        lines.append("</names>")
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")


# ----------------------------------------------------------------------------
# Excel workbook (port of ExcelScript.cs / ExcelScriptCollection.cs)
# ----------------------------------------------------------------------------

COL_ORIG_CHAR, COL_ORIG_LINE, COL_TL_CHAR, COL_TL, COL_TLC, COL_EDIT, COL_NOTES = range(1, 8)
EMPTY_TEXT_MARKER = "(empty)"
KINDS = ("say", "choice", "name", "string", "all")


def find_template(explicit=None):
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [explicit] if explicit else [
        os.path.join(here, "template.xlsx"),
        os.path.join(here, "..", "VNTextPatch.Shared", "template.xlsx"),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    if explicit:
        raise FileNotFoundError(explicit)
    return None


def _new_workbook(template):
    if template:
        return openpyxl.load_workbook(template)
    # No template.xlsx around: build an equivalent sheet by hand.
    from openpyxl.styles import Alignment, Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Script"
    for col, title in ((COL_ORIG_LINE, "Original"), (COL_TL, "TL"), (COL_TLC, "TLC"),
                       (COL_EDIT, "Edit"), (COL_NOTES, "Notes")):
        ws.cell(1, col, title).font = Font(bold=True)
    for letter, width in zip("ABCDEFG", (10, 54.25, 14.25, 54.25, 54.25, 54.25, 26.25)):
        ws.column_dimensions[letter].width = width
        ws.column_dimensions[letter].alignment = Alignment(vertical="top", wrap_text=letter not in "AG")
    return wb


def sheet_name_for(script_name, used=None):
    """Sheet name = script name without extension (like VNTextPatch), made Excel-safe."""
    base = re.sub(r"[\\/?*\[\]:]", "_", os.path.splitext(script_name)[0]).strip("'") or "_"
    name = base
    if len(name) > 31 or (used is not None and name.lower() in used):
        name = base[:24] + "~" + hashlib.sha1(script_name.encode("utf-8")).hexdigest()[:6]
    if used is not None:
        used.add(name.lower())
    return name


def _column_styles(ws):
    styles = {}
    for dim in ws.column_dimensions.values():
        lo = dim.min or openpyxl.utils.column_index_from_string(dim.index)
        hi = dim.max or lo
        for col in range(lo, min(hi, 16) + 1):
            styles[col] = dim._style
    return styles


def _cell_text(ws, row, col):
    value = ws.cell(row, col).value
    if value is None:
        return None
    value = str(value)
    return value if value != "" else None


def _set_cell(ws, row, col, value, styles):
    cell = ws.cell(row, col)
    cell.value = value
    cell.data_type = "s" if value is not None else "n"   # never let Excel see a formula
    if col in styles:
        cell._style = copy.copy(styles[col])


def save_workbook(wb, path):
    tmp = path + ".temp"
    wb.save(tmp)
    os.replace(tmp, path)


def join_names(names):
    def quote(n):
        if "/" not in n and '"' not in n:
            return n
        return '"' + n.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return "/".join(quote(n) for n in names)


def row_text(ws, row, stats=None):
    """ExcelScript.GetText(): Edit > TLC > TL > original."""
    original = _cell_text(ws, row, COL_ORIG_LINE)
    translated = _cell_text(ws, row, COL_TL)
    checked = _cell_text(ws, row, COL_TLC)
    edited = _cell_text(ws, row, COL_EDIT)
    if stats is not None:
        stats["total"] += original is not None
        stats["translated"] += translated is not None
        stats["checked"] += checked is not None
        stats["edited"] += edited is not None
    text = (edited if edited != "." else None) or (checked if checked != "." else None) or translated or original
    if text is None:
        return None
    return "" if text == EMPTY_TEXT_MARKER else text


# ----------------------------------------------------------------------------
# extract
# ----------------------------------------------------------------------------

def default_names_path(xlsx_path):
    return os.path.join(os.path.dirname(os.path.abspath(xlsx_path)), "names.xml")


def cmd_extract(args):
    scripts, _archives = collect_sources(args.input, args.unrpyc)
    if not scripts:
        sys.exit("No .rpy scripts found.")
    decoded = [(s, *decode_script(s.data)) for s in scripts]
    project = RenpyProject([text for _s, text, _bom in decoded], args.extra)
    names = CharacterNames(args.names or default_names_path(args.xlsx))

    if args.dump_scripts:
        for s in scripts:
            dest = os.path.join(args.dump_scripts, s.name)
            os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
            with open(dest, "wb") as f:
                f.write(s.data)

    # Remember translations from an existing workbook so re-extracting keeps them.
    previous = {}
    if os.path.exists(args.xlsx):
        old = openpyxl.load_workbook(args.xlsx)
        for ws in old.worksheets:
            previous[ws.title] = {
                row: [_cell_text(ws, row, c) for c in range(1, 8)] for row in range(2, ws.max_row + 1)}
        old.close()

    wb = _new_workbook(find_template(args.template))
    template_ws = wb.worksheets[0]
    styles = _column_styles(template_ws)
    used, first = set(), True
    total_lines = total_chars = carried = 0

    for source, text, _bom in decoded:
        ranges = get_ranges(text, project, args.mode)
        if not ranges:
            continue
        print(source.name)
        title = sheet_name_for(source.name, used)
        if first:
            ws, first = template_ws, False
            ws.title = title
        else:
            ws = wb.copy_worksheet(template_ws)
            ws.title = title
            ws.freeze_panes = template_ws.freeze_panes
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
        old_rows = previous.get(title, {})
        for i, r in enumerate(ranges):
            row = i + 2
            line = text[r.offset:r.end]
            speaker = range_speaker(r, project)
            cell_line = line if line else EMPTY_TEXT_MARKER
            if speaker:
                _set_cell(ws, row, COL_ORIG_CHAR, join_names([speaker]), styles)
                _set_cell(ws, row, COL_TL_CHAR, join_names([names.get(speaker)]), styles)
            _set_cell(ws, row, COL_ORIG_LINE, cell_line, styles)
            old_row = old_rows.get(row)
            if old_row and old_row[COL_ORIG_LINE - 1] == cell_line:
                carried += any(old_row[COL_TL - 1:COL_EDIT])
                for col in (COL_TL, COL_TLC, COL_EDIT, COL_NOTES):
                    if old_row[col - 1] is not None:
                        _set_cell(ws, row, col, old_row[col - 1], styles)
            if ws.cell(row, COL_NOTES).value is None and r.kind != "all":
                _set_cell(ws, row, COL_NOTES, r.kind, styles)
            total_lines += 1
            total_chars += len(line)

    if total_lines == 0:
        sys.exit("No translatable text found.")
    save_workbook(wb, args.xlsx)
    names.save()
    print(f"\nExtracted {total_lines} lines ({total_chars} characters) from {len(wb.worksheets)} scripts "
          f"into {args.xlsx}")
    if carried:
        print(f"Kept {carried} existing translations from the previous workbook.")
    return 0


# ----------------------------------------------------------------------------
# insert
# ----------------------------------------------------------------------------

def cmd_insert(args):
    scripts, archives = collect_sources(args.input, args.unrpyc)
    if not scripts:
        sys.exit("No .rpy scripts found.")
    decoded = [(s, *decode_script(s.data)) for s in scripts]
    project = RenpyProject([text for _s, text, _bom in decoded], args.extra)
    wb = openpyxl.load_workbook(args.xlsx, read_only=False)
    sheets = {ws.title: ws for ws in wb.worksheets}
    stats = {"total": 0, "translated": 0, "checked": 0, "edited": 0}
    used = set()
    patched = {}   # script name -> bytes
    errors = 0

    for source, text, bom in decoded:
        ranges = get_ranges(text, project, args.mode)
        if not ranges:
            patched[source.name] = source.data
            continue
        ws = sheets.get(sheet_name_for(source.name, used))
        if ws is None:
            patched[source.name] = source.data
            stats["total"] += len(ranges)
            continue
        print(source.name)
        texts = []
        for row in range(2, ws.max_row + 1):
            t = row_text(ws, row, stats)
            if t is not None:
                texts.append(t.replace("\r\n", "\n"))
        try:
            patched[source.name] = encode_script(patch_script(text, ranges, texts), bom)
        except ValueError as e:
            errors += 1
            print(f"  ERROR: {e} - is the workbook from another version or --mode? File left unpatched.",
                  file=sys.stderr)
            patched[source.name] = source.data

    os.makedirs(args.output, exist_ok=True)
    repacked = set()
    if args.repack:
        for arc in archives:
            members = []
            arc_names = set(arc.names())
            for name in arc.names():
                if name.lower().endswith(".rpyc") and name[:-1] in patched:
                    continue  # stale compiled script: Ren'Py recompiles the patched .rpy
                if name in patched and any(s.name == name and s.archive is arc for s in scripts):
                    continue
                members.append((name, arc.read(name)))
            for s in scripts:
                if s.archive is arc:
                    members.append((s.name, patched[s.name]))
                    repacked.add(s.name)
            dest = os.path.join(args.output, os.path.basename(arc.path))
            write_rpa(dest, members)
            print(f"Wrote {dest} ({len(members)} files, was {len(arc_names)})")
    for name, data in patched.items():
        if name in repacked:
            continue
        dest = os.path.join(args.output, name)
        os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
        with open(dest, "wb") as f:
            f.write(data)

    total = max(stats["total"], 1)
    print(f"\nTranslated: {stats['translated']}/{stats['total']} ({100 * stats['translated'] // total}%)")
    print(f"Checked:    {stats['checked']}/{stats['total']} ({100 * stats['checked'] // total}%)")
    print(f"Edited:     {stats['edited']}/{stats['total']} ({100 * stats['edited'] // total}%)")
    print(f"Patched scripts written to {args.output}")
    return 1 if errors else 0


# ----------------------------------------------------------------------------
# Gemini: key rotation, context caching, translation
# ----------------------------------------------------------------------------

class GeminiError(Exception):
    def __init__(self, status, message, payload=None):
        super().__init__(f"HTTP {status}: {message}")
        self.status = status
        self.message = message
        self.payload = payload or {}


class ContentBlocked(Exception):
    pass


def _http_json(method, url, api_key, body=None, timeout=180):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Content-Type": "application/json", "x-goog-api-key": api_key})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            payload = json.loads(raw).get("error", {})
        except ValueError:
            payload = {}
        raise GeminiError(e.code, payload.get("message") or raw[:300] or e.reason, payload)
    except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
        raise GeminiError(0, f"network error: {e}")


def load_keys(args):
    keys = []
    if getattr(args, "keys", None):
        keys += re.split(r"[,\s]+", args.keys)
    if getattr(args, "keys_file", None):
        with open(args.keys_file, encoding="utf-8") as f:
            keys += [line.split("#")[0].strip() for line in f]
    if not keys:
        keys += re.split(r"[,\s]+", os.environ.get("GEMINI_API_KEYS", ""))
        keys += [os.environ.get("GEMINI_API_KEY", "")]
    seen, out = set(), []
    for k in keys:
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out


class KeyPool:
    """Hands out API keys round-robin; keys that hit a rate limit rest for a while."""

    def __init__(self, keys):
        if not keys:
            raise ValueError("no Gemini API keys given")
        self.keys = list(keys)
        self.ready_at = {k: 0.0 for k in keys}
        self.dead = {}
        self.lock = threading.Lock()
        self.cursor = 0

    def label(self, key):
        return f"key#{self.keys.index(key) + 1}(...{key[-4:]})"

    def acquire(self):
        while True:
            with self.lock:
                alive = [k for k in self.keys if k not in self.dead]
                if not alive:
                    raise RuntimeError("every API key was rejected: " + "; ".join(
                        f"{self.label(k)}: {why}" for k, why in self.dead.items()))
                now = time.time()
                for step in range(len(self.keys)):
                    key = self.keys[(self.cursor + step) % len(self.keys)]
                    if key not in self.dead and self.ready_at[key] <= now:
                        self.cursor = (self.cursor + step + 1) % len(self.keys)
                        return key
                wait = min(self.ready_at[k] for k in alive) - now
            print(f"  all keys are rate limited, waiting {wait:.0f}s ...", flush=True)
            time.sleep(min(max(wait, 0.5), 300))

    def cooldown(self, key, seconds):
        with self.lock:
            self.ready_at[key] = max(self.ready_at[key], time.time() + seconds)

    def kill(self, key, why):
        with self.lock:
            if key not in self.dead:
                self.dead[key] = why
                print(f"  {self.label(key)} disabled: {why}", flush=True)


def _retry_delay(err):
    """Seconds to rest a key after a 429, from the error details when available."""
    delay, per_day = None, False
    for detail in err.payload.get("details", []) or []:
        if "retryDelay" in detail:
            m = re.match(r"([\d.]+)s", str(detail["retryDelay"]))
            if m:
                delay = float(m.group(1))
        for violation in detail.get("violations", []) or []:
            if "PerDay" in str(violation.get("quotaId", "")):
                per_day = True
    if per_day:
        return 3600.0
    return min(max(delay or 30.0, 2.0), 600.0)


class GeminiClient:
    """
    generateContent with:
      * rotation over many API keys (KeyPool)
      * explicit context caching: the fixed part of every prompt (rules, glossary,
        character sheet, story notes) is uploaded once per key as a cachedContent
        and referenced afterwards. If a key/model can't cache (content too small,
        free tier, ...) the same prefix is sent inline, which still benefits from
        Gemini's implicit caching because the prefix is byte-identical each time.
    """

    def __init__(self, pool, model, api_base, system_text, reference_text, state_path,
                 use_cache=True, ttl=3600, temperature=0.3, thinking_budget=None, max_attempts=12):
        self.pool = pool
        self.model = model
        self.api_base = api_base.rstrip("/")
        self.system_text = system_text
        self.reference_text = reference_text
        self.state_path = state_path
        self.use_cache = use_cache
        self.ttl = ttl
        self.temperature = temperature
        self.thinking_budget = thinking_budget
        self.max_attempts = max_attempts
        self.lock = threading.Lock()
        self.no_cache_keys = set()
        self.usage = {"requests": 0, "prompt": 0, "cached": 0, "output": 0}
        self.digest = hashlib.sha256(
            "\0".join((model, system_text, reference_text or "")).encode("utf-8")).hexdigest()[:16]
        self.state = {}
        if state_path and os.path.exists(state_path):
            try:
                with open(state_path, encoding="utf-8") as f:
                    self.state = json.load(f)
            except ValueError:
                self.state = {}

    # -- fixed prompt prefix ------------------------------------------------
    def _system(self):
        return {"parts": [{"text": self.system_text}]}

    def _prefix_turns(self):
        if not self.reference_text:
            return []
        return [
            {"role": "user", "parts": [{"text": self.reference_text}]},
            {"role": "model", "parts": [{"text": "Understood. I will follow this reference for every request."}]},
        ]

    # -- cache handling -----------------------------------------------------
    @staticmethod
    def _key_id(key):
        return hashlib.sha256(key.encode("utf-8")).hexdigest()[:16]   # never store the key itself

    def _save_state(self):
        if self.state_path:
            with open(self.state_path, "w", encoding="utf-8") as f:
                json.dump(self.state, f, indent=1)

    def _cache_for(self, key):
        if not self.use_cache or not self.reference_text:
            return None
        kid = self._key_id(key)
        with self.lock:
            if key in self.no_cache_keys:
                return None
            entry = self.state.get(kid)
            if entry and entry.get("digest") == self.digest and entry.get("expires", 0) > time.time() + 120:
                return entry["name"]
            # Caches belong to the key's project, so each key gets its own copy.
            body = {
                "model": f"models/{self.model}",
                "displayName": f"vntextpatch-{self.digest}",
                "systemInstruction": self._system(),
                "contents": self._prefix_turns(),
                "ttl": f"{int(self.ttl)}s",
            }
            try:
                created = _http_json("POST", f"{self.api_base}/cachedContents", key, body)
            except GeminiError as e:
                if e.status in (0, 429, 500, 502, 503, 504):
                    return None        # temporary: send the context inline this time
                self.no_cache_keys.add(key)
                print(f"  {self.pool.label(key)}: context cache unavailable ({e.message[:120]}); "
                      "sending context inline (implicit caching still applies)", flush=True)
                return None
            tokens = (created.get("usageMetadata") or {}).get("totalTokenCount", "?")
            print(f"  {self.pool.label(key)}: created context cache {created.get('name')} ({tokens} tokens)",
                  flush=True)
            self.state[kid] = {"name": created["name"], "digest": self.digest, "model": self.model,
                               "expires": time.time() + self.ttl}
            self._save_state()
            return created["name"]

    def _drop_cache(self, key):
        with self.lock:
            self.state.pop(self._key_id(key), None)
            self._save_state()

    def delete_caches(self):
        """Deletes the caches created for the given keys (stops storage billing)."""
        for key in self.pool.keys:
            entry = self.state.pop(self._key_id(key), None)
            if entry:
                try:
                    _http_json("DELETE", f"{self.api_base}/{entry['name']}", key)
                    print(f"  deleted {entry['name']}")
                except GeminiError as e:
                    print(f"  could not delete {entry['name']}: {e.message[:100]}")
        self._save_state()

    # -- generation ---------------------------------------------------------
    def generate_json(self, user_text, schema, use_context=True):
        last_error = None
        for attempt in range(self.max_attempts):
            key = self.pool.acquire()
            cache = self._cache_for(key) if use_context else None
            config = {"temperature": self.temperature, "responseMimeType": "application/json",
                      "responseSchema": schema}
            if self.thinking_budget is not None:
                config["thinkingConfig"] = {"thinkingBudget": self.thinking_budget}
            body = {
                "generationConfig": config,
                "safetySettings": [{"category": c, "threshold": "BLOCK_NONE"} for c in (
                    "HARM_CATEGORY_HARASSMENT", "HARM_CATEGORY_HATE_SPEECH",
                    "HARM_CATEGORY_SEXUALLY_EXPLICIT", "HARM_CATEGORY_DANGEROUS_CONTENT")],
            }
            user_turn = {"role": "user", "parts": [{"text": user_text}]}
            if cache:
                body["cachedContent"] = cache
                body["contents"] = [user_turn]
            else:
                body["systemInstruction"] = self._system()
                body["contents"] = (self._prefix_turns() if use_context else []) + [user_turn]

            try:
                resp = _http_json("POST", f"{self.api_base}/models/{self.model}:generateContent", key, body)
            except GeminiError as e:
                last_error = e
                text = (e.message or "").lower()
                if e.status == 429:
                    delay = _retry_delay(e)
                    print(f"  {self.pool.label(key)} hit its quota, resting {delay:.0f}s; switching key", flush=True)
                    self.pool.cooldown(key, delay)
                elif e.status in (0, 500, 502, 503, 504):
                    self.pool.cooldown(key, min(3 * 2 ** attempt, 60))
                elif cache and e.status in (400, 403, 404) and "cache" in text:
                    self._drop_cache(key)      # expired or deleted: recreate on the next try
                elif e.status in (401, 403) or "api key" in text or "api_key" in text:
                    self.pool.kill(key, e.message[:100])
                else:
                    raise
                continue

            usage = resp.get("usageMetadata") or {}
            with self.lock:
                self.usage["requests"] += 1
                self.usage["prompt"] += usage.get("promptTokenCount", 0)
                self.usage["cached"] += usage.get("cachedContentTokenCount", 0)
                self.usage["output"] += usage.get("candidatesTokenCount", 0)

            block = (resp.get("promptFeedback") or {}).get("blockReason")
            candidates = resp.get("candidates") or []
            finish = candidates[0].get("finishReason") if candidates else None
            if block or finish in ("SAFETY", "PROHIBITED_CONTENT", "RECITATION", "BLOCKLIST", "MAX_TOKENS"):
                raise ContentBlocked(block or finish)
            parts = ((candidates[0].get("content") or {}).get("parts") or []) if candidates else []
            text = "".join(p.get("text", "") for p in parts if not p.get("thought"))
            try:
                return json.loads(text)
            except ValueError:
                last_error = GeminiError(200, "model returned invalid JSON")
                continue
        raise RuntimeError(f"Gemini request failed after {self.max_attempts} attempts: {last_error}")


# -- prompts -----------------------------------------------------------------

SYSTEM_PROMPT = """You are a professional visual novel localizer translating a Ren'Py game from {source} into {target}.

Rules:
1. Translate meaning and tone, not word by word. Dialogue must sound like natural spoken {target}; narration like natural written prose.
2. Keep each character's voice consistent with the character sheet in the reference material: gender, age, personality, how they address each other, level of politeness.
3. Always use the glossary translations for names and terms. Never invent a different spelling for a name that is in the glossary.
4. Copy these verbatim, never translate or reorder their insides: Ren'Py interpolations in square brackets such as [player] or [s_name]; text tags in curly braces such as {{i}}...{{/i}}, {{w=0.5}}, {{nw}}, {{color=#f00}}; escape sequences such as \\n, \\" and %%. Text between an opening and a closing tag IS translated.
5. Keep leading/trailing whitespace, ellipses and stutters ("W-well...") in an equivalent form. Do not add quotation marks, notes or explanations.
6. Lines of type "name" are character names, "choice" are menu options (keep them short), "string" are interface texts (keep them short), "say" are dialogue or narration.
7. Lines made of deliberately corrupted or glitched text (random symbols/letters) must be returned unchanged.
8. Return exactly one result per input line, with the same id.
{language_rules}"""

BURMESE_RULES = """9. Write standard Burmese in Unicode (never Zawgyi encoding). Use colloquial spoken style (ပြောစကား) for dialogue - e.g. တယ်/မယ်/ဘူး endings - and do not switch to literary style (စာစကား: သည်/မည်/၏) except in poems, letters or formal writing inside the story.
10. Choose pronouns and sentence-final particles from the speaker's gender and their relationship with the listener (e.g. ကျွန်တော်/ခင်ဗျာ for male speakers, ကျွန်မ/ရှင် for female speakers, or casual ငါ/နင်/မင်း between close friends) and keep that choice consistent for each pair of characters throughout the game.
11. Transliterate personal names into Burmese script as given in the glossary; keep honorific habits natural for Burmese readers.
"""

CONTEXT_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "glossary": {"type": "ARRAY", "items": {
            "type": "OBJECT",
            "properties": {"source": {"type": "STRING"}, "translation": {"type": "STRING"},
                           "note": {"type": "STRING"}},
            "required": ["source", "translation"]}},
        "characters": {"type": "ARRAY", "items": {
            "type": "OBJECT",
            "properties": {"name": {"type": "STRING"}, "translation": {"type": "STRING"},
                           "gender": {"type": "STRING"}, "role": {"type": "STRING"},
                           "personality": {"type": "STRING"}, "speech_style": {"type": "STRING"},
                           "addresses_others_as": {"type": "STRING"}},
            "required": ["name", "translation", "gender", "speech_style"]}},
        "story_notes": {"type": "STRING"},
        "style_guide": {"type": "STRING"},
    },
    "required": ["glossary", "characters", "story_notes", "style_guide"],
}

TRANSLATION_SCHEMA = {
    "type": "ARRAY",
    "items": {"type": "OBJECT",
              "properties": {"id": {"type": "INTEGER"}, "text": {"type": "STRING"}},
              "required": ["id", "text"]},
}


def build_system_prompt(target, source):
    rules = BURMESE_RULES if re.search(r"burmese|myanmar|မြန်မာ", target, re.I) else ""
    return SYSTEM_PROMPT.format(source=source, target=target, language_rules=rules)


def render_reference(context):
    """The text that gets cached: glossary + character sheet + story notes."""
    lines = ["REFERENCE MATERIAL FOR THIS GAME - remember it and apply it to every translation request.", "",
             "## Glossary (source => translation)"]
    for g in context.get("glossary", []):
        note = f"   ({g['note']})" if g.get("note") else ""
        lines.append(f"- {g['source']} => {g['translation']}{note}")
    lines += ["", "## Characters"]
    for c in context.get("characters", []):
        lines.append(f"### {c.get('name')} => {c.get('translation')}")
        for field_name in ("gender", "role", "personality", "speech_style", "addresses_others_as"):
            if c.get(field_name):
                lines.append(f"- {field_name.replace('_', ' ')}: {c[field_name]}")
    lines += ["", "## Story notes", context.get("story_notes", ""), "", "## Style guide",
              context.get("style_guide", "")]
    if context.get("extra_notes"):
        lines += ["", "## Translator's extra notes", context["extra_notes"]]
    return "\n".join(lines)


# -- workbook helpers for translation ------------------------------------------

@dataclass
class Line:
    row: int
    speaker: str
    kind: str
    text: str
    translation: str = None
    note: str = None


def sheet_lines(ws):
    lines = []
    for row in range(2, ws.max_row + 1):
        text = _cell_text(ws, row, COL_ORIG_LINE)
        if text is None or text == EMPTY_TEXT_MARKER:
            continue
        note = (_cell_text(ws, row, COL_NOTES) or "").split()
        kind = note[0] if note and note[0] in KINDS else "say"
        lines.append(Line(row, _cell_text(ws, row, COL_ORIG_CHAR), kind, text, _cell_text(ws, row, COL_TL)))
    return lines


def _markup_signature(text):
    interpolations = sorted(re.findall(r"\[[^\[\]]+\]", text))
    tags = sorted(re.findall(r"\{[^{}]*\}", text))
    return interpolations, tags, text.count("\\n")


def build_context(client, wb, target, sample_chars):
    """Asks Gemini to read a sample of the script and write the translation bible."""
    speakers = {}
    name_lines = []
    samples = []
    per_sheet = max(sample_chars // max(len(wb.worksheets), 1), 1500)
    for ws in wb.worksheets:
        budget = per_sheet
        chunk = []
        for line in sheet_lines(ws):
            if line.speaker:
                speakers[line.speaker] = speakers.get(line.speaker, 0) + 1
            if line.kind == "name":
                name_lines.append(line.text)
            if budget > 0 and line.kind in ("say", "choice"):
                chunk.append(f"{line.speaker or '-'}: {line.text}")
                budget -= len(line.text) + 12
        if chunk:
            samples.append(f"### {ws.title}\n" + "\n".join(chunk))
    names = sorted(set(speakers) | set(name_lines))
    prompt = (
        f"Below is a sample of a visual novel script (format `speaker: line`, `-` means narration).\n"
        f"Prepare the reference material a translator needs to translate the whole game into {target} consistently.\n\n"
        f"Return:\n"
        f"- glossary: every character name listed under NAMES, plus recurring proper nouns, places, club/school "
        f"names and catchphrases you find, each with ONE fixed {target} rendering.\n"
        f"- characters: one entry per character: gender, role, personality, how they talk (speech_style: the "
        f"{target} pronouns, particles and politeness level they should use), and how they address the others.\n"
        f"- story_notes: setting, premise, relationships and anything a translator must remember "
        f"(running jokes, twists that change how lines should read). Write in English.\n"
        f"- style_guide: concrete {target} style decisions (register, punctuation, how to treat interjections, "
        f"honorifics, onomatopoeia). Write in English.\n\n"
        f"NAMES (speaker name, number of lines):\n"
        + "\n".join(f"- {n} ({speakers.get(n, 0)})" for n in names)
        + "\n\nSCRIPT SAMPLE:\n" + "\n\n".join(samples))
    context = client.generate_json(prompt, CONTEXT_SCHEMA, use_context=False)
    context["target_language"] = target
    return context


def translate_lines(client, lines, target, history, glossary, batch_size, batch_chars):
    """Translates the given lines of one sheet in order. Yields lists of finished Line objects."""
    todo = []
    for line in lines:
        if line.kind == "name" and line.text in glossary:
            line.translation = glossary[line.text]
            yield [line]
        else:
            todo.append(line)

    by_row = {l.row: l for l in history}
    ordered = sorted(by_row)

    def request(batch, attempt_note=""):
        first_row = batch[0].row
        before = [by_row[r] for r in ordered if r < first_row][-8:]
        payload = {
            "previous_lines_for_context": [
                {"speaker": l.speaker or "", "text": l.text, "translation": l.translation or ""} for l in before],
            "lines": [{"id": i, "type": l.kind, "speaker": l.speaker or "", "text": l.text}
                      for i, l in enumerate(batch)],
        }
        prompt = (f"Translate every entry of \"lines\" into {target}. \"previous_lines_for_context\" is only there "
                  f"for continuity; do not return it.{attempt_note}\n" + json.dumps(payload, ensure_ascii=False))
        result = client.generate_json(prompt, TRANSLATION_SCHEMA)
        got = {}
        for item in result if isinstance(result, list) else []:
            if isinstance(item, dict) and isinstance(item.get("id"), int) and isinstance(item.get("text"), str):
                got[item["id"]] = item["text"]
        return got

    def run(batch, depth=0):
        try:
            got = request(batch)
        except ContentBlocked as e:
            if len(batch) > 1:
                mid = len(batch) // 2
                return run(batch[:mid], depth + 1) + run(batch[mid:], depth + 1)
            batch[0].note = f"NOT TRANSLATED: blocked by Gemini ({e})"
            return batch
        retry = []
        for i, line in enumerate(batch):
            text = got.get(i)
            if text is None or not text.strip():
                retry.append(line)
            elif _markup_signature(text) != _markup_signature(line.text):
                line.translation, line.note = text, "CHECK: tags/placeholders differ from the original"
                retry.append(line)
            else:
                line.translation, line.note = text, None
        if retry and depth < 2:
            if len(retry) == len(batch) and len(batch) > 1:
                mid = len(batch) // 2
                run(batch[:mid], depth + 1)
                run(batch[mid:], depth + 1)
            else:
                run(retry, depth + 1) if len(retry) < len(batch) else retry_single(retry)
        return batch

    def retry_single(batch):
        note = (" IMPORTANT: copy every [placeholder], {tag} and \\n from the source text exactly, "
                "and return a non-empty translation.")
        try:
            got = request(batch, note)
        except ContentBlocked:
            return batch
        for i, line in enumerate(batch):
            text = got.get(i)
            if text and text.strip():
                ok = _markup_signature(text) == _markup_signature(line.text)
                line.translation = text
                line.note = None if ok else "CHECK: tags/placeholders differ from the original"
        return batch

    batch, size = [], 0
    for line in todo + [None]:
        if line is None or (batch and (len(batch) >= batch_size or size + len(line.text) > batch_chars)):
            if batch:
                done = run(batch)
                for l in done:
                    if not l.translation and not l.note:
                        l.note = "NOT TRANSLATED: no answer from Gemini"
                yield done
            batch, size = [], 0
        if line is not None:
            batch.append(line)
            size += len(line.text)


def default_side_path(xlsx_path, suffix):
    return os.path.splitext(os.path.abspath(xlsx_path))[0] + suffix


def cmd_translate(args):
    keys = load_keys(args)
    if not keys:
        sys.exit("No Gemini API keys. Use --keys-file FILE (one key per line), --keys k1,k2 or set GEMINI_API_KEYS.")
    pool = KeyPool(keys)
    wb = openpyxl.load_workbook(args.xlsx)
    styles = _column_styles(wb.worksheets[0])
    names = CharacterNames(args.names or default_names_path(args.xlsx))
    context_path = args.context or default_side_path(args.xlsx, ".context.json")
    state_path = default_side_path(args.xlsx, ".gemini_cache.json")
    system_text = build_system_prompt(args.language, args.source_language)
    common = dict(model=args.model, api_base=args.api_base, system_text=system_text, state_path=state_path,
                  use_cache=not args.no_cache, ttl=args.cache_ttl, temperature=args.temperature,
                  thinking_budget=args.thinking_budget)
    print(f"Using {len(keys)} API key(s), model {args.model}")

    # 1. the translation bible: names + important details, kept in a JSON file you can edit
    if os.path.exists(context_path) and not args.rebuild_context:
        with open(context_path, encoding="utf-8") as f:
            context = json.load(f)
        print(f"Loaded translation context from {context_path}")
    else:
        print("Building translation context (character names, personalities, story notes) ...")
        context = build_context(GeminiClient(pool, reference_text=None, **common), wb, args.language,
                                args.sample_chars)
        with open(context_path, "w", encoding="utf-8") as f:
            json.dump(context, f, ensure_ascii=False, indent=2)
        print(f"Saved translation context to {context_path} - edit it to fix names or add notes "
              "(\"extra_notes\"), the cache is rebuilt automatically when it changes.")

    glossary = {}
    for entry in list(context.get("characters", [])):
        if entry.get("name") and entry.get("translation"):
            glossary[entry["name"]] = entry["translation"]
    for entry in context.get("glossary", []):
        if entry.get("source") and entry.get("translation"):
            glossary[entry["source"]] = entry["translation"]
    # names.xml entries that the user translated by hand win over the model's choices
    for original, translation in names.translations.items():
        if translation != original:
            glossary[original] = translation
            for entry in context.get("glossary", []):
                if entry.get("source") == original:
                    entry["translation"] = translation
            for entry in context.get("characters", []):
                if entry.get("name") == original:
                    entry["translation"] = translation
    for ws in wb.worksheets:
        for row in range(2, ws.max_row + 1):
            speaker = _cell_text(ws, row, COL_ORIG_CHAR)
            if speaker and speaker in glossary:
                names.set(speaker, glossary[speaker])
                _set_cell(ws, row, COL_TL_CHAR, join_names([glossary[speaker]]), styles)
    names.save()

    if args.context_only:
        save_workbook(wb, args.xlsx)
        return 0

    # 2. translate, sheet by sheet
    client = GeminiClient(pool, reference_text=render_reference(context), **common)
    only = set(args.sheets.split(",")) if args.sheets else None
    jobs = []
    for ws in wb.worksheets:
        if only and ws.title not in only:
            continue
        lines = sheet_lines(ws)
        pending = [l for l in lines if args.retranslate or not l.translation]
        if args.limit:
            pending = pending[:args.limit]
        if pending:
            jobs.append((ws, lines, pending))
    total = sum(len(p) for _ws, _l, p in jobs)
    if not total:
        print("Nothing to translate: every line already has a TL entry (use --retranslate to redo).")
        return 0
    print(f"Translating {total} lines in {len(jobs)} sheets into {args.language} ...")

    book_lock = threading.Lock()
    progress = {"done": 0, "since_save": 0}

    def work(ws, lines, pending):
        for finished in translate_lines(client, pending, args.language, lines, glossary,
                                        args.batch_size, args.batch_chars):
            with book_lock:
                for line in finished:
                    if line.translation:
                        _set_cell(ws, line.row, COL_TL, line.translation, styles)
                    if line.note:
                        _set_cell(ws, line.row, COL_NOTES, f"{line.kind} {line.note}", styles)
                    elif line.translation:
                        _set_cell(ws, line.row, COL_NOTES, line.kind, styles)
                progress["done"] += len(finished)
                progress["since_save"] += len(finished)
                print(f"  [{progress['done']}/{total}] {ws.title}", flush=True)
                if progress["since_save"] >= args.save_every:
                    save_workbook(wb, args.xlsx)
                    progress["since_save"] = 0
        return ws.title

    failed = []
    workers = max(1, min(args.workers or len(keys), len(jobs), 8))
    try:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(work, *job): job[0].title for job in jobs}
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:   # keep the other sheets going, report at the end
                    failed.append(f"{futures[future]}: {e}")
    finally:
        with book_lock:
            save_workbook(wb, args.xlsx)

    u = client.usage
    print(f"\nDone: {progress['done']}/{total} lines. Gemini usage: {u['requests']} requests, "
          f"{u['prompt']} prompt tokens ({u['cached']} served from cache), {u['output']} output tokens.")
    if args.delete_cache:
        client.delete_caches()
    if failed:
        print("Some sheets stopped early (run the command again to resume):", file=sys.stderr)
        for f in failed:
            print("  " + f, file=sys.stderr)
        return 1
    return 0


# ----------------------------------------------------------------------------
# command line
# ----------------------------------------------------------------------------

def add_parser_options(p):
    p.add_argument("--mode", choices=("dialogue", "all"), default="dialogue",
                   help="dialogue: only player-visible text (default). all: every string literal, "
                        "exactly like VNTextPatch's RenpyScript.cs")
    p.add_argument("--extra", action="append", default=[], metavar="REGEX",
                   help="also extract strings whose line prefix (the text before the opening quote) matches "
                        "REGEX, e.g. --extra \"(title|text)\\s*=\\s*$\" for keyword arguments. Repeatable; "
                        "pass the same values to insert.")
    p.add_argument("--unrpyc", metavar="PATH", help="unrpyc.py / unrpyc executable for archives holding only .rpyc")


def add_gemini_options(p):
    p.add_argument("--keys-file", help="text file with one Gemini API key per line")
    p.add_argument("--keys", help="comma separated API keys (or set GEMINI_API_KEYS)")
    p.add_argument("--model", default="gemini-2.5-flash")
    p.add_argument("--language", default="Burmese", help="target language (default: Burmese)")
    p.add_argument("--source-language", default="English")
    p.add_argument("--context", metavar="JSON", help="translation context file (default: <xlsx>.context.json)")
    p.add_argument("--rebuild-context", action="store_true", help="regenerate the context file")
    p.add_argument("--context-only", action="store_true", help="only build the context / names, don't translate")
    p.add_argument("--sample-chars", type=int, default=120000, help="how much script to show Gemini for the context")
    p.add_argument("--no-cache", action="store_true", help="don't create explicit context caches")
    p.add_argument("--cache-ttl", type=int, default=3600, help="cache lifetime in seconds (default 3600)")
    p.add_argument("--delete-cache", action="store_true", help="delete the context caches when finished")
    p.add_argument("--batch-size", type=int, default=40, help="lines per request")
    p.add_argument("--batch-chars", type=int, default=6000, help="max source characters per request")
    p.add_argument("--workers", type=int, default=0, help="parallel sheets (default: number of keys, max 8)")
    p.add_argument("--sheets", help="comma separated sheet names to translate (default: all)")
    p.add_argument("--limit", type=int, default=0, help="translate at most N lines per sheet (for a test run)")
    p.add_argument("--retranslate", action="store_true", help="overwrite existing TL cells")
    p.add_argument("--temperature", type=float, default=0.3)
    p.add_argument("--thinking-budget", type=int, default=None, help="thinkingBudget for models that support it")
    p.add_argument("--save-every", type=int, default=200, help="save the workbook every N lines")
    p.add_argument("--api-base", default="https://generativelanguage.googleapis.com/v1beta")
    p.add_argument("--names", metavar="XML", help="names.xml path (default: next to the workbook)")


def ask_what_next(args):
    """The menu shown after extraction."""
    print("\nWhat do you want to do next?")
    print("  1) Translate into Burmese with Gemini (rotating API keys + context cache)")
    print("  2) Translate into another language with Gemini")
    print("  3) Only build the translation context (names, character notes) for review")
    print("  4) Nothing - I'll translate the workbook myself")
    choice = input("Choice [4]: ").strip() or "4"
    if choice not in ("1", "2", "3"):
        return 0
    argv = ["translate", args.xlsx]
    if choice == "2":
        argv += ["--language", input("Target language: ").strip() or "Burmese"]
    if choice == "3":
        argv += ["--context-only"]
    if not load_keys(argparse.Namespace()):
        answer = input("Path to a file with your Gemini API keys (one per line), or paste keys separated by commas: ").strip()
        argv += ["--keys-file", answer] if os.path.exists(answer) else ["--keys", answer]
    if args.names:
        argv += ["--names", args.names]
    return main(argv)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Extract Ren'Py script text from .rpa archives/folders into an .xlsx (VNTextPatch layout), "
                    "translate it with Gemini, and patch it back.")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("extract", help="archives/folder -> script.xlsx")
    p.add_argument("input", help="game folder, .rpa archive or folder of .rpy files")
    p.add_argument("xlsx")
    p.add_argument("--template", help="template.xlsx (default: the one shipped with VNTextPatch)")
    p.add_argument("--names", metavar="XML", help="names.xml path (default: next to the workbook)")
    p.add_argument("--dump-scripts", metavar="DIR", help="also write the scripts pulled from the archives to DIR")
    p.add_argument("--translate", action="store_true", help="run the Gemini translation right after extracting")
    p.add_argument("--no-prompt", action="store_true", help="don't ask what to do after extracting")
    add_parser_options(p)

    p = sub.add_parser("translate", help="fill the TL column of script.xlsx with Gemini")
    p.add_argument("xlsx")
    add_gemini_options(p)

    p = sub.add_parser("insert", help="archives/folder + script.xlsx -> patched scripts")
    p.add_argument("input", help="the same input that was given to extract")
    p.add_argument("xlsx")
    p.add_argument("output", help="folder to receive the patched files")
    p.add_argument("--repack", action="store_true",
                   help="rebuild each input .rpa (with the patched scripts) in the output folder "
                        "instead of writing loose .rpy files")
    add_parser_options(p)

    args, rest = parser.parse_known_args(argv)
    if args.command == "extract":
        if rest and not args.translate:
            parser.error("unrecognized arguments: " + " ".join(rest))
        code = cmd_extract(args)
        if code == 0 and args.translate:
            return main(["translate", args.xlsx] + (["--names", args.names] if args.names else []) + rest)
        if code == 0 and not args.no_prompt and sys.stdin.isatty():
            return ask_what_next(args)
        return code
    if rest:
        parser.error("unrecognized arguments: " + " ".join(rest))
    if args.command == "translate":
        return cmd_translate(args)
    return cmd_insert(args)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit("\nInterrupted. Progress is saved in the workbook; run the same command to resume.")
    except (OSError, ValueError, RuntimeError) as e:
        sys.exit(f"Error: {e}")
