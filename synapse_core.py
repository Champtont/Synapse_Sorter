import copy
import io
import os
import re

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# ----------------------------
# Patterns
# ----------------------------

# A question starts with its number, e.g. "12. ..." or "12) ..."
QUESTION_RE = re.compile(r'^(\s*)(\d+)(\s*)([\.\)])')
# Lines starting with this mark are section titles and are never copied
SECTION_MARK = "*"

R_ATTRS = (qn('r:embed'), qn('r:link'), qn('r:id'))


# ----------------------------
# Number Input
# ----------------------------

def parse_number_input(text):
    """Turn '20-26, 3, 8' into [20, ..., 26, 3, 8], keeping the order typed."""
    numbers = []
    for part in text.split(","):
        part = part.strip()
        if not part: continue
        if "-" in part:
            try:
                start, end = (int(x) for x in part.split("-"))
            except ValueError: continue
            step = 1 if end >= start else -1
            numbers.extend(range(start, end + step, step))
        else:
            try: numbers.append(int(part))
            except ValueError: continue
    return numbers


def _cell_to_number(value, coordinate):
    if isinstance(value, bool):
        raise ValueError(f"Cell {coordinate}: '{value}' is not a question number")
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if isinstance(value, str) and value.strip().isdigit():
        return int(value.strip())
    raise ValueError(f"Cell {coordinate}: '{value}' is not a question number")


def read_permutations(xlsx_path):
    """Read an Excel sheet where each column is one exam version (permuta).

    Row 1 holds the version names (e.g. 'Permuta_1'), the rows below hold the
    question numbers in the order they should appear. Returns a list of
    (name, [numbers]) tuples, one per non-empty column.
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows())
    finally:
        wb.close()
    if not rows:
        return []

    header = rows[0]
    has_header = any(isinstance(c.value, str) and not c.value.strip().isdigit() for c in header)
    data_rows = rows[1:] if has_header else rows

    permutations = []
    for col in range(max(len(r) for r in rows)):
        name = None
        if has_header and col < len(header) and header[col].value is not None:
            name = str(header[col].value).strip()
        numbers = []
        for row in data_rows:
            if col >= len(row) or row[col].value is None: continue
            if isinstance(row[col].value, str) and not row[col].value.strip(): continue
            numbers.append(_cell_to_number(row[col].value, row[col].coordinate))
        if numbers:
            permutations.append((name or f"Permuta_{len(permutations) + 1}", numbers))
    return permutations


# ----------------------------
# Document Processing
# ----------------------------

def _text(el):
    return "".join(t.text or "" for t in el.iter(qn('w:t')))


def _has_list_numbering(p):
    ppr = p.find(qn('w:pPr'))
    return ppr is not None and ppr.find(qn('w:numPr')) is not None


def _is_blank(el):
    if _text(el).strip():
        return False
    return next(el.iter(qn('w:drawing'), qn('w:pict'), qn('w:object')), None) is None


class QuestionBank:
    """Where each numbered question lives inside the source .docx.

    Questions are stored as positions of body elements, so every exam can be
    built from a fresh copy of the source and keep its styles, list
    numbering and images.
    """

    def __init__(self, path):
        self.path = path
        self.questions = {}     # number -> [body element index, ...]
        self.duplicates = []    # numbers that appear more than once in the bank

        body = Document(path).element.body
        current = None
        for i, el in enumerate(body.iterchildren()):
            if el.tag not in (qn('w:p'), qn('w:tbl')): continue
            text = _text(el)

            if el.tag == qn('w:p') and text.strip().startswith(SECTION_MARK):
                current = None
                continue

            match = QUESTION_RE.match(text) if el.tag == qn('w:p') and not _has_list_numbering(el) else None
            if match:
                current = int(match.group(2))
                if current in self.questions:
                    self.duplicates.append(current)
                self.questions[current] = [i]
                continue

            if current is not None:
                self.questions[current].append(i)

        # Trailing blank paragraphs are spacing, not part of the question
        children = list(body.iterchildren())
        for indexes in self.questions.values():
            while len(indexes) > 1 and _is_blank(children[indexes[-1]]):
                indexes.pop()

    def __len__(self):
        return len(self.questions)

    def is_empty(self, number, children):
        indexes = self.questions[number]
        stem = QUESTION_RE.sub("", _text(children[indexes[0]]), count=1)
        return not stem.strip() and all(_is_blank(children[i]) for i in indexes[1:])


def _renumber(p, new_number):
    """Replace the leading question number, which Word often splits over several runs."""
    t_nodes = list(p.iter(qn('w:t')))
    match = QUESTION_RE.match("".join(t.text or "" for t in t_nodes))
    if not match or not t_nodes: return
    to_remove = match.end()
    for t in t_nodes:
        if to_remove <= 0: break
        text = t.text or ""
        cut = min(len(text), to_remove)
        t.text = text[cut:]
        # Without this Word drops the space that now starts the text ("1.Adaptive")
        t.set(qn('xml:space'), 'preserve')
        to_remove -= cut
    t_nodes[0].text = f"{new_number}{match.group(4)}{t_nodes[0].text}"
    t_nodes[0].set(qn('xml:space'), 'preserve')


def _restart_lists(elements, numbering):
    """Give each copied question its own auto-numbered list so answers restart at a."""
    remap = {}
    for num_id in (n for el in elements for n in el.iter(qn('w:numId'))):
        old = num_id.get(qn('w:val'))
        if old == "0": continue
        if old not in remap:
            try:
                source = numbering.num_having_numId(int(old))
            except KeyError:
                continue
            new_num = numbering.add_num(source.abstractNumId.val)
            for ilvl in range(9):
                new_num.add_lvlOverride(ilvl=ilvl).add_startOverride(_level_start(numbering, source.abstractNumId.val, ilvl))
            remap[old] = str(new_num.numId)
        num_id.set(qn('w:val'), remap[old])


def _level_start(numbering, abstract_id, ilvl):
    found = numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]/w:lvl[@w:ilvl="{ilvl}"]/w:start/@w:val')
    return int(found[0]) if found else 1


def _is_red(hex_value):
    if not hex_value or len(hex_value) != 6: return False
    try:
        r, g, b = (int(hex_value[i:i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return False
    return r >= 0xC0 and g < 0x50 and b < 0x50


def _remove_red(elements):
    """Strip the red colour used to mark correct answers (for student copies)."""
    for el in elements:
        for color in list(el.iter(qn('w:color'))):
            if _is_red(color.get(qn('w:val'))):
                color.getparent().remove(color)


def _drop_unused_images(doc):
    used = {v for el in doc.element.body.iter() for k, v in el.attrib.items() if k in R_ATTRS}
    for r_id, rel in list(doc.part.rels.items()):
        if rel.reltype == RT.IMAGE and r_id not in used:
            doc.part.rels.pop(r_id)


def _add_title(doc, title):
    try:
        p = doc.add_heading(title, level=0)
    except KeyError:  # source has no Title style
        p = doc.add_paragraph()
        run = p.add_run(title)
        run.bold = True
        run.font.size = Pt(16)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _keep_together(elements):
    """Ask Word to move a question to the next page rather than split it.

    Every paragraph keeps its lines together and stays with the next one,
    except the last, so Word can only break the page between questions.
    """
    paragraphs = [el for el in elements if el.tag == qn('w:p')]
    for i, p in enumerate(paragraphs):
        ppr = p.get_or_add_pPr()
        ppr.keepLines_val = True
        ppr.keepNext_val = i < len(paragraphs) - 1 or None


# ----------------------------
# Exam Header
# ----------------------------

# Yellow text containing this word is the version spot, filled from the Excel column name
VERSION_RE = re.compile(r'permuta', re.IGNORECASE)


def _is_yellow(run):
    highlight = run.find(f"{qn('w:rPr')}/{qn('w:highlight')}")
    return highlight is not None and highlight.get(qn('w:val')) == "yellow"


def _run_text(run):
    parts = []
    for child in run:
        if child.tag == qn('w:t'): parts.append(child.text or "")
        elif child.tag == qn('w:tab'): parts.append("\t")
        elif child.tag in (qn('w:br'), qn('w:cr')): parts.append("\n")
    return "".join(parts)


def _yellow_groups(p):
    """Runs of yellow text that sit next to each other form one field."""
    groups, current = [], []
    for run in p.iter(qn('w:r')):
        if next(run.iterancestors(qn('w:p'))) is not p: continue  # text boxes are handled as their own paragraphs
        if _is_yellow(run):
            current.append(run)
        elif current:
            groups.append(current)
            current = []
    if current: groups.append(current)
    return groups


def _replace_group(group, text):
    first = group[0]
    for child in list(first):
        if child.tag != qn('w:rPr'):
            first.remove(child)
    highlight = first.find(f"{qn('w:rPr')}/{qn('w:highlight')}")
    highlight.getparent().remove(highlight)
    for i, line in enumerate(text.split("\n")):
        if i: first.append(OxmlElement('w:br'))
        for j, piece in enumerate(line.split("\t")):
            if j: first.append(OxmlElement('w:tab'))
            if piece:
                t = OxmlElement('w:t')
                t.text = piece
                t.set(qn('xml:space'), 'preserve')
                first.append(t)
    for run in group[1:]:
        run.getparent().remove(run)


def _version_text(template, version):
    """'Permuta<tab>1' + 'Permuta_3' -> 'Permuta<tab>3', keeping the header's own spacing."""
    if not version:
        return ""
    new_number = re.search(r'(\d+)\s*$', version)
    old_number = re.search(r'\d+(?=\s*$)', template)
    if new_number and old_number:
        return template[:old_number.start()] + new_number.group(1) + template[old_number.end():]
    return version.replace("_", " ")


class HeaderChangedError(Exception):
    pass


class ExamHeader:
    """The text placed at the top of every exam, read from a Word document.

    Yellow-highlighted text marks the parts that change per exam. Each one
    becomes a field; the one containing 'Permuta' is filled automatically.
    """

    def __init__(self, path):
        self.path = path
        self.fields = []    # [{"text": original text, "version": bool}, ...] in document order
        for p in Document(path).element.body.iter(qn('w:p')):
            for group in _yellow_groups(p):
                text = "".join(_run_text(r) for r in group)
                self.fields.append({"text": text, "version": bool(VERSION_RE.search(text))})

    @property
    def has_version(self):
        return any(f["version"] for f in self.fields)

    @property
    def inputs(self):
        """(index, original text) of every field the user fills in."""
        return [(i, f["text"]) for i, f in enumerate(self.fields) if not f["version"]]

    def filled(self, values, version=""):
        """Return (document, body elements) with every field filled in.

        `values` maps field index to text; fields left out keep their text.
        """
        doc = Document(self.path)
        groups = [g for p in list(doc.element.body.iter(qn('w:p'))) for g in _yellow_groups(p)]
        # Values are matched to fields by position, so the file must not have changed since loading
        if ["".join(_run_text(r) for r in g) for g in groups] != [f["text"] for f in self.fields]:
            raise HeaderChangedError(f"{os.path.basename(self.path)} was changed after it was loaded. Load the exam header again.")
        for i, (field, group) in enumerate(zip(self.fields, groups)):
            text = _version_text(field["text"], version) if field["version"] else values.get(i, field["text"])
            _replace_group(group, text)
        elements = [el for el in doc.element.body.iterchildren() if el.tag in (qn('w:p'), qn('w:tbl'))]
        return doc, elements


def _copy_style(style_id, source_styles, target_styles):
    if target_styles.xpath(f'./w:style[@w:styleId="{style_id}"]'): return
    found = source_styles.xpath(f'./w:style[@w:styleId="{style_id}"]')
    if not found: return
    target_styles.append(copy.deepcopy(found[0]))
    for based_on in found[0].xpath('./w:basedOn/@w:val | ./w:link/@w:val | ./w:next/@w:val'):
        _copy_style(based_on, source_styles, target_styles)


def _import_header(doc, header_doc, elements):
    """Make the header's styles, lists and images work inside the exam document."""
    source_styles, target_styles = header_doc.styles.element, doc.styles.element
    for el in elements:
        for ref in el.iter(qn('w:pStyle'), qn('w:rStyle'), qn('w:tblStyle')):
            _copy_style(ref.get(qn('w:val')), source_styles, target_styles)

    numbering = doc.part.numbering_part.element
    remap = {}
    for num_id in (n for el in elements for n in el.iter(qn('w:numId'))):
        old = num_id.get(qn('w:val'))
        if old == "0": continue
        if old not in remap:
            source_numbering = header_doc.part.numbering_part.element
            source_abstract_id = source_numbering.num_having_numId(int(old)).abstractNumId.val
            abstract = copy.deepcopy(source_numbering.xpath(f'./w:abstractNum[@w:abstractNumId="{source_abstract_id}"]')[0])
            new_abstract_id = max([int(a) for a in numbering.xpath('./w:abstractNum/@w:abstractNumId')] + [0]) + 1
            abstract.set(qn('w:abstractNumId'), str(new_abstract_id))
            first_num = numbering.find(qn('w:num'))
            if first_num is not None: first_num.addprevious(abstract)
            else: numbering.append(abstract)
            remap[old] = str(numbering.add_num(new_abstract_id).numId)
        num_id.set(qn('w:val'), remap[old])

    for el in elements:
        for node in el.iter():
            for attr in R_ATTRS:
                r_id = node.get(attr)
                if r_id is None or r_id not in header_doc.part.rels: continue
                rel = header_doc.part.rels[r_id]
                if rel.reltype == RT.IMAGE:
                    new_id, _ = doc.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                elif rel.is_external:
                    new_id = doc.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                else:
                    continue
                node.set(attr, new_id)


def build_exam(bank, numbers, title, out_path, remove_answer_marks=False, on_progress=None,
               header=None, header_values=None, version=""):
    """Write an exam with the given questions, in the given order, renumbered from 1.

    With a `header`, its text replaces the title at the top, filled in with
    `header_values` and `version` (see ExamHeader.filled).

    Returns a report dict: 'missing' (not in the bank), 'repeated' (asked for
    more than once; only the first is used) and 'empty' (no text in the bank).
    """
    report = {"missing": [], "repeated": [], "empty": []}
    doc = Document(bank.path)
    body = doc.element.body
    children = list(body.iterchildren())
    numbering = doc.part.numbering_part.element

    sect_pr = body.find(qn('w:sectPr'))
    for el in children:
        if el is not sect_pr:
            body.remove(el)
    if header:
        header_doc, header_elements = header.filled(header_values or {}, version)
        _import_header(doc, header_doc, header_elements)
        for el in header_elements:
            if sect_pr is not None: sect_pr.addprevious(el)
            else: body.append(el)
    else:
        _add_title(doc, title)

    seen = set()
    new_number = 1
    for step, n in enumerate(numbers, start=1):
        if n in seen:
            report["repeated"].append(n)
        elif n not in bank.questions:
            report["missing"].append(n)
        else:
            seen.add(n)
            if bank.is_empty(n, children):
                report["empty"].append(n)
            copies = [copy.deepcopy(children[i]) for i in bank.questions[n]]
            _renumber(copies[0], new_number)
            _restart_lists(copies, numbering)
            _keep_together(copies)
            if remove_answer_marks:
                _remove_red(copies)
            for el in copies + [OxmlElement('w:p')]:
                if sect_pr is not None: sect_pr.addprevious(el)
                else: body.append(el)
            new_number += 1
        if on_progress: on_progress(step)

    _drop_unused_images(doc)
    doc.save(out_path)
    return report


def safe_filename(text):
    return re.sub(r'[<>:"/\\|?*]+', "_", text).strip(" .") or "Exam"


def describe_report(report):
    lines = []
    if report["missing"]:
        lines.append("Not found in the question bank: " + ", ".join(map(str, report["missing"])))
    if report["repeated"]:
        lines.append("Listed more than once (used only once): " + ", ".join(map(str, report["repeated"])))
    if report["empty"]:
        lines.append("Empty in the question bank: " + ", ".join(map(str, report["empty"])))
    return "\n".join(lines)
