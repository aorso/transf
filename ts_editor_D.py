"""
TermSheetTransformer - Version D ultime

Transforme des documents Word (term sheets) avec :
- Version mark-up : modifications visibles (jaune + barré)
- Version finale : document modifié sans marquage
- Version PDF : conversion du document final

Règles de marquage :
- Remplacement (Mot1 → Mot2) : Mot1 barré + Mot2 surligné jaune
- Ajout (nouvelle section) : surligné jaune
- Modification (description) : surligné jaune
- Suppression : texte barré (conservé pour trace)
"""

import subprocess
import shutil
import sys
import tempfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches


# -----------------------------------------------------------------------------
# Conversion .doc → .docx
# -----------------------------------------------------------------------------

def _strip_all_highlights(doc: Document, keep_red: bool = True) -> None:
    """Supprime tout surlignage du document. Si keep_red=True, conserve le surlignage rouge."""
    def clear_run(r):
        if keep_red and r.font.highlight_color == WD_COLOR_INDEX.RED:
            return
        r.font.highlight_color = None

    for p in doc.paragraphs:
        for r in p.runs:
            clear_run(r)
    for table in doc.tables:
        _strip_highlights_in_table(table, clear_run)


def _strip_highlights_in_table(table, clear_run) -> None:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    clear_run(r)
            for nested in cell.tables:
                _strip_highlights_in_table(nested, clear_run)


def _convert_doc_to_docx(doc_path: Path) -> Tuple[Path, Path]:
    """Convertit .doc en .docx. Essaie Word (Windows) puis LibreOffice."""
    tmpdir = Path(tempfile.mkdtemp(prefix="ts_transformer_"))
    docx_path = tmpdir / (doc_path.stem + ".docx")

    # 1. Windows : Microsoft Word via COM
    if sys.platform == "win32":
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(doc_path.resolve()))
            doc.SaveAs2(str(docx_path), FileFormat=16)  # 16 = wdFormatDocumentDefault (.docx)
            doc.Close()
            word.Quit()
            if docx_path.exists():
                return docx_path, tmpdir
        except Exception:
            pass

    # 2. LibreOffice (soffice / libreoffice)
    cmd = None
    for exe in ("soffice", "libreoffice"):
        found = shutil.which(exe)
        if found:
            cmd = [found, "--headless", "--convert-to", "docx", "--outdir", str(tmpdir), str(doc_path)]
            break
    if cmd is None and Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists():
        cmd = ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless",
               "--convert-to", "docx", "--outdir", str(tmpdir), str(doc_path)]
    if cmd is not None:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode == 0 and docx_path.exists():
            return docx_path, tmpdir

    shutil.rmtree(tmpdir, ignore_errors=True)
    raise RuntimeError(
        "Conversion .doc impossible : ni Word (Windows) ni LibreOffice détecté. "
        "Convertissez manuellement en .docx (Word : Fichier > Enregistrer sous > Word Document)."
    )


# -----------------------------------------------------------------------------
# Enregistrements des modifications
# -----------------------------------------------------------------------------

@dataclass
class ReplaceOp:
    old: str
    new: str
    occurrence: int = 1


@dataclass
class AddSectionOp:
    after_section: str
    title: str
    description: str
    occurrence: int = 1


@dataclass
class UpdateDescriptionOp:
    section_title: str
    new_description: str
    occurrence: int = 1


@dataclass
class DeleteSectionOp:
    section_title: str
    occurrence: int = 1


@dataclass
class AddLogoOp:
    logo_path: str
    width_inches: float = 1.0
    all_sections: bool = True


@dataclass
class AddContentOp:
    """Ajoute un titre ou une description (paragraphe hors tableau)."""
    text: str
    after_paragraph: Optional[str] = None  # Texte du paragraphe après lequel insérer (None = fin)
    red_highlight_in_final: bool = False


@dataclass
class RemoveParagraphOp:
    """Supprime un paragraphe contenant le texte donné."""
    text_contains: str
    occurrence: int = 1


@dataclass
class UpdateParagraphOp:
    """Modifie le contenu d'un paragraphe."""
    text_contains: str
    new_text: str
    occurrence: int = 1


# -----------------------------------------------------------------------------
# Éditeur de base
# -----------------------------------------------------------------------------

class _TermSheetEditor:
    """
    Éditeur interne pour term sheets.
    markup_mode=True : applique surlignage jaune et barré sur les modifications.
    """

    def __init__(self, doc: Document, markup_mode: bool = False):
        self.doc = doc
        self.markup_mode = markup_mode

    def replace_text(self, old: str, new: str):
        for p in self.doc.paragraphs:
            self._replace_in_paragraph_runs(p, old, new)
        for t in self.doc.tables:
            self._replace_in_table(t, old, new)
        return self

    def _replace_in_table(self, table, old: str, new: str):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    self._replace_in_paragraph_runs(p, old, new)
                for nested in cell.tables:
                    self._replace_in_table(nested, old, new)

    def _replace_in_paragraph_runs(self, paragraph, old: str, new: str):
        runs = list(paragraph.runs)
        if not runs:
            return

        full_text = "".join(r.text for r in runs)
        if old not in full_text:
            return

        char_to_run = []
        for i, r in enumerate(runs):
            char_to_run.extend([i] * len(r.text))

        segments = []
        i = 0
        L = len(full_text)
        old_len = len(old)

        while i < L:
            if full_text.startswith(old, i):
                start_run_idx = char_to_run[i] if i < len(char_to_run) else 0
                if self.markup_mode:
                    segments.append(("strike", old, start_run_idx))
                    segments.append(("highlight", new, start_run_idx))
                else:
                    segments.append((new, start_run_idx))
                i += old_len
            else:
                run_idx = char_to_run[i] if i < len(char_to_run) else 0
                segments.append((full_text[i], run_idx))
                i += 1

        merged = []
        for seg in segments:
            if seg[0] in ("strike", "highlight"):
                merged.append(seg)
            else:
                text, ridx = seg
                if merged and merged[-1][0] not in ("strike", "highlight") and merged[-1][1] == ridx:
                    merged[-1] = (merged[-1][0] + text, ridx)
                else:
                    merged.append(seg)

        for r in runs:
            r._element.getparent().remove(r._element)

        for seg in merged:
            if seg[0] == "strike":
                _, text, ridx = seg
                run = paragraph.add_run(text)
                self._copy_run_format(runs[ridx], run)
                run.font.strike = True
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif seg[0] == "highlight":
                _, text, ridx = seg
                run = paragraph.add_run(text)
                self._copy_run_format(runs[ridx], run)
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            else:
                text, ridx = seg
                if text == "":
                    continue
                run = paragraph.add_run(text)
                self._copy_run_format(runs[ridx], run)

    def _copy_run_format(self, src_run, dst_run):
        """Copie le format via XML (rPr) pour préserver tous les styles, y compris hérités."""
        src_rpr = src_run._element.find(qn("w:rPr"))
        if src_rpr is not None:
            dst_rpr = dst_run._element.find(qn("w:rPr"))
            if dst_rpr is not None:
                dst_run._element.remove(dst_rpr)
            dst_run._element.insert(0, deepcopy(src_rpr))
        else:
            # Fallback : copie propriété par propriété
            dst_run.bold = src_run.bold
            dst_run.italic = src_run.italic
            dst_run.underline = src_run.underline
            dst_run.font.name = src_run.font.name
            dst_run.font.size = src_run.font.size
            if src_run.font.color is not None and src_run.font.color.rgb is not None:
                dst_run.font.color.rgb = src_run.font.color.rgb
        if not self.markup_mode:
            dst_run.font.highlight_color = None

    def update_section_description(self, section_title: str, new_description: str, occurrence: int = 1):
        row, table = self._find_section_row(section_title, occurrence)
        if row is None:
            raise ValueError(f"Section introuvable : {section_title}")
        if len(row.cells) >= 2:
            self._set_cell_text(row.cells[1], new_description, highlight=self.markup_mode)
        elif len(row.cells) == 1:
            title = self._normalize(row.cells[0].text)
            self._set_cell_text(row.cells[0], f"{title}: {new_description}", highlight=self.markup_mode)
        return self

    def delete_section(self, section_title: str, occurrence: int = 1):
        row, table = self._find_section_row(section_title, occurrence)
        if row is None:
            raise ValueError(f"Section introuvable : {section_title}")
        if self.markup_mode:
            self._strike_row(row)
        else:
            table._tbl.remove(row._tr)
        return self

    def _strike_row(self, row):
        """Barré tout le contenu d'une ligne (mode mark-up pour suppression)."""
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.strike = True
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW

    def insert_section_after(self, after_section: str, new_title: str, new_description: str, occurrence: int = 1):
        ref_row, table = self._find_section_row(after_section, occurrence)
        if ref_row is None:
            raise ValueError(f"Section de référence introuvable : {after_section}")
        new_row = self._create_minimal_row_after(table, ref_row, new_title, new_description)
        return self

    def _set_cell_text(self, cell, text: str, highlight: bool = False, red_highlight: bool = False):
        if not cell.paragraphs:
            cell.add_paragraph("")
        p = cell.paragraphs[0]
        if p.runs:
            ref_run = p.runs[0]
            ref_run.text = text
            if highlight:
                ref_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif red_highlight:
                ref_run.font.highlight_color = WD_COLOR_INDEX.RED
            elif not self.markup_mode:
                ref_run.font.highlight_color = None
            for r in p.runs[1:]:
                r.text = ""
            for extra_p in cell.paragraphs[1:]:
                extra_p._element.getparent().remove(extra_p._element)
        else:
            r = p.add_run(text)
            if highlight:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif red_highlight:
                r.font.highlight_color = WD_COLOR_INDEX.RED
            elif not self.markup_mode:
                r.font.highlight_color = None

    def _normalize(self, s: str) -> str:
        return " ".join((s or "").replace("\n", " ").split()).strip()

    def _find_section_row(self, section_title: str, occurrence: int = 1):
        target = self._normalize(section_title)
        count = 0
        for table in self.doc.tables:
            row = self._find_section_row_in_table(table, target, occurrence, count)
            if row is not None:
                return row, table
            count += self._count_occurrences_in_table(table, target)
        return None, None

    def _count_occurrences_in_table(self, table, target: str) -> int:
        c = 0
        for row in table.rows:
            if len(row.cells) >= 1 and self._normalize(row.cells[0].text) == target:
                c += 1
            for cell in row.cells:
                for nested in cell.tables:
                    c += self._count_occurrences_in_table(nested, target)
        return c

    def _find_section_row_in_table(self, table, target: str, occurrence: int, count_so_far: int = 0):
        count = count_so_far
        for row in table.rows:
            if len(row.cells) >= 1 and self._normalize(row.cells[0].text) == target:
                count += 1
                if count == occurrence:
                    return row
            for cell in row.cells:
                for nested in cell.tables:
                    found = self._find_section_row_in_table(nested, target, occurrence, count)
                    if found is not None:
                        return found
                    count += self._count_occurrences_in_table(nested, target)
        return None

    def _create_minimal_row_after(self, table, ref_row, new_title: str, new_description: str):
        """Crée une nouvelle ligne minimale (sans dupliquer la structure de la référence)."""
        tbl = table._tbl
        ref_tr = ref_row._tr
        ref_cells = ref_row.cells
        texts = [new_title, new_description] if len(ref_cells) >= 2 else [f"{new_title}\n{new_description}"]

        new_tr = OxmlElement("w:tr")
        ref_trpr = ref_tr.find(qn("w:trPr"))
        if ref_trpr is not None:
            new_tr.append(deepcopy(ref_trpr))

        for i, text in enumerate(texts[: len(ref_cells)]):
            ref_cell = ref_cells[i]
            ref_tc = ref_cell._tc
            new_tc = OxmlElement("w:tc")
            ref_tcpr = ref_tc.find(qn("w:tcPr"))
            if ref_tcpr is not None:
                new_tc.append(deepcopy(ref_tcpr))

            new_p = OxmlElement("w:p")
            ref_para = ref_cell.paragraphs[0] if ref_cell.paragraphs else None
            ref_ppr = ref_para._element.find(qn("w:pPr")) if ref_para is not None else None
            if ref_ppr is not None:
                new_p.append(deepcopy(ref_ppr))

            new_r = OxmlElement("w:r")
            ref_rpr = None
            if ref_cell.paragraphs and ref_cell.paragraphs[0].runs:
                ref_rpr = ref_cell.paragraphs[0].runs[0]._element.find(qn("w:rPr"))
            if ref_rpr is not None:
                new_r.append(deepcopy(ref_rpr))

            new_t = OxmlElement("w:t")
            new_t.text = text
            new_r.append(new_t)
            new_p.append(new_r)
            new_tc.append(new_p)
            new_tr.append(new_tc)

            if self.markup_mode:
                rpr = new_r.find(qn("w:rPr"))
                if rpr is None:
                    rpr = OxmlElement("w:rPr")
                    new_r.insert(0, rpr)
                hl = OxmlElement("w:highlight")
                hl.set(qn("w:val"), "yellow")
                rpr.append(hl)

        ref_tr.addnext(new_tr)
        tr_list = list(tbl)
        ref_idx = tr_list.index(ref_tr)
        return table.rows[ref_idx + 1]

    def add_content(
        self,
        text: str,
        after_paragraph: Optional[str] = None,
        highlight: bool = False,
        red_highlight: bool = False,
    ):
        """Ajoute un paragraphe (titre/description) hors tableau."""
        if after_paragraph:
            target = self._find_body_paragraph(after_paragraph, occurrence=1)
            if target is None:
                raise ValueError(f"Paragraphe introuvable contenant : {after_paragraph!r}")
            new_p = self.doc.add_paragraph(text)
            new_p._element.getparent().remove(new_p._element)
            target._element.addnext(new_p._element)
            target_ppr = target._element.find(qn("w:pPr"))
            if target_ppr is not None:
                new_ppr = new_p._element.find(qn("w:pPr"))
                if new_ppr is not None:
                    new_p._element.remove(new_ppr)
                new_p._element.insert(0, deepcopy(target_ppr))
            if target.runs:
                self._copy_run_format(target.runs[-1], new_p.runs[0])
        else:
            new_p = self.doc.add_paragraph(text)
            last_ppr, last_rpr = self._get_last_paragraph_and_run_pr_before(new_p._element)
            if last_ppr is not None:
                new_ppr = new_p._element.find(qn("w:pPr"))
                if new_ppr is not None:
                    new_p._element.remove(new_ppr)
                new_p._element.insert(0, deepcopy(last_ppr))
            if last_rpr is not None and new_p.runs:
                dst_rpr = new_p.runs[0]._element.find(qn("w:rPr"))
                if dst_rpr is not None:
                    new_p.runs[0]._element.remove(dst_rpr)
                new_p.runs[0]._element.insert(0, deepcopy(last_rpr))
        for run in new_p.runs:
            if highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif red_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.RED
            elif not self.markup_mode:
                run.font.highlight_color = None
        return self

    def _get_last_paragraph_and_run_pr_before(self, exclude_element):
        """Retourne (pPr, rPr) du dernier bloc de contenu avant l'élément exclu."""
        body = self.doc.element.body
        last_ppr, last_rpr = None, None
        for block in body:
            if block is exclude_element:
                break
            if block.tag == qn("w:p"):
                last_ppr = block.find(qn("w:pPr"))
                runs = block.findall(qn("w:r"))
                if runs:
                    rpr = runs[-1].find(qn("w:rPr"))
                    if rpr is not None:
                        last_rpr = rpr
            elif block.tag == qn("w:tbl"):
                ppr, rpr = self._get_last_ppr_rpr_in_tbl(block)
                if ppr is not None:
                    last_ppr = ppr
                if rpr is not None:
                    last_rpr = rpr
        return last_ppr, last_rpr

    def _get_last_ppr_rpr_in_tbl(self, tbl_elem):
        """Dernier (pPr, rPr) dans un tableau (parcours récursif)."""
        last_ppr, last_rpr = None, None
        for tc in tbl_elem.findall(qn("w:tr")):
            for cell in tc.findall(qn("w:tc")):
                for p in cell.findall(qn("w:p")):
                    ppr = p.find(qn("w:pPr"))
                    if ppr is not None:
                        last_ppr = ppr
                    for r in p.findall(qn("w:r")):
                        rpr = r.find(qn("w:rPr"))
                        if rpr is not None:
                            last_rpr = rpr
                for nested_tbl in cell.findall(qn("w:tbl")):
                    ppr, rpr = self._get_last_ppr_rpr_in_tbl(nested_tbl)
                    if ppr is not None:
                        last_ppr = ppr
                    if rpr is not None:
                        last_rpr = rpr
        return last_ppr, last_rpr

    def remove_paragraph(self, text_contains: str, occurrence: int = 1):
        """Supprime un paragraphe hors tableau contenant le texte donné.
        Mode mark-up : barre le texte au lieu de le supprimer."""
        target = self._find_body_paragraph(text_contains, occurrence)
        if target is None:
            raise ValueError(f"Paragraphe introuvable contenant : {text_contains!r}")
        if self.markup_mode:
            for run in target.runs:
                run.font.strike = True
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            target._element.getparent().remove(target._element)
        return self

    def update_paragraph(self, text_contains: str, new_text: str, occurrence: int = 1):
        """Modifie le contenu d'un paragraphe hors tableau.
        Mode mark-up : surligne le nouveau texte en jaune."""
        target = self._find_body_paragraph(text_contains, occurrence)
        if target is None:
            raise ValueError(f"Paragraphe introuvable contenant : {text_contains!r}")
        target.clear()
        run = target.add_run(new_text)
        if self.markup_mode:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return self

    def _find_body_paragraph(self, text_contains: str, occurrence: int = 1):
        """Trouve le n-ième paragraphe du body (hors tableaux) contenant le texte."""
        count = 0
        for p in self.doc.paragraphs:
            if text_contains in p.text:
                count += 1
                if count == occurrence:
                    return p
        return None

    def add_logo_to_header(self, logo_path: str, width_inches: float = 1.0, all_sections: bool = True) -> bool:
        """Supprime le header existant et place uniquement le logo en haut à gauche."""
        try:
            sections = self.doc.sections if all_sections else [self.doc.sections[0]]
            for section in sections:
                header = section.header
                # Supprimer tout le contenu du header
                for paragraph in list(header.paragraphs):
                    paragraph._element.getparent().remove(paragraph._element)
                for table in list(header.tables):
                    table._element.getparent().remove(table._element)
                # Nouveau paragraphe avec uniquement le logo
                paragraph = header.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = paragraph.add_run()
                run.add_picture(logo_path, width=Inches(width_inches))
                paragraph.add_run("\n")
            return True
        except Exception as e:
            print(f"Erreur lors de l'ajout du logo : {e}")
            return False

    def save(self, output_path: str):
        self.doc.save(output_path)


# -----------------------------------------------------------------------------
# Transformer principal
# -----------------------------------------------------------------------------

class TermSheetTransformer:
    """
    Transforme un document Word term sheet et produit :
    - Document mark-up : modifications visibles (jaune + barré)
    - Document final : modifications appliquées
    - Document PDF : version PDF du final

    Exemple :
        t = TermSheetTransformer("term_sheet.docx")
        t.replace_word("Mot1", "Mot2")
        t.add_section_after("Listing", "Issuer", "Paul Berber")
        t.update_section_description("Country", "France")
        t.remove_section("Old Section")
        t.add_logo("logo.png")
        t.execute_and_export("./output", "modifications.docx", "final.docx", "final.pdf")
    """

    def __init__(self, docx_path: str):
        self._temp_dir: Optional[Path] = None
        path = Path(docx_path)
        if not path.exists():
            raise FileNotFoundError(f"Fichier introuvable : {docx_path}")
        if path.suffix.lower() == ".doc":
            self.docx_path, self._temp_dir = _convert_doc_to_docx(path)
        else:
            self.docx_path = path
        self.operations: List = []

    def __del__(self):
        if hasattr(self, "_temp_dir") and self._temp_dir is not None and self._temp_dir.exists():
            try:
                shutil.rmtree(self._temp_dir, ignore_errors=True)
            except Exception:
                pass

    def replace_word(self, old: str, new: str, occurrence: int = 1) -> "TermSheetTransformer":
        self.operations.append(ReplaceOp(old=old, new=new, occurrence=occurrence))
        return self

    def add_section_after(self, after_section: str, title: str, description: str,
                         occurrence: int = 1) -> "TermSheetTransformer":
        self.operations.append(AddSectionOp(
            after_section=after_section, title=title, description=description, occurrence=occurrence
        ))
        return self

    def update_section_description(self, section_title: str, new_description: str,
                                   occurrence: int = 1) -> "TermSheetTransformer":
        self.operations.append(UpdateDescriptionOp(
            section_title=section_title, new_description=new_description, occurrence=occurrence
        ))
        return self

    def remove_section(self, section_title: str, occurrence: int = 1) -> "TermSheetTransformer":
        self.operations.append(DeleteSectionOp(section_title=section_title, occurrence=occurrence))
        return self

    def add_logo(self, logo_path: str, width_inches: float = 1.0,
                 all_sections: bool = True) -> "TermSheetTransformer":
        self.operations.append(AddLogoOp(
            logo_path=logo_path, width_inches=width_inches, all_sections=all_sections
        ))
        return self

    def add_content(self, text: str, after_paragraph: Optional[str] = None,
                    red_highlight_in_final: bool = False) -> "TermSheetTransformer":
        """Ajoute un titre ou une description (paragraphe hors tableau)."""
        self.operations.append(AddContentOp(
            text=text, after_paragraph=after_paragraph, red_highlight_in_final=red_highlight_in_final
        ))
        return self

    def remove_paragraph(self, text_contains: str, occurrence: int = 1) -> "TermSheetTransformer":
        """Supprime un paragraphe contenant le texte donné."""
        self.operations.append(RemoveParagraphOp(text_contains=text_contains, occurrence=occurrence))
        return self

    def update_paragraph(self, text_contains: str, new_text: str,
                         occurrence: int = 1) -> "TermSheetTransformer":
        """Modifie le contenu d'un paragraphe."""
        self.operations.append(UpdateParagraphOp(
            text_contains=text_contains, new_text=new_text, occurrence=occurrence
        ))
        return self

    def execute_and_export(
        self,
        output_dir: str,
        markup_docx: str = "modifications.docx",
        final_docx: str = "document_final.docx",
        final_pdf: str = "document_final.pdf",
    ) -> None:
        """
        Exécute toutes les modifications et exporte les 3 fichiers.
        """
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        doc_markup = Document(str(self.docx_path))
        doc_final = Document(str(self.docx_path))

        editor_markup = _TermSheetEditor(doc_markup, markup_mode=True)
        editor_final = _TermSheetEditor(doc_final, markup_mode=False)

        for op in self.operations:
            if isinstance(op, ReplaceOp):
                editor_markup.replace_text(op.old, op.new)
                editor_final.replace_text(op.old, op.new)
            elif isinstance(op, AddSectionOp):
                editor_markup.insert_section_after(
                    op.after_section, op.title, op.description, op.occurrence
                )
                editor_final.insert_section_after(
                    op.after_section, op.title, op.description, op.occurrence
                )
            elif isinstance(op, UpdateDescriptionOp):
                editor_markup.update_section_description(
                    op.section_title, op.new_description, op.occurrence
                )
                editor_final.update_section_description(
                    op.section_title, op.new_description, op.occurrence
                )
            elif isinstance(op, DeleteSectionOp):
                editor_markup.delete_section(op.section_title, op.occurrence)
                editor_final.delete_section(op.section_title, op.occurrence)
            elif isinstance(op, AddLogoOp):
                editor_markup.add_logo_to_header(
                    op.logo_path, op.width_inches, op.all_sections
                )
                editor_final.add_logo_to_header(
                    op.logo_path, op.width_inches, op.all_sections
                )
            elif isinstance(op, AddContentOp):
                editor_markup.add_content(
                    op.text, op.after_paragraph,
                    highlight=True,  # jaune en mark-up
                    red_highlight=False,
                )
                editor_final.add_content(
                    op.text, op.after_paragraph,
                    highlight=False,
                    red_highlight=op.red_highlight_in_final,
                )
            elif isinstance(op, RemoveParagraphOp):
                editor_markup.remove_paragraph(op.text_contains, op.occurrence)
                editor_final.remove_paragraph(op.text_contains, op.occurrence)
            elif isinstance(op, UpdateParagraphOp):
                editor_markup.update_paragraph(op.text_contains, op.new_text, op.occurrence)
                editor_final.update_paragraph(op.text_contains, op.new_text, op.occurrence)

        path_markup = output_dir / markup_docx
        path_final = output_dir / final_docx
        path_pdf = output_dir / final_pdf

        editor_markup.save(str(path_markup))
        _strip_all_highlights(doc_final)
        editor_final.save(str(path_final))

        try:
            import docx2pdf
            docx2pdf.convert(str(path_final), str(path_pdf))
        except Exception as e:
            print(f"Erreur lors de la conversion PDF : {e}")
            print("Assurez-vous que Word est installé et que docx2pdf est correctement configuré.")

        print(f"Mark-up sauvegardé : {path_markup}")
        print(f"Document final sauvegardé : {path_final}")
        print(f"PDF sauvegardé : {path_pdf}")


# -----------------------------------------------------------------------------
# Exemple d'utilisation
# -----------------------------------------------------------------------------

if __name__ == "__main__":
    # Accepte .doc (conversion auto via LibreOffice) ou .docx
    transformer = TermSheetTransformer("mon_term_sheet.docx")

    transformer.replace_word("Mot1", "Mot2")
    transformer.add_section_after("Listing", "Issuer", "Paul Berber")
    transformer.update_section_description("Country", "France")
    transformer.remove_section("Old Section")
    transformer.add_content("Notice importante", red_highlight_in_final=True)
    transformer.add_content("Paragraphe après un autre", after_paragraph="Texte existant")
    transformer.update_paragraph("Ancien texte", "Nouveau texte")
    transformer.remove_paragraph("Paragraphe à supprimer")
    transformer.add_logo("logo_nbc.png", width_inches=0.8)

    transformer.execute_and_export(
        output_dir="./output",
        markup_docx="modifications.docx",
        final_docx="document_final.docx",
        final_pdf="document_final.pdf",
    )
