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

from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.shared import Inches


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


# -----------------------------------------------------------------------------
# Éditeur de base (logique version C + logo version A)
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
        dst_run.bold = src_run.bold
        dst_run.italic = src_run.italic
        dst_run.underline = src_run.underline
        dst_run.font.all_caps = src_run.font.all_caps
        dst_run.font.small_caps = src_run.font.small_caps
        dst_run.font.strike = src_run.font.strike
        dst_run.font.double_strike = src_run.font.double_strike
        dst_run.font.subscript = src_run.font.subscript
        dst_run.font.superscript = src_run.font.superscript
        dst_run.font.shadow = src_run.font.shadow
        dst_run.font.outline = src_run.font.outline
        dst_run.font.rtl = src_run.font.rtl
        dst_run.font.imprint = src_run.font.imprint
        dst_run.font.emboss = src_run.font.emboss
        dst_run.font.hidden = src_run.font.hidden
        dst_run.font.name = src_run.font.name
        dst_run.font.size = src_run.font.size
        if src_run.font.color is not None and src_run.font.color.rgb is not None:
            dst_run.font.color.rgb = src_run.font.color.rgb
        if not self.markup_mode:
            dst_run.font.highlight_color = src_run.font.highlight_color

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
        new_row = self._clone_row_after(table, ref_row)
        if len(new_row.cells) >= 2:
            self._set_cell_text(new_row.cells[0], new_title, highlight=self.markup_mode)
            self._set_cell_text(new_row.cells[1], new_description, highlight=self.markup_mode)
        elif len(new_row.cells) == 1:
            self._set_cell_text(
                new_row.cells[0], f"{new_title}\n{new_description}", highlight=self.markup_mode
            )
        return self

    def _set_cell_text(self, cell, text: str, highlight: bool = False):
        if not cell.paragraphs:
            cell.add_paragraph("")
        p = cell.paragraphs[0]
        if p.runs:
            ref_run = p.runs[0]
            ref_run.text = text
            if highlight:
                ref_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            for r in p.runs[1:]:
                r.text = ""
            for extra_p in cell.paragraphs[1:]:
                extra_p._element.getparent().remove(extra_p._element)
        else:
            r = p.add_run(text)
            if highlight:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW

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

    def _clone_row_after(self, table, ref_row):
        tbl = table._tbl
        new_tr = deepcopy(ref_row._tr)
        ref_row._tr.addnext(new_tr)
        all_rows = list(table.rows)
        ref_idx = all_rows.index(ref_row)
        return table.rows[ref_idx + 1]

    def add_logo_to_header(self, logo_path: str, width_inches: float = 1.0, all_sections: bool = True) -> bool:
        try:
            sections = self.doc.sections if all_sections else [self.doc.sections[0]]
            for section in sections:
                header = section.header
                if not header.paragraphs:
                    paragraph = header.add_paragraph()
                else:
                    paragraph = header.paragraphs[0]
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
        self.docx_path = Path(docx_path)
        if not self.docx_path.exists():
            raise FileNotFoundError(f"Fichier introuvable : {docx_path}")
        self.operations: List = []

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

        path_markup = output_dir / markup_docx
        path_final = output_dir / final_docx
        path_pdf = output_dir / final_pdf

        editor_markup.save(str(path_markup))
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
    transformer = TermSheetTransformer("mon_term_sheet.docx")

    transformer.replace_word("Mot1", "Mot2")
    transformer.add_section_after("Listing", "Issuer", "Paul Berber")
    transformer.update_section_description("Country", "France")
    transformer.remove_section("Old Section")
    transformer.add_logo("logo_nbc.png", width_inches=0.8)

    transformer.execute_and_export(
        output_dir="./output",
        markup_docx="modifications.docx",
        final_docx="document_final.docx",
        final_pdf="document_final.pdf",
    )
