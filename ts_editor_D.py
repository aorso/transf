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
class SetSectionOp:
    """Ajoute ou met à jour une section (fusion add_section_after + update_section_description)."""
    title: str
    description: str
    after_section: Optional[str] = None  # Si None, ajoute à la fin du tableau
    red_highlight_in_final: bool = False
    occurrence: int = 1


@dataclass
class DeleteSectionOp:
    section_title: str
    occurrence: int = 1


@dataclass
class SetSectionOrderOp:
    """Définit l'ordre des sections dans le tableau."""
    section_order: List[str]


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


@dataclass
class SetDisclaimerSectionOp:
    """Ajoute ou met à jour une section dans la partie Disclaimer (après les tableaux)."""
    title: str
    content: str  # Peut contenir \n pour plusieurs paragraphes
    after_title: Optional[str] = None  # Titre après lequel insérer (None = fin)
    red_highlight_in_final: bool = False
    occurrence: int = 1


@dataclass
class RemoveDisclaimerSectionOp:
    """Supprime une section de disclaimer complète (titre + contenu)."""
    title: str
    occurrence: int = 1


@dataclass
class UpdateDisclaimerContentOp:
    """Met à jour le contenu d'une section de disclaimer existante."""
    title: str
    new_content: str  # Peut contenir \n pour plusieurs paragraphes
    red_highlight_in_final: bool = False
    occurrence: int = 1


@dataclass
class AddDisclaimerContentOp:
    """Ajoute du contenu à la fin d'une section de disclaimer existante."""
    title: str
    additional_content: str  # Peut contenir \n pour plusieurs paragraphes
    red_highlight_in_final: bool = False
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

    def set_section(
        self, 
        title: str, 
        description: str, 
        after_section: Optional[str] = None,
        red_highlight: bool = False,
        occurrence: int = 1
    ):
        """
        Ajoute ou met à jour une section.
        - Si la section existe : met à jour la description
        - Si la section n'existe pas : l'ajoute après after_section (ou à la fin si None)
        """
        row, table = self._find_section_row(title, occurrence)
        
        if row is not None:
            # Section existe : mise à jour de la description
            if len(row.cells) >= 2:
                self._set_cell_text(
                    row.cells[1], 
                    description, 
                    highlight=self.markup_mode,
                    red_highlight=red_highlight and not self.markup_mode
                )
            elif len(row.cells) == 1:
                norm_title = self._normalize(row.cells[0].text)
                self._set_cell_text(
                    row.cells[0], 
                    f"{norm_title}: {description}",
                    highlight=self.markup_mode,
                    red_highlight=red_highlight and not self.markup_mode
                )
        else:
            # Section n'existe pas : ajout
            if after_section:
                # Ajouter après une section spécifique
                ref_row, table = self._find_section_row(after_section, occurrence)
                if ref_row is None:
                    raise ValueError(f"Section de référence introuvable : {after_section}")
                self._create_minimal_row_after(table, ref_row, title, description, red_highlight)
            else:
                # Ajouter à la fin du premier tableau
                if not self.doc.tables:
                    raise ValueError("Aucun tableau trouvé dans le document")
                table = self.doc.tables[0]
                if not table.rows:
                    raise ValueError("Le tableau est vide")
                # Utiliser la dernière ligne comme référence
                last_row = table.rows[-1]
                self._create_minimal_row_after(table, last_row, title, description, red_highlight)
        return self

    def set_section_order(self, section_order: List[str]):
        """
        Réorganise les lignes du tableau pour respecter l'ordre des sections spécifié.
        Les sections non listées sont placées à la fin.
        """
        if not self.doc.tables:
            return self
        
        table = self.doc.tables[0]
        rows_data = []
        
        # Récupérer toutes les lignes avec leur titre normalisé
        for row in table.rows:
            if len(row.cells) >= 1:
                title = self._normalize(row.cells[0].text)
                rows_data.append((title, row))
        
        # Trier selon l'ordre spécifié
        def get_order_index(item):
            title = item[0]
            try:
                return section_order.index(title)
            except ValueError:
                return len(section_order)  # Placer à la fin si non spécifié
        
        sorted_rows = sorted(rows_data, key=get_order_index)
        
        # Réorganiser les lignes dans le tableau
        tbl = table._tbl
        for title, row in sorted_rows:
            tbl.append(row._tr)
        
        return self

    def _set_cell_text(self, cell, text: str, highlight: bool = False, red_highlight: bool = False):
        """
        Définit le texte d'une cellule en gérant les retours à la ligne (\n).
        """
        # Gérer les retours à la ligne
        lines = text.split('\n')
        
        if not cell.paragraphs:
            cell.add_paragraph("")
        
        # Supprimer tous les paragraphes sauf le premier
        for extra_p in cell.paragraphs[1:]:
            extra_p._element.getparent().remove(extra_p._element)
        
        # Premier paragraphe avec la première ligne
        p = cell.paragraphs[0]
        if p.runs:
            ref_run = p.runs[0]
            ref_run.text = lines[0] if lines else ""
            if highlight:
                ref_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif red_highlight:
                ref_run.font.highlight_color = WD_COLOR_INDEX.RED
            elif not self.markup_mode:
                ref_run.font.highlight_color = None
            # Supprimer les autres runs
            for r in p.runs[1:]:
                r.text = ""
        else:
            r = p.add_run(lines[0] if lines else "")
            if highlight:
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif red_highlight:
                r.font.highlight_color = WD_COLOR_INDEX.RED
            elif not self.markup_mode:
                r.font.highlight_color = None
        
        # Ajouter les lignes suivantes comme nouveaux paragraphes
        for line in lines[1:]:
            new_p = cell.add_paragraph(line)
            if new_p.runs:
                if highlight:
                    new_p.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
                elif red_highlight:
                    new_p.runs[0].font.highlight_color = WD_COLOR_INDEX.RED
                elif not self.markup_mode:
                    new_p.runs[0].font.highlight_color = None

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

    def _create_minimal_row_after(self, table, ref_row, new_title: str, new_description: str, red_highlight: bool = False):
        """Clone la ligne (format exact), modifie uniquement le texte (w:t) sans toucher pPr/rPr."""
        from docx.oxml.ns import nsmap
        w_ns = nsmap['w']
        
        tbl = table._tbl
        ref_tr = ref_row._tr
        new_tr = deepcopy(ref_tr)
        ref_tr.addnext(new_tr)

        texts = [new_title, new_description] if len(ref_row.cells) >= 2 else [f"{new_title}\n{new_description}"]
        tcs = new_tr.findall(qn("w:tc"))
        
        # Vérification robuste pour éviter index out of range
        for i, text in enumerate(texts):
            if i >= len(tcs):  # Sortir si plus de cellules disponibles
                break
            
            tc = tcs[i]
            t_elements = tc.findall(f".//{{{w_ns}}}t")
            
            # Vérifier que t_elements n'est pas vide
            if t_elements and len(t_elements) > 0:
                t_elements[0].text = text
                for extra_t in t_elements[1:]:
                    extra_t.text = ""
            else:
                # Si pas de t_elements, créer un run et un text element
                # Trouver ou créer un paragraph dans la cellule
                p_elem = tc.find(qn("w:p"))
                if p_elem is not None:
                    r_elem = OxmlElement("w:r")
                    t_elem = OxmlElement("w:t")
                    t_elem.text = text
                    r_elem.append(t_elem)
                    p_elem.append(r_elem)
            
            # Appliquer le surlignage approprié
            highlight_color = None
            if self.markup_mode:
                highlight_color = "yellow"
            elif red_highlight and i == 1:  # Surlignage rouge seulement sur la description (colonne 2)
                highlight_color = "red"
            
            if highlight_color:
                for r_elem in tc.findall(f".//{{{w_ns}}}r"):
                    rpr = r_elem.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr")
                        r_elem.insert(0, rpr)
                    hl = rpr.find(qn("w:highlight"))
                    if hl is None:
                        hl = OxmlElement("w:highlight")
                        rpr.append(hl)
                    hl.set(qn("w:val"), highlight_color)

        # Pas besoin de retourner la nouvelle ligne
        # car elle est déjà ajoutée au tableau
        return self

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
            if target.runs and new_p.runs:
                self._copy_run_format(target.runs[-1], new_p.runs[0])
        else:
            # Créer un paragraphe temporaire pour trouver la position
            temp_p = self.doc.add_paragraph()
            last_run, last_para = self._get_last_run_and_paragraph_before(temp_p._element)
            
            # Supprimer le paragraphe temporaire
            temp_p._element.getparent().remove(temp_p._element)
            
            # Cloner le dernier paragraphe (structure complète + style) puis modifier le texte
            if last_para is not None:
                new_p_elem = deepcopy(last_para._element)
                self.doc.element.body.append(new_p_elem)
                
                # Récupérer l'objet Paragraph python-docx
                new_p = None
                for p in self.doc.paragraphs:
                    if p._element is new_p_elem:
                        new_p = p
                        break
                
                # Modifier le texte du premier run (garde le rPr avec formatage)
                # et supprimer les runs suivants
                if new_p.runs:
                    new_p.runs[0].text = text
                    for run in list(new_p.runs[1:]):
                        run._element.getparent().remove(run._element)
                else:
                    # Si pas de run, en créer un (fallback)
                    new_p.add_run(text)
            else:
                # Fallback : créer un paragraphe normal si aucun paragraphe précédent
                new_p = self.doc.add_paragraph(text)
        for run in new_p.runs:
            if highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            elif red_highlight:
                run.font.highlight_color = WD_COLOR_INDEX.RED
            elif not self.markup_mode:
                run.font.highlight_color = None
        return self

    def _get_last_run_and_paragraph_before(self, exclude_element):
        """Retourne (dernier Run, dernier Paragraph) avant l'élément exclu (objets python-docx)."""
        body = self.doc.element.body
        last_run, last_para = None, None
        try:
            idx = list(body).index(exclude_element)
        except ValueError:
            idx = len(body)
        blocks_before = list(body)[:idx]

        for block in reversed(blocks_before):
            if block.tag == qn("w:p"):
                for p in self.doc.paragraphs:
                    if p._element is block:
                        last_para = p
                        last_run = p.runs[-1] if p.runs else None
                        return last_run, last_para
            elif block.tag == qn("w:tbl"):
                for table in self.doc.tables:
                    if table._tbl is block:
                        run, para = self._get_last_run_para_in_table(table)
                        return (run, para) if (run or para) else (last_run, last_para)
        return last_run, last_para

    def _get_last_run_para_in_table(self, table):
        """Dernier (Run, Paragraph) dans un tableau (objets python-docx)."""
        last_run, last_para = None, None
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    last_para = p
                    last_run = p.runs[-1] if p.runs else None
                for nested in cell.tables:
                    r, p = self._get_last_run_para_in_table(nested)
                    if r is not None:
                        last_run = r
                    if p is not None:
                        last_para = p
        return last_run, last_para

    def _copy_run_format_from_api(self, src_run, dst_run):
        """Copie le format via l'API python-docx (gère l'héritage des styles)."""
        try:
            dst_run.bold = src_run.bold
            dst_run.italic = src_run.italic
            dst_run.underline = src_run.underline
            if src_run.font.name:
                dst_run.font.name = src_run.font.name
            if src_run.font.size:
                dst_run.font.size = src_run.font.size
            if src_run.font.color and src_run.font.color.rgb:
                dst_run.font.color.rgb = src_run.font.color.rgb
        except (AttributeError, TypeError):
            pass

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

    # -------------------------------------------------------------------------
    # Gestion des sections de Disclaimer (paragraphes après les tableaux)
    # -------------------------------------------------------------------------

    def _is_disclaimer_title(self, paragraph) -> bool:
        """
        Détecte si un paragraphe est un titre de disclaimer.
        Critère : tous les runs en gras ET aucun souligné.
        """
        if not paragraph.text.strip():
            return False
        
        non_empty_runs = [r for r in paragraph.runs if r.text.strip()]
        if not non_empty_runs:
            return False
        
        all_bold = all(r.bold for r in non_empty_runs)
        no_underline = not any(r.underline for r in non_empty_runs)
        
        return all_bold and no_underline

    def _get_last_table_index(self):
        """Retourne l'index du dernier tableau dans le body."""
        last_table_idx = -1
        for i, elem in enumerate(self.doc.element.body):
            if elem.tag.split('}')[-1] == 'tbl':
                last_table_idx = i
        return last_table_idx

    def _get_disclaimer_paragraphs(self):
        """Retourne tous les paragraphes après le dernier tableau."""
        last_table_idx = self._get_last_table_index()
        
        disclaimer_paras = []
        for i, elem in enumerate(self.doc.element.body):
            if i > last_table_idx and elem.tag.split('}')[-1] == 'p':
                for p in self.doc.paragraphs:
                    if p._element == elem:
                        disclaimer_paras.append(p)
                        break
        return disclaimer_paras

    def _find_disclaimer_section(self, title: str, occurrence: int = 1):
        """
        Trouve une section de disclaimer par son titre.
        Retourne: (titre_paragraph, content_paragraphs[], last_content_paragraph)
        """
        disclaimer_paras = self._get_disclaimer_paragraphs()
        target = self._normalize(title)
        
        count = 0
        for i, para in enumerate(disclaimer_paras):
            if self._is_disclaimer_title(para) and self._normalize(para.text) == target:
                count += 1
                if count == occurrence:
                    # Récupérer le contenu (paragraphes suivants jusqu'au prochain titre)
                    content = []
                    last_content = None
                    for j in range(i + 1, len(disclaimer_paras)):
                        if self._is_disclaimer_title(disclaimer_paras[j]):
                            break
                        if disclaimer_paras[j].text.strip():
                            content.append(disclaimer_paras[j])
                            last_content = disclaimer_paras[j]
                    return para, content, last_content
        
        return None, [], None

    def _get_last_disclaimer_paragraph(self):
        """Retourne le dernier paragraphe de la zone disclaimer."""
        disclaimer_paras = self._get_disclaimer_paragraphs()
        return disclaimer_paras[-1] if disclaimer_paras else None

    def set_disclaimer_section(
        self,
        title: str,
        content: str,
        after_title: Optional[str] = None,
        red_highlight: bool = False,
        occurrence: int = 1
    ):
        """
        Ajoute ou met à jour une section dans la partie Disclaimer.
        - Si la section existe : remplace le contenu
        - Si la section n'existe pas : l'ajoute après after_title (ou à la fin)
        
        Le content peut contenir des \n pour séparer les paragraphes.
        """
        title_para, content_paras, _ = self._find_disclaimer_section(title, occurrence)
        
        if title_para is not None:
            # Section existe : remplacer le contenu
            if self.markup_mode:
                # Barrer l'ancien contenu
                for cp in content_paras:
                    for r in cp.runs:
                        r.font.strike = True
                        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            else:
                # Supprimer l'ancien contenu
                for cp in content_paras:
                    cp._element.getparent().remove(cp._element)
            
            # Ajouter le nouveau contenu après le titre
            content_lines = content.split('\n')
            last_elem = title_para._element
            ref_run = content_paras[0].runs[0] if content_paras and content_paras[0].runs else None
            
            for line in content_lines:
                if line.strip():
                    new_p = self.doc.add_paragraph()
                    new_p._element.getparent().remove(new_p._element)
                    last_elem.addnext(new_p._element)
                    last_elem = new_p._element
                    
                    # Copier le format du premier paragraphe de contenu
                    if content_paras and content_paras[0]._element.find(qn("w:pPr")) is not None:
                        ppr = content_paras[0]._element.find(qn("w:pPr"))
                        new_ppr = new_p._element.find(qn("w:pPr"))
                        if new_ppr is not None:
                            new_p._element.remove(new_ppr)
                        new_p._element.insert(0, deepcopy(ppr))
                    
                    # Ajouter le texte
                    run = new_p.add_run(line)
                    if ref_run:
                        self._copy_run_format(ref_run, run)
                    
                    # Appliquer highlighting
                    if self.markup_mode:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    elif red_highlight:
                        run.font.highlight_color = WD_COLOR_INDEX.RED
        else:
            # Section n'existe pas : créer titre + contenu
            if after_title:
                ref_title_para, ref_content, last_content = self._find_disclaimer_section(after_title, 1)
                if ref_title_para is None:
                    raise ValueError(f"Titre de référence introuvable : {after_title}")
                # Insérer après le dernier paragraphe de contenu de la section de référence
                insert_after = last_content if last_content else ref_title_para
            else:
                # Insérer à la fin
                insert_after = self._get_last_disclaimer_paragraph()
                if insert_after is None:
                    # Pas de disclaimer existant, ajouter à la fin du document
                    all_paras = list(self.doc.paragraphs)
                    if all_paras:
                        insert_after = all_paras[-1]
                    else:
                        insert_after = None
            
            if insert_after is None:
                raise ValueError("Impossible de trouver où insérer la section - le document ne contient aucun paragraphe")
            
            # Créer le titre
            title_para = self.doc.add_paragraph()
            title_para._element.getparent().remove(title_para._element)
            insert_after._element.addnext(title_para._element)
            
            # Mettre le titre en gras
            title_run = title_para.add_run(title)
            title_run.font.bold = True
            if self.markup_mode:
                title_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            
            # Créer le contenu
            content_lines = content.split('\n')
            last_elem = title_para._element
            for line in content_lines:
                if line.strip():
                    new_p = self.doc.add_paragraph()
                    new_p._element.getparent().remove(new_p._element)
                    last_elem.addnext(new_p._element)
                    last_elem = new_p._element
                    
                    run = new_p.add_run(line)
                    # Highlighting
                    if self.markup_mode:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    elif red_highlight:
                        run.font.highlight_color = WD_COLOR_INDEX.RED
        
        return self

    def remove_disclaimer_section(self, title: str, occurrence: int = 1):
        """
        Supprime une section de disclaimer complète (titre + contenu).
        En mode mark-up : barre tout au lieu de supprimer.
        """
        title_para, content_paras, _ = self._find_disclaimer_section(title, occurrence)
        if title_para is None:
            raise ValueError(f"Section de disclaimer introuvable : {title}")
        
        if self.markup_mode:
            # Barrer le titre
            for r in title_para.runs:
                r.font.strike = True
                r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            # Barrer le contenu
            for cp in content_paras:
                for r in cp.runs:
                    r.font.strike = True
                    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            # Supprimer le titre
            title_para._element.getparent().remove(title_para._element)
            # Supprimer le contenu
            for cp in content_paras:
                cp._element.getparent().remove(cp._element)
        
        return self

    def update_disclaimer_content(
        self,
        title: str,
        new_content: str,
        red_highlight: bool = False,
        occurrence: int = 1
    ):
        """
        Met à jour uniquement le contenu d'une section de disclaimer existante.
        Équivalent à set_disclaimer_section mais force que la section existe.
        """
        title_para, content_paras, _ = self._find_disclaimer_section(title, occurrence)
        if title_para is None:
            raise ValueError(f"Section de disclaimer introuvable : {title}")
        
        # Utiliser set_disclaimer_section qui gère déjà la mise à jour
        return self.set_disclaimer_section(title, new_content, red_highlight=red_highlight, occurrence=occurrence)

    def add_disclaimer_content(
        self,
        title: str,
        additional_content: str,
        red_highlight: bool = False,
        occurrence: int = 1
    ):
        """
        Ajoute du contenu à la fin d'une section de disclaimer existante.
        """
        title_para, content_paras, last_content = self._find_disclaimer_section(title, occurrence)
        if title_para is None:
            raise ValueError(f"Section de disclaimer introuvable : {title}")
        
        # Trouver où insérer
        insert_after = last_content if last_content else title_para
        ref_run = content_paras[0].runs[0] if content_paras and content_paras[0].runs else None
        
        # Ajouter le nouveau contenu
        content_lines = additional_content.split('\n')
        last_elem = insert_after._element
        
        for line in content_lines:
            if line.strip():
                new_p = self.doc.add_paragraph()
                new_p._element.getparent().remove(new_p._element)
                last_elem.addnext(new_p._element)
                last_elem = new_p._element
                
                # Copier le format
                if content_paras and content_paras[0]._element.find(qn("w:pPr")) is not None:
                    ppr = content_paras[0]._element.find(qn("w:pPr"))
                    new_ppr = new_p._element.find(qn("w:pPr"))
                    if new_ppr is not None:
                        new_p._element.remove(new_ppr)
                    new_p._element.insert(0, deepcopy(ppr))
                
                run = new_p.add_run(line)
                if ref_run:
                    self._copy_run_format(ref_run, run)
                
                # Highlighting
                if self.markup_mode:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                elif red_highlight:
                    run.font.highlight_color = WD_COLOR_INDEX.RED
        
        return self

    def add_logo_to_header(self, logo_path: str, width_inches: float = 1.0, all_sections: bool = True) -> bool:
        """Supprime le header existant et place uniquement le logo en haut à gauche."""
        try:
            if not self.doc.sections:
                raise ValueError("Le document ne contient aucune section")
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

    Fonctionnalités :
    1. Sections de tableau : Modifier les lignes du tableau principal
    2. Paragraphes non structurés : Modifier des paragraphes individuels
    3. Sections de Disclaimer : Modifier les sections titre/contenu après les tableaux
    4. Logo : Ajouter un logo dans le header

    Exemple :
        t = TermSheetTransformer("term_sheet.docx")
        
        # Remplacements de mots
        t.replace_word("Mot1", "Mot2")
        
        # Sections de tableau (méthode unifiée)
        t.set_section("Issuer", "Paul Berber", after_section="Listing")
        t.set_section("New Section", "Description")  # Ajoute à la fin si pas de référence
        t.set_section("Important", "Texte important", red_highlight_in_final=True)
        t.set_section_order(["Issuer", "Country", "Currency"])
        t.remove_section("Old Section")
        
        # Paragraphes non structurés
        t.add_content("Notice importante", red_highlight_in_final=True)
        t.update_paragraph("Ancien texte", "Nouveau texte")
        t.remove_paragraph("Paragraphe à supprimer")
        
        # Sections de Disclaimer (titre/contenu après tableaux)
        t.set_disclaimer_section("Important note", "Nouveau texte\\nLigne 2")
        t.update_disclaimer_content("Wichtiger Hinweis", "Contenu mis à jour")
        t.add_disclaimer_content("Important Risks", "Paragraphe supplémentaire")
        t.remove_disclaimer_section("Wesentliche Risiken")
        
        # Logo
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

    def set_section(
        self, 
        title: str, 
        description: str,
        after_section: Optional[str] = None,
        red_highlight_in_final: bool = False,
        occurrence: int = 1
    ) -> "TermSheetTransformer":
        """
        Ajoute ou met à jour une section dans le tableau.
        - Si la section existe : met à jour la description
        - Si la section n'existe pas : l'ajoute après after_section (ou à la fin si None)
        - red_highlight_in_final : surligne en rouge dans la version finale (pas dans le markup)
        """
        self.operations.append(SetSectionOp(
            title=title,
            description=description,
            after_section=after_section,
            red_highlight_in_final=red_highlight_in_final,
            occurrence=occurrence
        ))
        return self

    def set_section_order(self, section_order: List[str]) -> "TermSheetTransformer":
        """
        Définit l'ordre des sections dans le tableau.
        Les sections non listées sont placées à la fin.
        """
        self.operations.append(SetSectionOrderOp(section_order=section_order))
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

    # -------------------------------------------------------------------------
    # Gestion des sections de Disclaimer (paragraphes structurés après les tableaux)
    # -------------------------------------------------------------------------

    def set_disclaimer_section(
        self,
        title: str,
        content: str,
        after_title: Optional[str] = None,
        red_highlight_in_final: bool = False,
        occurrence: int = 1
    ) -> "TermSheetTransformer":
        """
        Ajoute ou met à jour une section dans la partie Disclaimer (après les tableaux).
        
        - Si la section existe : remplace son contenu
        - Si la section n'existe pas : l'ajoute après after_title (ou à la fin si None)
        - Le content peut contenir des \\n pour séparer les paragraphes
        - red_highlight_in_final : surligne en rouge dans la version finale
        
        Exemple:
            t.set_disclaimer_section("Important note", "Texte ligne 1\\nTexte ligne 2")
            t.set_disclaimer_section("New Warning", "Attention!", after_title="Important note")
        """
        self.operations.append(SetDisclaimerSectionOp(
            title=title,
            content=content,
            after_title=after_title,
            red_highlight_in_final=red_highlight_in_final,
            occurrence=occurrence
        ))
        return self

    def remove_disclaimer_section(self, title: str, occurrence: int = 1) -> "TermSheetTransformer":
        """
        Supprime une section de disclaimer complète (titre + contenu).
        
        Exemple:
            t.remove_disclaimer_section("Important note")
        """
        self.operations.append(RemoveDisclaimerSectionOp(title=title, occurrence=occurrence))
        return self

    def update_disclaimer_content(
        self,
        title: str,
        new_content: str,
        red_highlight_in_final: bool = False,
        occurrence: int = 1
    ) -> "TermSheetTransformer":
        """
        Met à jour uniquement le contenu d'une section de disclaimer existante.
        La section doit exister, sinon une erreur sera levée.
        
        Équivalent à set_disclaimer_section mais force que la section existe.
        
        Exemple:
            t.update_disclaimer_content("Important note", "Nouveau contenu\\nDeuxième ligne")
        """
        self.operations.append(UpdateDisclaimerContentOp(
            title=title,
            new_content=new_content,
            red_highlight_in_final=red_highlight_in_final,
            occurrence=occurrence
        ))
        return self

    def add_disclaimer_content(
        self,
        title: str,
        additional_content: str,
        red_highlight_in_final: bool = False,
        occurrence: int = 1
    ) -> "TermSheetTransformer":
        """
        Ajoute du contenu à la fin d'une section de disclaimer existante.
        La section doit exister, sinon une erreur sera levée.
        
        Exemple:
            t.add_disclaimer_content("Important note", "Paragraphe supplémentaire")
        """
        self.operations.append(AddDisclaimerContentOp(
            title=title,
            additional_content=additional_content,
            red_highlight_in_final=red_highlight_in_final,
            occurrence=occurrence
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
            elif isinstance(op, SetSectionOp):
                editor_markup.set_section(
                    op.title, op.description, op.after_section, 
                    red_highlight=False, occurrence=op.occurrence
                )
                editor_final.set_section(
                    op.title, op.description, op.after_section,
                    red_highlight=op.red_highlight_in_final, occurrence=op.occurrence
                )
            elif isinstance(op, SetSectionOrderOp):
                editor_markup.set_section_order(op.section_order)
                editor_final.set_section_order(op.section_order)
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
            elif isinstance(op, SetDisclaimerSectionOp):
                editor_markup.set_disclaimer_section(
                    op.title, op.content, op.after_title,
                    red_highlight=False, occurrence=op.occurrence
                )
                editor_final.set_disclaimer_section(
                    op.title, op.content, op.after_title,
                    red_highlight=op.red_highlight_in_final, occurrence=op.occurrence
                )
            elif isinstance(op, RemoveDisclaimerSectionOp):
                editor_markup.remove_disclaimer_section(op.title, op.occurrence)
                editor_final.remove_disclaimer_section(op.title, op.occurrence)
            elif isinstance(op, UpdateDisclaimerContentOp):
                editor_markup.update_disclaimer_content(
                    op.title, op.new_content,
                    red_highlight=False, occurrence=op.occurrence
                )
                editor_final.update_disclaimer_content(
                    op.title, op.new_content,
                    red_highlight=op.red_highlight_in_final, occurrence=op.occurrence
                )
            elif isinstance(op, AddDisclaimerContentOp):
                editor_markup.add_disclaimer_content(
                    op.title, op.additional_content,
                    red_highlight=False, occurrence=op.occurrence
                )
                editor_final.add_disclaimer_content(
                    op.title, op.additional_content,
                    red_highlight=op.red_highlight_in_final, occurrence=op.occurrence
                )

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

    # Remplacements de mots
    transformer.replace_word("Mot1", "Mot2")
    
    # Nouvelle méthode set_section : ajoute ou met à jour une section
    # Si "Issuer" n'existe pas, l'ajoute après "Listing"
    # Si "Issuer" existe, met à jour sa description
    transformer.set_section("Issuer", "Paul Berber", after_section="Listing")
    
    # Si after_section n'est pas spécifié, ajoute à la fin du tableau
    transformer.set_section("New Section", "Description de la nouvelle section")
    
    # Avec surlignage rouge dans la version finale
    transformer.set_section("Important Note", "Texte important", red_highlight_in_final=True)
    
    # Pour des retours à la ligne dans la description, utiliser \n
    transformer.set_section("Multi-line", "Ligne 1\nLigne 2\nLigne 3")
    
    # Anciennes méthodes toujours disponibles
    transformer.add_section_after("Listing", "Other Section", "Other Description")
    transformer.update_section_description("Country", "France")
    
    # Définir l'ordre des sections
    transformer.set_section_order(["Issuer", "Country", "Currency", "Amount"])
    
    # Suppression
    transformer.remove_section("Old Section")
    
    # Contenu hors tableau (paragraphes non structurés)
    transformer.add_content("Notice importante", red_highlight_in_final=True)
    transformer.add_content("Paragraphe après un autre", after_paragraph="Texte existant")
    transformer.update_paragraph("Ancien texte", "Nouveau texte")
    transformer.remove_paragraph("Paragraphe à supprimer")
    
    # -------------------------------------------------------------------------
    # Sections de Disclaimer (paragraphes structurés titre/contenu après tableaux)
    # -------------------------------------------------------------------------
    
    # Modifier une section de disclaimer existante
    transformer.set_disclaimer_section(
        "Important note",
        "Nouveau texte pour cette section\nDeuxième paragraphe"
    )
    
    # Ajouter une nouvelle section de disclaimer
    transformer.set_disclaimer_section(
        "New Warning",
        "Ceci est un nouveau disclaimer\nAvec plusieurs lignes",
        after_title="Important note",
        red_highlight_in_final=True
    )
    
    # Mettre à jour uniquement le contenu (la section doit exister)
    transformer.update_disclaimer_content(
        "Wichtiger Hinweis",
        "Contenu mis à jour"
    )
    
    # Ajouter du contenu à la fin d'une section existante
    transformer.add_disclaimer_content(
        "Important Risks",
        "Paragraphe supplémentaire ajouté"
    )
    
    # Supprimer une section de disclaimer complète
    transformer.remove_disclaimer_section("Wesentliche Risiken")
    
    # Logo
    transformer.add_logo("logo_nbc.png", width_inches=0.8)

    transformer.execute_and_export(
        output_dir="./output",
        markup_docx="modifications.docx",
        final_docx="document_final.docx",
        final_pdf="document_final.pdf",
    )
