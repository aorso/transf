import pdfplumber
import pandas as pd
import numpy as np
from math import isfinite


class PdfTwoColumnExtractor:
    """
    Extract 2-column key/value tables from a PDF into a tidy DataFrame
    that follows the exact pipeline used in your notebook:
      1) Detect the vertical split (x_mid) on page 0 via modes (histogram peaks)
         or fallback to 1-D k-means on first-word x0s.
      2) Compute the list of row separators (Y coordinates) on each page.
      3) Reuse the same abscissas (left x, x_mid, right x) for all pages; only
         recompute Y lines per page.
      4) Slice characters in each cell box and concatenate them into strings.
      5) Pivot to two columns and clean.

    Public API
    ----------
    to_dataframe(pdf_path: str) -> pd.DataFrame
        Returns a DataFrame with columns ['titre', 'description'].

    Notes
    -----
    • PDF coordinate system (pdfplumber): origin at top-left, y increases downward.
    • We build row bands between consecutive horizontal Y separators (top to bottom).
    • Abscissas reused across pages ensure stable column boundaries.
    """

    def __init__(
        self,
        y_tol: float = 2.0,
        y_shift: float = -8.0,
        use_modes: bool = True,
        bin_size: float = 1.0,
        min_count: int = 4,
        delta: float = 3.0,
        extract_chars_for_modes: bool = False,
        left_margin: float = 12.0,
        right_margin: float = 12.0,
        min_row_gap: float = 12.0,
    ):
        self.y_tol = y_tol
        self.y_shift = y_shift
        self.use_modes = use_modes
        self.bin_size = bin_size
        self.min_count = min_count
        self.delta = delta
        self.extract_chars_for_modes = extract_chars_for_modes
        self.left_margin = left_margin
        self.right_margin = right_margin
        self.min_row_gap = min_row_gap

    # ----------------------------- PUBLIC API ----------------------------- #

    def to_dataframe(self, pdf_path: str) -> pd.DataFrame:
        """Process the whole PDF and return a 2-column DataFrame."""
        # Compute abscissas once on page 0
        coords0 = self._get_table_coordinates(pdf_path, page_no=0)
        abscisses = coords0['abscisses']  # [x_left, x_mid, x_right]

        all_rows = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_no in range(len(pdf.pages)):
                try:
                    df_chars = self._extract_chars_coordinates_page(pdf_path, page_no, y_mode="top")
                    coords = self._get_table_coordinates_with_fixed_abscisses(
                        pdf_path, page_no, abscisses
                    )
                    table_df = self._extract_table_cells(df_chars, coords)
                    if len(table_df) == 0:
                        continue
                    # pivot to two columns
                    table_pivot = table_df.pivot(index='ligne', columns='colonne', values='contenu')
                    for col in table_pivot.columns:
                        table_pivot[col] = table_pivot[col].str.replace('_', '', regex=False)
                    if table_pivot.shape[1] >= 2:
                        two_cols = table_pivot.iloc[:, :2].copy()
                        two_cols.columns = ['titre', 'description']
                        two_cols = two_cols[(two_cols['titre'].str.len() > 0) | (two_cols['description'].str.len() > 0)]
                        if len(two_cols) > 0:
                            all_rows.append(two_cols)
                except Exception as e:
                    print(f"Erreur sur la page {page_no + 1}: {e}")
                    continue

        if all_rows:
            return pd.concat(all_rows, ignore_index=True)
        return pd.DataFrame(columns=['titre', 'description'])

    # ------------------------- PAGE/CHAR EXTRACTION ----------------------- #

    def _extract_chars_coordinates_page(self, pdf_path: str, page_no: int, y_mode: str = "top") -> pd.DataFrame:
        """Return DataFrame with columns ['caractere','x','y'] for one page."""
        rows = []
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_no]
            for ch in page.chars:
                text = ch.get("text", "")
                if text is None:
                    continue
                x = float(ch.get("x0", 0.0))
                if y_mode == "center":
                    y = float((ch.get("top", 0.0) + ch.get("bottom", 0.0)) / 2.0)
                elif y_mode == "bottom":
                    y = float(ch.get("bottom", 0.0))
                else:
                    y = float(ch.get("top", 0.0))
                rows.append((text, x, y))
        return pd.DataFrame(rows, columns=["caractere", "x", "y"])

    # -------------------------- TABLE COORDINATES ------------------------- #

    def _get_table_coordinates(self, pdf_path: str, page_no: int = 0) -> dict:
        """Compute abscissas and Y lines on a page."""
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_no]
            page_width, page_height = page.width, page.height

            words = page.extract_words(
                use_text_flow=True,
                keep_blank_chars=False,
                y_tolerance=self.y_tol,
                x_tolerance=1.0,
            )

            # x_mid detection
            if self.use_modes:
                x_mid = self._compute_x_mid_by_modes(page, words)
            else:
                # use first-word x0 per line
                lines = self._group_lines(words, y_tol=self.y_tol)
                firsts = [self._first_word(line)["x0"] for line in lines]
                c1, c2 = self._kmeans_1d_two_centers(firsts)
                x_mid = (c1 + c2) / 2.0

            # gather left-column line starters
            lines = self._group_lines(words, y_tol=self.y_tol)
            firsts = [self._first_word(line) for line in lines]
            left = sorted([f for f in firsts if f["x0"] < x_mid], key=lambda d: d["y"])

            # keep only plausible starts (as in your notebook: filter by vertical gaps)
            df = pd.DataFrame({"col1_premier_mot": [f["text"] for f in left], "y": [f["y"] for f in left]})
            df["ecart_y"] = df["y"].diff()
            df_filtre = df[df['ecart_y'].isna() | (df['ecart_y'] > self.min_row_gap)].copy()

            # --- IMPORTANT: return ordonnees in CANVAS coordinates to match notebook ---
            # Extract plumber y-values, then convert to canvas coordinates applying y_shift
            y_vals = df_filtre["y"].dropna().astype(float).tolist()
            ordonnees_plumber = sorted({round(y, 3) for y in y_vals})
            ordonnees_canvas = [page_height - (float(y) + float(self.y_shift)) for y in ordonnees_plumber]
            ordonnees_canvas = [y for y in ordonnees_canvas if 0 <= y <= page_height]

            ordonnees = ordonnees_canvas
            # Guard: ensure at least 2 lines to define at least one row band
            if len(ordonnees) < 2:
                # fallback: coarse canvas bands (top -> bottom)
                ordonnees = [self.left_margin, page_height / 2.0, page_height - self.right_margin]

            x_left = self.left_margin
            x_right = page_width - self.right_margin

            return {
                'abscisses': [x_left, x_mid, x_right],
                'ordonnees': ordonnees,
                'dimensions': {'width': page_width, 'height': page_height},
                'metadata': {'page_no': page_no}
            }

    def _get_table_coordinates_with_fixed_abscisses(
        self,
        pdf_path: str,
        page_no: int,
        abscisses_fixes: list,
    ) -> dict:
        """Reuse x_left, x_mid, x_right but recompute Y lines for this page."""
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_no]
            page_width, page_height = page.width, page.height
            words = page.extract_words(
                use_text_flow=True,
                keep_blank_chars=False,
                y_tolerance=self.y_tol,
                x_tolerance=1.0,
            )
            x_mid = abscisses_fixes[1]
            lines = self._group_lines(words, y_tol=self.y_tol)
            firsts = [self._first_word(line) for line in lines]
            left = sorted([f for f in firsts if f["x0"] < x_mid], key=lambda d: d["y"])

            df = pd.DataFrame({"col1_premier_mot": [f["text"] for f in left], "y": [f["y"] for f in left]})
            df["ecart_y"] = df["y"].diff()
            df_filtre = df[df['ecart_y'].isna() | (df['ecart_y'] > self.min_row_gap)].copy()

            # Convert plumber y-values to canvas coordinates and apply y_shift
            y_vals = df_filtre["y"].dropna().astype(float).tolist()
            ordonnees_plumber = sorted({round(y, 3) for y in y_vals})
            ordonnees_canvas = [page_height - (float(y) + float(self.y_shift)) for y in ordonnees_plumber]
            ordonnees_canvas = [y for y in ordonnees_canvas if 0 <= y <= page_height]

            ordonnees = ordonnees_canvas

            x_left = abscisses_fixes[0]
            x_right = abscisses_fixes[2]
            return {
                'abscisses': [x_left, x_mid, x_right],
                'ordonnees': ordonnees,
                'dimensions': {'width': page_width, 'height': page_height},
                'metadata': {'page_no': page_no}
            }

    # ----------------------------- CELL EXTRACTION ------------------------ #

    def _extract_table_cells(self, df_chars: pd.DataFrame, coords: dict) -> pd.DataFrame:
        table_data = []
        page_height = coords['dimensions']['height']

        # Convert row lines to pdfplumber y (top-origin)
        ord_plumber = [page_height - y for y in coords['ordonnees']]
        ord_plumber.sort(reverse=True)  # top -> bottom

        abscisses = coords['abscisses']  # [x_left, x_mid, x_right]

        # iterate row bands
        for i in range(len(ord_plumber) - 1):
            top_y = ord_plumber[i]
            bot_y = ord_plumber[i + 1]
            # two columns: [left,x_mid] and [x_mid,right]
            for j in range(2):
                x1, x2 = abscisses[j], abscisses[j + 1]
                cell = self._slice_df_by_box(df_chars, x1, x2, top_y, bot_y)
                if len(cell) == 0:
                    s = ""
                else:
                    cell_sorted = cell.sort_values(by=["y", "x"]).copy()
                    s = "".join(cell_sorted["caractere"].tolist())
                    # simple cleanups
                    s = s.replace("\n", " ")
                    s = s.replace("__", "_")
                    s = s.strip()
                table_data.append({"ligne": i, "colonne": j, "contenu": s})

        return pd.DataFrame(table_data, columns=["ligne", "colonne", "contenu"]) if table_data else pd.DataFrame(columns=["ligne", "colonne", "contenu"]) 

    def _slice_df_by_box(self, df: pd.DataFrame, x1: float, x2: float, y1: float, y2: float) -> pd.DataFrame:
        xmin, xmax = min(x1, x2), max(x1, x2)
        ymin, ymax = min(y1, y2), max(y1, y2)
        mask = (df["x"] >= xmin) & (df["x"] < xmax) & (df["y"] >= ymin) & (df["y"] < ymax)
        return df.loc[mask, ["caractere", "x", "y"]]

    # ---------------------------- HELPERS (x_mid) ------------------------- #

    def _group_lines(self, words, y_tol: float):
        lines = []
        for w in words:
            inserted = False
            wy = (float(w["top"]) + float(w["bottom"])) / 2.0
            for line in lines:
                if abs(((min(v['top'] for v in line) + max(v['bottom'] for v in line)) / 2.0) - wy) <= y_tol:
                    line.append(w)
                    inserted = True
                    break
            if not inserted:
                lines.append([w])
        return lines

    def _first_word(self, line):
        w0 = min(line, key=lambda w: float(w['x0']))
        return {
            "text": w0['text'],
            "x0": float(w0['x0']),
            "y": (min(float(w['top']) for w in line) + max(float(w['bottom']) for w in line)) / 2.0,
        }

    def _kmeans_1d_two_centers(self, xs, max_iter: int = 50):
        xs = np.array([x for x in xs if isfinite(x)], dtype=float)
        if len(xs) < 4:
            raise ValueError("Pas assez de points de données pour le clustering")
        c1, c2 = np.percentile(xs, 25), np.percentile(xs, 75)
        for _ in range(max_iter):
            labels = (np.abs(xs - c2) < np.abs(xs - c1)).astype(int)
            new_c1 = xs[labels == 0].mean() if np.any(labels == 0) else c1
            new_c2 = xs[labels == 1].mean() if np.any(labels == 1) else c2
            if abs(new_c1 - c1) < 1e-6 and abs(new_c2 - c2) < 1e-6:
                break
            c1, c2 = new_c1, new_c2
        return tuple(sorted([float(c1), float(c2)]))

    def _compute_x_mid_by_modes(self, page, words) -> float:
        # get x0 list for each line's first word (or first char of that word)
        if self.extract_chars_for_modes:
            chars = page.chars
            lines = self._group_lines(words, y_tol=self.y_tol)
            x0_list = []
            for line in lines:
                w = min(line, key=lambda ww: float(ww["x0"]))
                wx0, wx1 = float(w["x0"]), float(w["x1"])
                wtop = float(w["top"])
                candidates = [c for c in chars if (float(c["x0"]) >= wx0 - 0.5 and float(c["x0"]) < wx1 + 0.5 and abs(float(c["top"]) - wtop) <= self.y_tol + 1.5)]
                if candidates:
                    first_char = min(candidates, key=lambda c: float(c["x0"]))
                    x0_list.append(float(first_char["x0"]))
                else:
                    x0_list.append(float(w["x0"]))
        else:
            lines = self._group_lines(words, y_tol=self.y_tol)
            x0_list = [float(min(line, key=lambda w: float(w["x0"]))["x0"]) for line in lines]

        # histogram bins
        xs = np.array(x0_list, dtype=float)
        if len(xs) == 0:
            # fallback: center of page
            return page.width / 2.0

        bins = np.floor(xs / self.bin_size) * self.bin_size
        uniq, counts = np.unique(bins, return_counts=True)
        # top two peaks
        top = sorted(zip(uniq, counts), key=lambda t: t[1], reverse=True)
        top = [b for (b, c) in top if c >= self.min_count][:2] or [page.width/3.0, 2*page.width/3.0]
        page_mid = page.width / 2.0
        centers_sorted_by_centrality = sorted(top, key=lambda x: abs(x - page_mid))
        x_col2 = centers_sorted_by_centrality[0]
        x_mid = float(x_col2) - float(self.delta)
        return x_mid


# ------------------------------- EXAMPLE -------------------------------- #
# if __name__ == "__main__":
#     extractor = PdfTwoColumnExtractor()
#     df = extractor.to_dataframe("/path/to/file.pdf")
#     print(df.head())
