"""
Modulo per la generazione di codice LaTeX dagli elementi del documento parsato.
Produce un file .tex completo e compilabile.
"""

import os
import re
from pathlib import Path
from typing import Optional

from .parser import (
    DocumentElement, ElementType, TextRun, TextStyle,
    Alignment, ListType, TableCell
)


class LaTeXGenerator:
    """Genera codice LaTeX a partire dagli elementi estratti dal documento."""

    def __init__(
        self,
        elements: list,
        metadata: dict = None,
        image_dir: str = None,
        document_class: str = 'article',
        font_size: int = 11,
        paper_size: str = 'a4paper',
        language: str = 'italian',
        use_microtype: bool = True,
        use_hyperref: bool = True,
    ):
        self.elements = elements
        self.metadata = metadata or {}
        self.image_dir = image_dir
        self.document_class = document_class
        self.font_size = font_size
        self.paper_size = paper_size
        self.language = language
        self.use_microtype = use_microtype
        self.use_hyperref = use_hyperref

        self._has_images = any(e.element_type == ElementType.IMAGE for e in elements)
        self._has_tables = any(e.element_type == ElementType.TABLE for e in elements)
        self._has_hyperlinks = any(
            any(r.hyperlink for r in e.runs) for e in elements if e.runs
        )
        self._has_colors = any(
            any(r.font_color for r in e.runs) for e in elements if e.runs
        )
        self._has_strikethrough = any(
            any(TextStyle.STRIKETHROUGH in r.styles for r in e.runs)
            for e in elements if e.runs
        )
        self._has_lists = any(e.element_type == ElementType.LIST_ITEM for e in elements)
        self._has_code_blocks = any(
            e.element_type == ElementType.CODE_BLOCK for e in elements
        )

    def generate(self) -> str:
        """Genera il codice LaTeX completo."""
        parts = []
        parts.append(self._generate_preamble())
        parts.append(self._generate_begin_document())
        parts.append(self._generate_title_block())
        parts.append(self._generate_body())
        parts.append(self._generate_end_document())

        latex = '\n'.join(parts)
        # Pulizia: rimuovi righe vuote excessive
        latex = re.sub(r'\n{4,}', '\n\n\n', latex)
        return latex

    def _generate_preamble(self) -> str:
        """Genera il preambolo LaTeX con tutti i pacchetti necessari."""
        lines = []

        # Document class
        lines.append(
            f'\\documentclass[{self.font_size}pt, {self.paper_size}]'
            f'{{{self.document_class}}}'
        )
        lines.append('')

        # Encoding e font
        lines.append('% === Encoding e Font ===')
        lines.append('\\usepackage[utf8]{inputenc}')
        lines.append('\\usepackage[T1]{fontenc}')
        lines.append('\\usepackage{lmodern}')  # Latin Modern fonts
        lines.append('')

        # Lingua
        lines.append('% === Lingua ===')
        lines.append(f'\\usepackage[{self.language}]{{babel}}')
        lines.append('')

        # Geometria pagina
        lines.append('% === Layout Pagina ===')
        lines.append('\\usepackage{geometry}')
        lines.append('\\geometry{')
        lines.append('  top=2.5cm,')
        lines.append('  bottom=2.5cm,')
        lines.append('  left=2.5cm,')
        lines.append('  right=2.5cm,')
        lines.append('}')
        lines.append('')

        # Microtypography
        if self.use_microtype:
            lines.append('% === Microtipografia ===')
            lines.append('\\usepackage{microtype}')
            lines.append('')

        # Float (necessario per l'opzione [H] su figure e tabelle)
        if self._has_images or self._has_tables:
            lines.append('% === Float ===')
            lines.append('\\usepackage{float}')
            lines.append('')

        # Immagini
        if self._has_images:
            lines.append('% === Immagini ===')
            lines.append('\\usepackage{graphicx}')
            if self.image_dir:
                # Percorso relativo per graphicspath
                lines.append(f'\\graphicspath{{{{{self.image_dir}/}}}}')
            lines.append('')

        # Tabelle
        if self._has_tables:
            lines.append('% === Tabelle ===')
            lines.append('\\usepackage{array}')
            lines.append('\\usepackage{booktabs}')
            lines.append('\\usepackage{longtable}')
            lines.append('\\usepackage{multirow}')
            lines.append('\\usepackage{tabularx}')
            lines.append('')

        # Colori
        if self._has_colors or self._has_hyperlinks or self.use_hyperref:
            lines.append('% === Colori ===')
            lines.append('\\usepackage[dvipsnames,svgnames,x11names]{xcolor}')
            lines.append('')

        # Barrato
        if self._has_strikethrough:
            lines.append('% === Testo Barrato ===')
            lines.append('\\usepackage[normalem]{ulem}')
            lines.append('')

        # Liste
        if self._has_lists:
            lines.append('% === Liste ===')
            lines.append('\\usepackage{enumitem}')
            lines.append('')

        # Blocchi di codice
        if self._has_code_blocks:
            lines.append('% === Codice ===')
            lines.append('\\usepackage{listings}')
            lines.append('\\lstset{')
            lines.append('  basicstyle=\\ttfamily\\small,')
            lines.append('  breaklines=true,')
            lines.append('  columns=fullflexible,')
            lines.append('  frame=single,')
            lines.append('}')
            lines.append('')

        # Hyperref (sempre utile per i link interni)
        if self.use_hyperref:
            lines.append('% === Hyperlink ===')
            lines.append('\\usepackage{hyperref}')
            lines.append('\\hypersetup{')
            lines.append('  colorlinks=true,')
            lines.append('  linkcolor=blue!70!black,')
            lines.append('  urlcolor=blue!70!black,')
            lines.append('  citecolor=green!50!black,')
            lines.append('  pdfauthor={' + self._escape(self.metadata.get('author', '')) + '},')
            lines.append('  pdftitle={' + self._escape(self.metadata.get('title', '')) + '},')
            lines.append('}')
            lines.append('')

        # Utilità varie
        lines.append('% === Utilità ===')
        lines.append('\\usepackage{parskip}       % Spaziatura tra paragrafi')
        lines.append('\\usepackage{setspace}       % Interlinea')
        lines.append('\\usepackage{fancyhdr}       % Header/Footer')
        lines.append('\\usepackage{lastpage}       % Numero ultima pagina')
        lines.append('\\usepackage{caption}        % Caption per figure/tabelle')
        lines.append('\\usepackage{textcomp}       % Simboli aggiuntivi')
        lines.append('')

        # Header/Footer
        lines.append('% === Header e Footer ===')
        lines.append('\\pagestyle{fancy}')
        lines.append('\\fancyhf{}')
        title = self.metadata.get('title', '')
        if title:
            lines.append(f'\\fancyhead[L]{{\\small {self._escape(title)}}}')
        lines.append('\\fancyhead[R]{\\small \\thepage\\ / \\pageref{LastPage}}')
        lines.append('\\renewcommand{\\headrulewidth}{0.4pt}')
        lines.append('\\renewcommand{\\footrulewidth}{0pt}')
        lines.append('')

        # Interlinea
        lines.append('% === Interlinea ===')
        lines.append('\\onehalfspacing')
        lines.append('')

        return '\n'.join(lines)

    def _generate_begin_document(self) -> str:
        return '\\begin{document}\n'

    def _generate_end_document(self) -> str:
        return '\n\\end{document}'

    def _generate_title_block(self) -> str:
        """Genera il blocco titolo se presente nei metadati."""
        title = self.metadata.get('title', '')
        author = self.metadata.get('author', '')

        if not title:
            # Cerca un titolo nel primo heading di livello 0
            for elem in self.elements:
                if elem.element_type == ElementType.HEADING and elem.level == 0:
                    title = self._runs_to_plain_text(elem.runs)
                    break

        if not title:
            return ''

        lines = []
        lines.append(f'\\title{{{self._escape(title)}}}')
        if author:
            lines.append(f'\\author{{{self._escape(author)}}}')
        lines.append('\\date{}')  # Nessuna data
        lines.append('\\maketitle')
        lines.append('\\thispagestyle{fancy}')
        lines.append('')
        return '\n'.join(lines)

    def _generate_body(self) -> str:
        """Genera il corpo del documento LaTeX."""
        lines = []
        i = 0
        in_list = False
        current_list_type = None
        current_list_depth = 0
        skip_title = False

        while i < len(self.elements):
            elem = self.elements[i]

            # Salta il primo heading level 0 se è stato usato come titolo
            if (elem.element_type == ElementType.HEADING and elem.level == 0
                    and not skip_title):
                skip_title = True
                i += 1
                continue

            # Gestione liste: raggruppa elementi lista consecutivi
            if elem.element_type == ElementType.LIST_ITEM:
                if not in_list:
                    # Inizia nuova lista
                    in_list = True
                    current_list_type = elem.list_type
                    current_list_depth = elem.list_depth
                    lines.append(self._list_begin(elem.list_type))

                elif elem.list_depth > current_list_depth:
                    # Livello più profondo
                    for _ in range(elem.list_depth - current_list_depth):
                        lines.append(self._list_begin(elem.list_type))
                    current_list_depth = elem.list_depth

                elif elem.list_depth < current_list_depth:
                    # Livello meno profondo
                    for _ in range(current_list_depth - elem.list_depth):
                        lines.append(self._list_end(current_list_type))
                    current_list_depth = elem.list_depth

                text = self._runs_to_latex(elem.runs)
                lines.append(f'  \\item {text}')

            else:
                # Chiudi la lista se aperta
                if in_list:
                    for _ in range(current_list_depth + 1):
                        lines.append(self._list_end(current_list_type))
                    in_list = False
                    current_list_depth = 0
                    lines.append('')

                # Genera l'elemento
                latex = self._element_to_latex(elem)
                if latex:
                    lines.append(latex)

            i += 1

        # Chiudi lista eventualmente rimasta aperta
        if in_list:
            for _ in range(current_list_depth + 1):
                lines.append(self._list_end(current_list_type))

        return '\n'.join(lines)

    def _element_to_latex(self, elem: DocumentElement) -> str:
        """Converte un singolo elemento in LaTeX."""
        if elem.element_type == ElementType.HEADING:
            return self._heading_to_latex(elem)
        elif elem.element_type == ElementType.PARAGRAPH:
            return self._paragraph_to_latex(elem)
        elif elem.element_type == ElementType.IMAGE:
            return self._image_to_latex(elem)
        elif elem.element_type == ElementType.TABLE:
            return self._table_to_latex(elem)
        elif elem.element_type == ElementType.PAGE_BREAK:
            return '\\newpage\n'
        elif elem.element_type == ElementType.HORIZONTAL_RULE:
            return '\\noindent\\rule{\\textwidth}{0.4pt}\n'
        elif elem.element_type == ElementType.CODE_BLOCK:
            return self._code_block_to_latex(elem)
        elif elem.element_type == ElementType.FOOTNOTE:
            return ''  # Gestite inline
        return ''

    def _heading_to_latex(self, elem: DocumentElement) -> str:
        """Converte un heading in LaTeX."""
        text = self._runs_to_latex(elem.runs)
        if not text.strip():
            return ''

        level_map = {
            0: 'title',
            1: 'section',
            2: 'subsection',
            3: 'subsubsection',
            4: 'paragraph',
            5: 'subparagraph',
        }

        cmd = level_map.get(elem.level, 'subparagraph')
        if cmd == 'title':
            return ''  # Titolo gestito nel title block

        result = f'\\{cmd}{{{text}}}'

        # Allineamento per heading (raro ma possibile)
        if elem.alignment == Alignment.CENTER:
            result = f'\\begin{{center}}\n{result}\n\\end{{center}}'

        return result + '\n'

    def _paragraph_to_latex(self, elem: DocumentElement) -> str:
        """Converte un paragrafo in LaTeX."""
        text = self._runs_to_latex(elem.runs)

        # Paragrafo vuoto -> riga vuota
        if not text.strip():
            return ''

        # Gestione allineamento
        if elem.alignment == Alignment.CENTER:
            return f'\\begin{{center}}\n{text}\n\\end{{center}}\n'
        elif elem.alignment == Alignment.RIGHT:
            return f'\\begin{{flushright}}\n{text}\n\\end{{flushright}}\n'
        elif elem.alignment == Alignment.JUSTIFY:
            return text + '\n'
        else:
            # Gestisci indentazione
            if elem.indent_level > 0:
                indent = '\\quad ' * elem.indent_level
                return f'{indent}{text}\n'
            return text + '\n'

    def _code_block_to_latex(self, elem: DocumentElement) -> str:
        """Converte un blocco di codice in LaTeX (listings)."""
        if not elem.runs:
            return ''

        code_text = ''.join(run.text for run in elem.runs).strip('\n')
        if not code_text.strip():
            return ''

        return f'\\begin{{lstlisting}}\n{code_text}\n\\end{{lstlisting}}\n'

    def _image_to_latex(self, elem: DocumentElement) -> str:
        """Converte un'immagine in LaTeX."""
        if not elem.image_path:
            return ''

        img_path = elem.image_path
        # Converti il percorso per LaTeX
        img_basename = os.path.basename(img_path)
        # Rimuovi estensione (LaTeX la aggiunge da solo)
        img_name = os.path.splitext(img_basename)[0]

        # Determina le opzioni di dimensione
        size_opts = []
        if elem.image_width:
            width_cm = elem.image_width
            # Se l'immagine è più larga della textwidth, scala
            if width_cm > 15:
                size_opts.append('width=\\textwidth')
            else:
                size_opts.append(f'width={width_cm:.1f}cm')
        else:
            size_opts.append('width=0.8\\textwidth')

        if elem.image_height and not elem.image_width:
            size_opts.append(f'height={elem.image_height:.1f}cm')

        size_opts.append('keepaspectratio')
        size_str = ', '.join(size_opts)

        lines = [
            '\\begin{figure}[H]',
            '  \\centering',
            f'  \\includegraphics[{size_str}]{{{img_basename}}}',
        ]

        if elem.image_caption:
            lines.append(f'  \\caption{{{self._escape(elem.image_caption)}}}')

        lines.append('\\end{figure}')
        lines.append('')

        return '\n'.join(lines)

    def _table_to_latex(self, elem: DocumentElement) -> str:
        """Converte una tabella in LaTeX usando booktabs."""
        if not elem.table_rows:
            return ''

        # Calcola il numero di colonne
        max_cols = max(len(row) for row in elem.table_rows) if elem.table_rows else 0
        if max_cols == 0:
            return ''

        # Usa tabularx per tabelle adattive
        col_spec = '|' + '|'.join(['X'] * max_cols) + '|'

        lines = [
            '\\begin{table}[H]',
            '  \\centering',
            f'  \\begin{{tabularx}}{{\\textwidth}}{{{col_spec}}}',
            '    \\hline',
        ]

        for row_idx, row in enumerate(elem.table_rows):
            cells_text = []
            for cell in row:
                if cell.rowspan == 0:
                    cells_text.append('')  # Cella merged
                    continue

                # Per header, forza il bold (rimuovi bold esistente per evitare doppioni)
                if row_idx < elem.table_header_rows:
                    cell_content = self._runs_to_plain_text(cell.runs)
                    cell_content = self._escape(cell_content.strip())
                    cell_content = f'\\textbf{{{cell_content}}}'
                else:
                    cell_content = self._runs_to_latex(cell.runs)
                    cell_content = cell_content.strip()

                if cell.colspan > 1:
                    cell_content = (
                        f'\\multicolumn{{{cell.colspan}}}{{|c|}}{{{cell_content}}}'
                    )

                cells_text.append(cell_content)

            # Filtra celle vuote da merge
            line = ' & '.join(cells_text)
            lines.append(f'    {line} \\\\')
            lines.append('    \\hline')

        lines.append('  \\end{tabularx}')
        lines.append('\\end{table}')
        lines.append('')

        return '\n'.join(lines)

    # =========================================================================
    # TEXT RUN PROCESSING
    # =========================================================================

    def _runs_to_latex(self, runs: list) -> str:
        """Converte una lista di TextRun in codice LaTeX."""
        if not runs:
            return ''

        parts = []
        for run in runs:
            text = self._escape(run.text)
            if not text:
                continue

            # Applica stili
            if TextStyle.MONOSPACE in run.styles:
                text = f'\\texttt{{{text}}}'
            if TextStyle.BOLD in run.styles:
                text = f'\\textbf{{{text}}}'
            if TextStyle.ITALIC in run.styles:
                text = f'\\textit{{{text}}}'
            if TextStyle.UNDERLINE in run.styles and not run.hyperlink:
                text = f'\\underline{{{text}}}'
            if TextStyle.STRIKETHROUGH in run.styles:
                text = f'\\sout{{{text}}}'
            if TextStyle.SUPERSCRIPT in run.styles:
                text = f'\\textsuperscript{{{text}}}'
            if TextStyle.SUBSCRIPT in run.styles:
                text = f'\\textsubscript{{{text}}}'
            if TextStyle.SMALL_CAPS in run.styles:
                text = f'\\textsc{{{text}}}'

            # Colore
            if run.font_color and run.font_color != '000000':
                r = int(run.font_color[0:2], 16)
                g = int(run.font_color[2:4], 16)
                b = int(run.font_color[4:6], 16)
                text = (
                    f'\\textcolor[RGB]{{{r},{g},{b}}}{{{text}}}'
                )

            # Hyperlink
            if run.hyperlink:
                text = f'\\href{{{self._escape_url(run.hyperlink)}}}{{{text}}}'

            parts.append(text)

        return ''.join(parts)

    def _runs_to_plain_text(self, runs: list) -> str:
        """Converte TextRun in testo semplice (senza formattazione)."""
        return ''.join(run.text for run in runs)

    # =========================================================================
    # ESCAPING
    # =========================================================================

    def _escape(self, text: str) -> str:
        """Esegue l'escape dei caratteri speciali LaTeX."""
        if not text:
            return ''

        # Ordine importante: backslash per primo
        replacements = [
            ('\\', '\\textbackslash{}'),
            ('&', '\\&'),
            ('%', '\\%'),
            ('$', '\\$'),
            ('#', '\\#'),
            ('_', '\\_'),
            ('{', '\\{'),
            ('}', '\\}'),
            ('~', '\\textasciitilde{}'),
            ('^', '\\textasciicircum{}'),
        ]

        for old, new in replacements:
            text = text.replace(old, new)

        # Gestisci virgolette
        text = text.replace('``', "''")  # Fix doppi backtick
        text = text.replace('"', "''")
        text = text.replace('"', "``")
        text = text.replace('"', "''")
        text = text.replace(''', "'")
        text = text.replace(''', "`")

        # Gestisci em-dash e en-dash
        text = text.replace('—', '---')
        text = text.replace('–', '--')
        text = text.replace('…', '\\ldots{}')

        return text

    def _escape_url(self, url: str) -> str:
        """Esegue l'escape di un URL per LaTeX."""
        if not url:
            return ''
        # Solo # e % necessitano escape negli URL
        url = url.replace('%', '\\%')
        return url

    # =========================================================================
    # LIST HELPERS
    # =========================================================================

    def _list_begin(self, list_type: ListType) -> str:
        if list_type == ListType.NUMBERED:
            return '\\begin{enumerate}'
        return '\\begin{itemize}'

    def _list_end(self, list_type: ListType) -> str:
        if list_type == ListType.NUMBERED:
            return '\\end{enumerate}'
        return '\\end{itemize}'

    # =========================================================================
    # FILE OUTPUT
    # =========================================================================

    def write_to_file(self, output_path: str) -> str:
        """Scrive il codice LaTeX su file."""
        latex_code = self.generate()
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(latex_code)

        return str(output_path)
