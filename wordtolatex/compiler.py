"""
Modulo per la compilazione del file LaTeX in PDF.
Usa pdflatex o lualatex o xelatex.
"""

import os
import subprocess
import shutil
import tempfile
import locale
from pathlib import Path
from typing import Optional


class PDFCompiler:
    """Compila un file .tex in .pdf usando un motore LaTeX."""

    # Motori supportati in ordine di preferenza
    ENGINES = ['pdflatex', 'lualatex', 'xelatex']

    def __init__(
        self,
        engine: str = None,
        num_passes: int = 2,
        shell_escape: bool = False,
        clean_aux: bool = True,
    ):
        self.engine = engine or self._find_engine()
        self.num_passes = num_passes
        self.shell_escape = shell_escape
        self.clean_aux = clean_aux

    def _find_engine(self) -> str:
        """Trova il primo motore LaTeX disponibile nel sistema."""
        for engine in self.ENGINES:
            if shutil.which(engine):
                return engine

        raise RuntimeError(
            "Nessun motore LaTeX trovato nel sistema.\n"
            "Installa una distribuzione TeX:\n"
            "  Ubuntu/Debian: sudo apt install texlive-full\n"
            "  oppure:        sudo apt install texlive-latex-extra texlive-fonts-recommended "
            "texlive-lang-italian texlive-fonts-extra latexmk\n"
            "  Fedora:        sudo dnf install texlive-scheme-full\n"
            "  macOS:         brew install --cask mactex\n"
            "  Arch:          sudo pacman -S texlive-most"
        )

    def compile(self, tex_path: str, output_dir: str = None) -> str:
        """
        Compila un file .tex in .pdf.

        Args:
            tex_path: Percorso del file .tex
            output_dir: Directory di output (default: stessa del .tex)

        Returns:
            Percorso del file PDF generato.
        """
        tex_path = Path(tex_path).resolve()
        if not tex_path.exists():
            raise FileNotFoundError(f"File .tex non trovato: {tex_path}")

        if output_dir:
            output_dir = Path(output_dir).resolve()
            output_dir.mkdir(parents=True, exist_ok=True)
        else:
            output_dir = tex_path.parent

        # Compila nella directory del file .tex per gestire i percorsi relativi
        work_dir = tex_path.parent

        cmd = [self.engine]
        cmd.extend(['-interaction=nonstopmode'])
        cmd.extend(['-halt-on-error'])
        cmd.extend([f'-output-directory={output_dir}'])

        if self.shell_escape:
            cmd.append('-shell-escape')

        cmd.append(str(tex_path))

        print(f"  Compilazione con {self.engine}...")

        for pass_num in range(1, self.num_passes + 1):
            print(f"  Pass {pass_num}/{self.num_passes}...")

            result = subprocess.run(
                cmd,
                cwd=str(work_dir),
                capture_output=True,
                text=False,
                timeout=300,
            )

            stdout_text = self._decode_process_output(result.stdout)
            stderr_text = self._decode_process_output(result.stderr)

            if result.returncode != 0:
                # Cerca l'errore specifico nel log
                error_msg = self._extract_latex_error(stdout_text)
                if not error_msg:
                    error_msg = self._extract_latex_error(stderr_text)
                if not error_msg:
                    # Prendi le ultime righe dello stdout
                    lines = stdout_text.strip().split('\n')
                    error_msg = '\n'.join(lines[-20:])

                raise RuntimeError(
                    f"Errore nella compilazione LaTeX (pass {pass_num}):\n{error_msg}"
                )

        # Percorso del PDF generato
        pdf_name = tex_path.stem + '.pdf'
        pdf_path = output_dir / pdf_name

        if not pdf_path.exists():
            raise RuntimeError(
                f"Il PDF non è stato generato. Controlla il file .tex per errori."
            )

        # Pulizia file ausiliari
        if self.clean_aux:
            self._clean_auxiliary_files(output_dir, tex_path.stem)

        print(f"  PDF generato: {pdf_path}")
        return str(pdf_path)

    def _decode_process_output(self, data: bytes) -> str:
        """Decodifica output bytes dei processi esterni in modo robusto."""
        if not data:
            return ''

        for encoding in ('utf-8', locale.getpreferredencoding(False), 'latin-1'):
            try:
                return data.decode(encoding)
            except (UnicodeDecodeError, LookupError):
                continue

        return data.decode('utf-8', errors='replace')

    def _extract_latex_error(self, log_text: str) -> Optional[str]:
        """Estrae il messaggio di errore dal log LaTeX."""
        if not log_text:
            return None

        error_lines = []
        in_error = False

        for line in log_text.split('\n'):
            if line.startswith('!'):
                in_error = True
                error_lines.append(line)
            elif in_error:
                error_lines.append(line)
                if line.strip() == '' or len(error_lines) > 10:
                    in_error = False

        if error_lines:
            return '\n'.join(error_lines[:15])

        return None

    def _clean_auxiliary_files(self, directory: Path, stem: str):
        """Rimuove i file ausiliari della compilazione LaTeX."""
        aux_extensions = [
            '.aux', '.log', '.out', '.toc', '.lof', '.lot',
            '.bbl', '.blg', '.nav', '.snm', '.vrb',
            '.fdb_latexmk', '.fls', '.synctex.gz',
        ]

        for ext in aux_extensions:
            aux_file = directory / (stem + ext)
            if aux_file.exists():
                try:
                    aux_file.unlink()
                except OSError:
                    pass


def check_latex_installation() -> dict:
    """
    Verifica lo stato dell'installazione LaTeX nel sistema.

    Returns:
        Dizionario con lo stato dei componenti.
    """
    status = {
        'engines': {},
        'packages': {},
        'available': False,
    }

    # Controlla motori
    for engine in PDFCompiler.ENGINES:
        path = shutil.which(engine)
        status['engines'][engine] = {
            'installed': path is not None,
            'path': path or 'non trovato',
        }
        if path:
            status['available'] = True

    # Controlla kpsewhich per verificare i pacchetti
    kpsewhich = shutil.which('kpsewhich')
    if kpsewhich:
        important_packages = [
            'geometry.sty', 'graphicx.sty', 'hyperref.sty',
            'booktabs.sty', 'xcolor.sty', 'microtype.sty',
            'babel.sty', 'fancyhdr.sty', 'enumitem.sty',
            'parskip.sty', 'setspace.sty', 'tabularx.sty',
            'lmodern.sty', 'caption.sty', 'float.sty',
        ]

        for pkg in important_packages:
            try:
                result = subprocess.run(
                    [kpsewhich, pkg],
                    capture_output=True, text=True, timeout=10,
                )
                status['packages'][pkg] = {
                    'installed': result.returncode == 0,
                    'path': result.stdout.strip() if result.returncode == 0 else 'non trovato',
                }
            except Exception:
                status['packages'][pkg] = {
                    'installed': False,
                    'path': 'errore nel controllo',
                }

    return status
