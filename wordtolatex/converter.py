"""
Modulo principale del convertitore: orchestra parser, generatore LaTeX e compilatore PDF.
"""

import os
import shutil
import tempfile
from pathlib import Path
from typing import Optional

from .parser import DocumentParser
from .latex_generator import LaTeXGenerator
from .compiler import PDFCompiler


class WordToLatexConverter:
    """
    Convertitore da Word/ODT a PDF (stile LaTeX).

    Flusso:
        1. Parsing del documento sorgente (docx/doc/odt/rtf)
        2. Generazione codice LaTeX
        3. Compilazione LaTeX → PDF
    """

    def __init__(
        self,
        document_class: str = 'article',
        font_size: int = 11,
        paper_size: str = 'a4paper',
        language: str = 'italian',
        latex_engine: str = None,
        keep_tex: bool = False,
        keep_images: bool = False,
    ):
        self.document_class = document_class
        self.font_size = font_size
        self.paper_size = paper_size
        self.language = language
        self.latex_engine = latex_engine
        self.keep_tex = keep_tex
        self.keep_images = keep_images

    def convert(
        self,
        input_path: str,
        output_path: str = None,
    ) -> str:
        """
        Converte un documento Word/ODT in PDF (stile LaTeX).

        Args:
            input_path:  Percorso del file sorgente (.docx, .doc, .odt, .rtf)
            output_path: Percorso del file PDF di output (opzionale)

        Returns:
            Percorso del file PDF generato.
        """
        input_path = Path(input_path).resolve()

        if not input_path.exists():
            raise FileNotFoundError(f"File non trovato: {input_path}")

        # Determina output path
        if output_path:
            output_path = Path(output_path).resolve()
        else:
            output_path = input_path.with_suffix('.pdf')

        output_dir = output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)

        # Directory di lavoro temporanea
        work_dir = tempfile.mkdtemp(prefix='wordtolatex_')
        image_dir = os.path.join(work_dir, 'images')
        os.makedirs(image_dir, exist_ok=True)

        try:
            # === FASE 1: Parsing ===
            print(f"\n{'='*60}")
            print(f"  WordToLaTeX Converter")
            print(f"{'='*60}")
            print(f"\n  Input:  {input_path}")
            print(f"  Output: {output_path}\n")
            print(f"[1/3] Parsing del documento...")

            parser = DocumentParser(str(input_path), image_output_dir=image_dir)
            elements = parser.parse()
            metadata = parser.metadata

            print(f"  Trovati {len(elements)} elementi")
            self._print_element_summary(elements)

            # === FASE 2: Generazione LaTeX ===
            print(f"\n[2/3] Generazione codice LaTeX...")

            generator = LaTeXGenerator(
                elements=elements,
                metadata=metadata,
                image_dir=image_dir,
                document_class=self.document_class,
                font_size=self.font_size,
                paper_size=self.paper_size,
                language=self.language,
            )

            tex_filename = input_path.stem + '.tex'
            tex_path = os.path.join(work_dir, tex_filename)
            generator.write_to_file(tex_path)
            print(f"  File .tex generato: {tex_path}")

            # === FASE 3: Compilazione PDF ===
            print(f"\n[3/3] Compilazione PDF...")

            compiler = PDFCompiler(
                engine=self.latex_engine,
                num_passes=2,
                clean_aux=True,
            )

            pdf_temp = compiler.compile(tex_path, output_dir=work_dir)

            # Copia il PDF nella destinazione finale
            shutil.copy2(pdf_temp, str(output_path))

            # Opzionalmente mantieni il file .tex
            if self.keep_tex:
                tex_output = output_path.with_suffix('.tex')
                shutil.copy2(tex_path, str(tex_output))
                print(f"  File .tex salvato: {tex_output}")

            # Opzionalmente mantieni le immagini
            if self.keep_images and os.listdir(image_dir):
                img_output_dir = output_path.parent / 'images'
                if img_output_dir.exists():
                    shutil.rmtree(str(img_output_dir))
                shutil.copytree(image_dir, str(img_output_dir))
                print(f"  Immagini salvate in: {img_output_dir}")

            print(f"\n{'='*60}")
            print(f"  Conversione completata!")
            print(f"  PDF: {output_path}")
            print(f"{'='*60}\n")

            return str(output_path)

        finally:
            # Pulizia directory temporanea
            try:
                shutil.rmtree(work_dir)
            except OSError:
                pass

    def convert_to_tex(
        self,
        input_path: str,
        output_path: str = None,
    ) -> str:
        """
        Converte un documento Word/ODT in LaTeX (solo .tex, senza compilazione).

        Args:
            input_path:  Percorso del file sorgente
            output_path: Percorso del file .tex di output (opzionale)

        Returns:
            Percorso del file .tex generato.
        """
        input_path = Path(input_path).resolve()

        if not input_path.exists():
            raise FileNotFoundError(f"File non trovato: {input_path}")

        if output_path:
            output_path = Path(output_path).resolve()
        else:
            output_path = input_path.with_suffix('.tex')

        output_dir = output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        image_dir = str(output_dir / 'images')

        print(f"\n  Parsing di {input_path.name}...")

        parser = DocumentParser(str(input_path), image_output_dir=image_dir)
        elements = parser.parse()
        metadata = parser.metadata

        print(f"  Generazione LaTeX...")

        generator = LaTeXGenerator(
            elements=elements,
            metadata=metadata,
            image_dir=image_dir,
            document_class=self.document_class,
            font_size=self.font_size,
            paper_size=self.paper_size,
            language=self.language,
        )

        generator.write_to_file(str(output_path))
        print(f"  File .tex generato: {output_path}")

        return str(output_path)

    def _print_element_summary(self, elements):
        """Stampa un riepilogo degli elementi trovati."""
        from .parser import ElementType
        counts = {}
        for elem in elements:
            name = elem.element_type.name
            counts[name] = counts.get(name, 0) + 1

        for name, count in sorted(counts.items()):
            print(f"    - {name}: {count}")
