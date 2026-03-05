#!/usr/bin/env python3
"""
WordToLaTeX - Interfaccia a linea di comando

Converte documenti Word (docx, doc, odt, rtf, pptx, html, epub, txt, md, ipynb)
in PDF con stile LaTeX.

Utilizzo:
    python -m wordtolatex documento.docx
    python -m wordtolatex documento.docx -o output.pdf
    python -m wordtolatex documento.docx --keep-tex
    python -m wordtolatex --gui
    python -m wordtolatex --check  (verifica installazione)
"""

import argparse
import sys
import os

from .converter import WordToLatexConverter
from .compiler import check_latex_installation
from .parser import DocumentParser


def main():
    parser = argparse.ArgumentParser(
        prog='wordtolatex',
        description=(
            'WordToLaTeX - Converte documenti in PDF con stile LaTeX.\n'
            'Formati supportati: .docx, .doc, .odt, .rtf, .pptx, .html, .epub, .txt, .md, .ipynb'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'Esempi:\n'
            '  wordtolatex documento.docx\n'
            '  wordtolatex documento.docx -o risultato.pdf\n'
            '  wordtolatex documento.docx --keep-tex --keep-images\n'
            '  wordtolatex presentazione.pptx\n'
            '  wordtolatex pagina.html --font-size 12\n'
            '  wordtolatex documento.md\n'
            '  wordtolatex notebook.ipynb\n'
            '  wordtolatex --gui\n'
            '  wordtolatex --tex-only documento.docx -o documento.tex\n'
            '  wordtolatex --check\n'
        ),
    )

    parser.add_argument(
        'input',
        nargs='?',
        help='File di input (docx, doc, odt, rtf, pptx, html, epub, txt, md, ipynb)',
    )

    parser.add_argument(
        '-o', '--output',
        help='Percorso del file di output (default: stesso nome con .pdf)',
    )

    parser.add_argument(
        '--tex-only',
        action='store_true',
        help='Genera solo il file .tex senza compilare in PDF',
    )

    parser.add_argument(
        '--keep-tex',
        action='store_true',
        help='Mantieni il file .tex intermedio dopo la compilazione',
    )

    parser.add_argument(
        '--keep-images',
        action='store_true',
        help='Mantieni le immagini estratte in una cartella separata',
    )

    # Opzioni di formattazione
    format_group = parser.add_argument_group('Opzioni di formattazione')

    format_group.add_argument(
        '--document-class',
        default='article',
        choices=['article', 'report', 'book', 'scrartcl', 'scrreprt', 'scrbook'],
        help='Classe del documento LaTeX (default: article)',
    )

    format_group.add_argument(
        '--font-size',
        type=int,
        default=11,
        choices=[10, 11, 12],
        help='Dimensione del font in pt (default: 11)',
    )

    format_group.add_argument(
        '--paper-size',
        default='a4paper',
        choices=['a4paper', 'letterpaper', 'a5paper', 'b5paper'],
        help='Formato della carta (default: a4paper)',
    )

    format_group.add_argument(
        '--language',
        default='italian',
        help='Lingua per babel (default: italian)',
    )

    format_group.add_argument(
        '--engine',
        choices=['pdflatex', 'lualatex', 'xelatex'],
        help='Motore LaTeX da usare (default: auto-detect)',
    )

    # Utilità
    parser.add_argument(
        '--gui',
        action='store_true',
        help='Avvia l\'interfaccia grafica',
    )

    parser.add_argument(
        '--check',
        action='store_true',
        help='Verifica l\'installazione di LaTeX e le dipendenze',
    )

    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 1.0.0',
    )

    args = parser.parse_args()

    # === Modalità GUI ===
    if args.gui:
        from .gui import main as gui_main
        gui_main()
        return

    # === Modalità check ===
    if args.check:
        run_check()
        return

    # === Conversione ===
    if not args.input:
        parser.print_help()
        print("\nErrore: specifica un file di input.")
        sys.exit(1)

    if not os.path.exists(args.input):
        print(f"Errore: file non trovato: {args.input}")
        sys.exit(1)

    # Verifica estensione
    ext = os.path.splitext(args.input)[1].lower()
    if ext not in DocumentParser.SUPPORTED_EXTENSIONS:
        print(
            f"Errore: formato '{ext}' non supportato.\n"
            f"Formati supportati: {', '.join(DocumentParser.SUPPORTED_EXTENSIONS)}"
        )
        sys.exit(1)

    try:
        converter = WordToLatexConverter(
            document_class=args.document_class,
            font_size=args.font_size,
            paper_size=args.paper_size,
            language=args.language,
            latex_engine=args.engine,
            keep_tex=args.keep_tex,
            keep_images=args.keep_images,
        )

        if args.tex_only:
            result = converter.convert_to_tex(args.input, args.output)
        else:
            result = converter.convert(args.input, args.output)

        print(f"Output: {result}")

    except FileNotFoundError as e:
        print(f"Errore: {e}")
        sys.exit(1)
    except ImportError as e:
        print(f"Errore dipendenza: {e}")
        print("Installa le dipendenze con: pip install -r requirements.txt")
        sys.exit(1)
    except RuntimeError as e:
        print(f"Errore: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Errore imprevisto: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def run_check():
    """Verifica lo stato dell'installazione."""
    print("\n" + "=" * 60)
    print("  WordToLaTeX - Verifica Installazione")
    print("=" * 60)

    # Verifica dipendenze Python
    print("\n📦 Dipendenze Python:")
    python_deps = {
        'python-docx': 'docx',
        'odfpy': 'odf',
        'python-pptx': 'pptx',
    }

    all_ok = True
    for pkg_name, import_name in python_deps.items():
        try:
            __import__(import_name)
            print(f"  ✅ {pkg_name}: installato")
        except ImportError:
            print(f"  ❌ {pkg_name}: NON installato")
            all_ok = False

    # Verifica LaTeX
    print("\n📄 Distribuzione LaTeX:")
    status = check_latex_installation()

    if not status['available']:
        print("  ❌ Nessun motore LaTeX trovato!")
        all_ok = False
    else:
        for engine, info in status['engines'].items():
            if info['installed']:
                print(f"  ✅ {engine}: {info['path']}")
            else:
                print(f"  ⬜ {engine}: non trovato")

    if status['packages']:
        print("\n📦 Pacchetti LaTeX:")
        missing = []
        for pkg, info in status['packages'].items():
            if info['installed']:
                print(f"  ✅ {pkg}")
            else:
                print(f"  ❌ {pkg}: NON trovato")
                missing.append(pkg)
                all_ok = False

        if missing:
            print(f"\n  ⚠️  Pacchetti mancanti: {', '.join(missing)}")
            print("  Installa con:")
            print("    sudo apt install texlive-latex-extra texlive-fonts-recommended "
                  "texlive-lang-italian texlive-fonts-extra")

    # Verifica LibreOffice (per .doc/.rtf)
    print("\n📎 LibreOffice (per file .doc/.rtf):")
    import shutil
    lo_path = shutil.which('libreoffice') or shutil.which('soffice')
    if lo_path:
        print(f"  ✅ LibreOffice: {lo_path}")
    else:
        print("  ⚠️  LibreOffice: non trovato (necessario solo per .doc e .rtf)")

    print("\n" + "=" * 60)
    if all_ok:
        print("  ✅ Tutto pronto! Puoi usare WordToLaTeX.")
    else:
        print("  ⚠️  Alcune dipendenze mancano. Vedi sopra per i dettagli.")
    print("=" * 60 + "\n")


if __name__ == '__main__':
    main()
