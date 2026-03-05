# WordToLaTeX

Convertitore da documenti **Word / ODT / PowerPoint / HTML / Markdown / EPUB / TXT** a **PDF con stile LaTeX**.

Dispone di un'**interfaccia grafica** (GUI) e di un'interfaccia a **linea di comando** (CLI).

---

## Funzionalità

| Elemento           | Supporto |
|-------------------|----------|
| Testo e paragrafi | ✅       |
| Grassetto, corsivo, sottolineato | ✅ |
| Barrato, apice, pedice | ✅ |
| Intestazioni (H1-H6) | ✅ |
| Liste puntate e numerate | ✅ |
| Tabelle            | ✅       |
| Immagini inline    | ✅       |
| Hyperlink          | ✅       |
| Colori testo       | ✅       |
| Page break         | ✅       |
| Allineamento (sx, dx, centro, giustificato) | ✅ |
| Small caps         | ✅       |

### Formati supportati in input

| Formato    | Estensioni     | Libreria usata               |
|-----------|---------------|------------------------------|
| Word       | `.docx`        | python-docx                  |
| Word legacy| `.doc`         | LibreOffice (conversione)    |
| OpenDocument | `.odt`       | odfpy / LibreOffice          |
| Rich Text  | `.rtf`         | LibreOffice (conversione)    |
| PowerPoint | `.pptx`        | python-pptx                  |
| HTML       | `.html`, `.htm`| html.parser (builtin)        |
| EPUB       | `.epub`        | zipfile + html.parser        |
| Testo      | `.txt`         | builtin                      |
| Markdown   | `.md`          | builtin (parser custom)      |
| Jupyter Notebook | `.ipynb` | json (builtin, parser custom) |

---

## Requisiti di Sistema

### Python
- Python 3.8+

### Distribuzione LaTeX

```bash
# Ubuntu / Debian (consigliata installazione completa)
sudo apt install texlive-full

# oppure i pacchetti minimi necessari:
sudo apt install texlive-latex-extra texlive-fonts-recommended \
    texlive-lang-italian texlive-fonts-extra
```

### LibreOffice (opzionale, per .doc e .rtf)

```bash
sudo apt install libreoffice
```

---

## Installazione

```bash
cd Wordtolatex
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## Utilizzo

### Interfaccia Grafica (GUI)

```bash
python -m wordtolatex --gui
```

Nella GUI puoi anche trascinare un file direttamente nel campo **Input** (drag&drop),
oltre a usare il pulsante **Sfoglia...**.

### Da linea di comando (CLI)

```bash
python -m wordtolatex documento.docx
python -m wordtolatex documento.docx -o risultato.pdf
python -m wordtolatex presentazione.pptx
python -m wordtolatex pagina.html
python -m wordtolatex readme.md
python -m wordtolatex notebook.ipynb
python -m wordtolatex libro.epub
python -m wordtolatex --tex-only documento.docx
python -m wordtolatex --check
python -m wordtolatex --gui
```

### Da codice Python

```python
from wordtolatex.converter import WordToLatexConverter

converter = WordToLatexConverter()
pdf_path = converter.convert('documento.docx')
```

---

## Opzioni CLI

| Opzione              | Descrizione                                    |
|---------------------|------------------------------------------------|
| `--gui`              | Avvia l'interfaccia grafica                    |
| `-o, --output`       | Percorso file di output                        |
| `--tex-only`         | Genera solo il .tex                            |
| `--keep-tex`         | Mantieni il .tex dopo la compilazione PDF      |
| `--keep-images`      | Salva le immagini estratte                     |
| `--document-class`   | article, report, book, scrartcl...             |
| `--font-size`        | 10, 11, 12 pt                                 |
| `--paper-size`       | a4paper, letterpaper, a5paper, b5paper         |
| `--language`         | Lingua per babel (default: italian)            |
| `--engine`           | pdflatex, lualatex, xelatex                   |
| `--check`            | Verifica installazione                         |

---

## Struttura Progetto

```
Wordtolatex/
├── wordtolatex/
│   ├── __init__.py          # Package init
│   ├── __main__.py          # CLI entry point
│   ├── gui.py               # Interfaccia grafica (Tkinter)
│   ├── converter.py         # Orchestratore principale
│   ├── parser.py            # Parser documenti (tutti i formati)
│   ├── latex_generator.py   # Generatore codice LaTeX
│   └── compiler.py          # Compilatore LaTeX → PDF
├── requirements.txt
├── pyproject.toml
└── README.md
```

---

## Licenza

MIT
# WordToLatex
