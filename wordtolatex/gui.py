"""
WordToLaTeX - Interfaccia Grafica (GUI) con Tkinter.

Permette di:
- Selezionare file di input (docx, doc, odt, rtf, pptx, html, epub, txt, md, ipynb)
- Configurare le opzioni di conversione
- Convertire in PDF o solo in .tex
- Visualizzare il progresso e i log in tempo reale
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from io import StringIO

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_TK_DND = True
except ImportError:
    HAS_TK_DND = False

try:
    from .converter import WordToLatexConverter
    from .parser import DocumentParser
    from .compiler import check_latex_installation
except ImportError:
    from wordtolatex.converter import WordToLatexConverter
    from wordtolatex.parser import DocumentParser
    from wordtolatex.compiler import check_latex_installation


# ===========================================================================
# Colori e stile
# ===========================================================================
COLORS = {
    'bg':            '#1e1e2e',
    'bg_secondary':  '#2a2a3d',
    'bg_card':       '#313145',
    'accent':        '#7c3aed',
    'accent_hover':  '#6d28d9',
    'accent_light':  '#a78bfa',
    'success':       '#22c55e',
    'error':         '#ef4444',
    'warning':       '#f59e0b',
    'text':          '#e2e8f0',
    'text_dim':      '#94a3b8',
    'text_bright':   '#f8fafc',
    'border':        '#404060',
    'input_bg':      '#252538',
    'input_border':  '#4a4a6a',
    'btn_secondary': '#3b3b55',
}

FONT_FAMILY = 'Segoe UI'
if sys.platform == 'linux':
    FONT_FAMILY = 'Ubuntu'
elif sys.platform == 'darwin':
    FONT_FAMILY = 'SF Pro Display'


class WordToLatexGUI:
    """Interfaccia grafica principale."""

    SUPPORTED_FILETYPES = [
        ("Tutti i formati supportati",
         "*.docx *.doc *.odt *.rtf *.pptx *.html *.htm *.epub *.txt *.md *.ipynb"),
        ("Word", "*.docx *.doc"),
        ("OpenDocument", "*.odt"),
        ("Rich Text", "*.rtf"),
        ("PowerPoint", "*.pptx"),
        ("HTML", "*.html *.htm"),
        ("EPUB", "*.epub"),
        ("Testo", "*.txt"),
        ("Markdown", "*.md"),
        ("Jupyter Notebook", "*.ipynb"),
        ("Tutti i file", "*.*"),
    ]

    def __init__(self):
        if HAS_TK_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("WordToLaTeX Converter")
        self.root.geometry("820x720")
        self.root.minsize(700, 620)
        self.root.configure(bg=COLORS['bg'])
        self.dragdrop_enabled = HAS_TK_DND

        # Variabili
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.doc_class = tk.StringVar(value='article')
        self.font_size = tk.StringVar(value='11')
        self.paper_size = tk.StringVar(value='a4paper')
        self.language = tk.StringVar(value='italian')
        self.engine = tk.StringVar(value='auto')
        self.keep_tex = tk.BooleanVar(value=False)
        self.keep_images = tk.BooleanVar(value=False)
        self.tex_only = tk.BooleanVar(value=False)

        self.is_converting = False

        self._setup_styles()
        self._build_ui()
        self._enable_drag_and_drop()
        self._center_window()

    def _setup_styles(self):
        """Configura gli stili ttk."""
        style = ttk.Style()
        style.theme_use('clam')

        # Frame
        style.configure('Card.TFrame', background=COLORS['bg_card'])
        style.configure('Main.TFrame', background=COLORS['bg'])

        # Label
        style.configure('Title.TLabel',
                        background=COLORS['bg'],
                        foreground=COLORS['accent_light'],
                        font=(FONT_FAMILY, 22, 'bold'))
        style.configure('Subtitle.TLabel',
                        background=COLORS['bg'],
                        foreground=COLORS['text_dim'],
                        font=(FONT_FAMILY, 10))
        style.configure('Section.TLabel',
                        background=COLORS['bg_card'],
                        foreground=COLORS['text_bright'],
                        font=(FONT_FAMILY, 11, 'bold'))
        style.configure('Field.TLabel',
                        background=COLORS['bg_card'],
                        foreground=COLORS['text'],
                        font=(FONT_FAMILY, 9))
        style.configure('Status.TLabel',
                        background=COLORS['bg'],
                        foreground=COLORS['text_dim'],
                        font=(FONT_FAMILY, 9))

        # Checkbutton
        style.configure('Card.TCheckbutton',
                        background=COLORS['bg_card'],
                        foreground=COLORS['text'],
                        font=(FONT_FAMILY, 9))
        style.map('Card.TCheckbutton',
                  background=[('active', COLORS['bg_card'])])

        # Combobox
        style.configure('Card.TCombobox',
                        foreground=COLORS['text'],
                        fieldbackground=COLORS['input_bg'],
                        background=COLORS['input_bg'],
                        font=(FONT_FAMILY, 9))
        style.map('Card.TCombobox',
                  fieldbackground=[('readonly', COLORS['input_bg'])],
                  selectbackground=[('readonly', COLORS['accent'])],
                  selectforeground=[('readonly', COLORS['text_bright'])])

        # Progressbar
        style.configure('Accent.Horizontal.TProgressbar',
                        troughcolor=COLORS['bg_secondary'],
                        background=COLORS['accent'],
                        borderwidth=0,
                        thickness=6)

    def _build_ui(self):
        """Costruisce l'interfaccia utente."""
        # Container principale con padding
        main = tk.Frame(self.root, bg=COLORS['bg'], padx=24, pady=16)
        main.pack(fill='both', expand=True)

        # === HEADER ===
        header = tk.Frame(main, bg=COLORS['bg'])
        header.pack(fill='x', pady=(0, 16))

        ttk.Label(header, text="WordToLaTeX", style='Title.TLabel').pack(
            side='left')
        ttk.Label(header, text="  Converti documenti in PDF con stile LaTeX",
                  style='Subtitle.TLabel').pack(side='left', padx=(8, 0), pady=(8, 0))

        # === FILE INPUT / OUTPUT ===
        file_frame = tk.Frame(main, bg=COLORS['bg_card'], bd=0,
                              highlightbackground=COLORS['border'],
                              highlightthickness=1, padx=16, pady=14)
        file_frame.pack(fill='x', pady=(0, 10))

        ttk.Label(file_frame, text="📂  File", style='Section.TLabel').pack(
            anchor='w', pady=(0, 10))

        # Input row
        input_row = tk.Frame(file_frame, bg=COLORS['bg_card'])
        input_row.pack(fill='x', pady=(0, 6))
        ttk.Label(input_row, text="Input:", style='Field.TLabel').pack(
            side='left', padx=(0, 8))
        self.input_entry = tk.Entry(
            input_row, textvariable=self.input_path,
            font=(FONT_FAMILY, 9),
            bg=COLORS['input_bg'], fg=COLORS['text'],
            insertbackground=COLORS['text'],
            relief='flat', bd=0, highlightthickness=1,
            highlightcolor=COLORS['accent'],
            highlightbackground=COLORS['input_border'],
        )
        self.input_entry.pack(side='left', fill='x', expand=True, ipady=4, padx=(0, 6))
        self._make_button(input_row, "Sfoglia…", self._browse_input,
                          bg=COLORS['btn_secondary'], width=9).pack(side='right')

        # Output row
        output_row = tk.Frame(file_frame, bg=COLORS['bg_card'])
        output_row.pack(fill='x')
        ttk.Label(output_row, text="Output:", style='Field.TLabel').pack(
            side='left', padx=(0, 4))
        self.output_entry = tk.Entry(
            output_row, textvariable=self.output_path,
            font=(FONT_FAMILY, 9),
            bg=COLORS['input_bg'], fg=COLORS['text'],
            insertbackground=COLORS['text'],
            relief='flat', bd=0, highlightthickness=1,
            highlightcolor=COLORS['accent'],
            highlightbackground=COLORS['input_border'],
        )
        self.output_entry.pack(side='left', fill='x', expand=True, ipady=4, padx=(0, 6))
        self._make_button(output_row, "Sfoglia…", self._browse_output,
                          bg=COLORS['btn_secondary'], width=9).pack(side='right')

        dnd_text = "Trascina qui un file per impostare automaticamente l'input"
        if not self.dragdrop_enabled:
            dnd_text += " (installa tkinterdnd2 per attivare drag&drop)"
        ttk.Label(file_frame, text=dnd_text, style='Subtitle.TLabel').pack(
            anchor='w', pady=(8, 0))

        # === OPZIONI ===
        opts_frame = tk.Frame(main, bg=COLORS['bg_card'], bd=0,
                              highlightbackground=COLORS['border'],
                              highlightthickness=1, padx=16, pady=14)
        opts_frame.pack(fill='x', pady=(0, 10))

        ttk.Label(opts_frame, text="⚙️  Opzioni", style='Section.TLabel').pack(
            anchor='w', pady=(0, 10))

        # Griglia opzioni
        grid = tk.Frame(opts_frame, bg=COLORS['bg_card'])
        grid.pack(fill='x')

        # Riga 1: Classe, Font, Carta
        self._add_combo(grid, "Classe documento:", self.doc_class,
                        ['article', 'report', 'book', 'scrartcl', 'scrreprt'],
                        row=0, col=0)
        self._add_combo(grid, "Font (pt):", self.font_size,
                        ['10', '11', '12'],
                        row=0, col=2)
        self._add_combo(grid, "Formato carta:", self.paper_size,
                        ['a4paper', 'letterpaper', 'a5paper', 'b5paper'],
                        row=0, col=4)

        # Riga 2: Lingua, Motore
        self._add_combo(grid, "Lingua:", self.language,
                        ['italian', 'english', 'french', 'german', 'spanish',
                         'portuguese'],
                        row=1, col=0)
        self._add_combo(grid, "Motore LaTeX:", self.engine,
                        ['auto', 'pdflatex', 'lualatex', 'xelatex'],
                        row=1, col=2)

        # Checkbox
        checks_frame = tk.Frame(opts_frame, bg=COLORS['bg_card'])
        checks_frame.pack(fill='x', pady=(10, 0))

        ttk.Checkbutton(checks_frame, text="Mantieni .tex",
                        variable=self.keep_tex,
                        style='Card.TCheckbutton').pack(side='left', padx=(0, 16))
        ttk.Checkbutton(checks_frame, text="Mantieni immagini",
                        variable=self.keep_images,
                        style='Card.TCheckbutton').pack(side='left', padx=(0, 16))
        ttk.Checkbutton(checks_frame, text="Solo .tex (no PDF)",
                        variable=self.tex_only,
                        style='Card.TCheckbutton').pack(side='left')

        # === PULSANTI AZIONE ===
        action_frame = tk.Frame(main, bg=COLORS['bg'])
        action_frame.pack(fill='x', pady=(0, 10))

        self.convert_btn = self._make_button(
            action_frame, "▶  Converti", self._start_conversion,
            bg=COLORS['accent'], fg=COLORS['text_bright'],
            font_size=11, width=18, height=2,
        )
        self.convert_btn.pack(side='left')

        self.check_btn = self._make_button(
            action_frame, "🔍  Verifica Sistema", self._run_check,
            bg=COLORS['btn_secondary'], width=16,
        )
        self.check_btn.pack(side='left', padx=(10, 0))

        # Progressbar
        self.progress = ttk.Progressbar(
            action_frame, mode='indeterminate',
            style='Accent.Horizontal.TProgressbar',
            length=160,
        )
        self.progress.pack(side='right', padx=(10, 0))

        # === LOG / OUTPUT ===
        log_frame = tk.Frame(main, bg=COLORS['bg_card'], bd=0,
                             highlightbackground=COLORS['border'],
                             highlightthickness=1)
        log_frame.pack(fill='both', expand=True)

        log_header = tk.Frame(log_frame, bg=COLORS['bg_card'], padx=16, pady=(10))
        log_header.pack(fill='x')
        ttk.Label(log_header, text="📋  Log", style='Section.TLabel').pack(
            side='left')
        self._make_button(log_header, "Pulisci", self._clear_log,
                          bg=COLORS['btn_secondary'], width=7,
                          font_size=8).pack(side='right')

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            font=('JetBrains Mono', 9) if sys.platform != 'darwin' else ('Menlo', 9),
            bg=COLORS['input_bg'],
            fg=COLORS['text'],
            insertbackground=COLORS['text'],
            relief='flat',
            bd=0,
            padx=14,
            pady=8,
            wrap='word',
            state='disabled',
            height=10,
        )
        self.log_text.pack(fill='both', expand=True, padx=1, pady=(0, 1))

        # Tag colori per il log
        self.log_text.tag_configure('info', foreground=COLORS['text'])
        self.log_text.tag_configure('success', foreground=COLORS['success'])
        self.log_text.tag_configure('error', foreground=COLORS['error'])
        self.log_text.tag_configure('warning', foreground=COLORS['warning'])
        self.log_text.tag_configure('accent', foreground=COLORS['accent_light'])

        # === STATUS BAR ===
        self.status_var = tk.StringVar(value="Pronto")
        status_bar = tk.Frame(main, bg=COLORS['bg'])
        status_bar.pack(fill='x', pady=(6, 0))
        ttk.Label(status_bar, textvariable=self.status_var,
                  style='Status.TLabel').pack(side='left')
        ttk.Label(status_bar,
                  text=f"Formati: {', '.join(sorted(DocumentParser.SUPPORTED_EXTENSIONS))}",
                  style='Status.TLabel').pack(side='right')

    # =========================================================================
    # WIDGET HELPERS
    # =========================================================================

    def _make_button(self, parent, text, command, bg=None, fg=None,
                     font_size=9, width=None, height=1):
        """Crea un bottone stilizzato."""
        bg = bg or COLORS['accent']
        fg = fg or COLORS['text']
        btn = tk.Button(
            parent, text=text, command=command,
            font=(FONT_FAMILY, font_size),
            bg=bg, fg=fg,
            activebackground=COLORS['accent_hover'],
            activeforeground=COLORS['text_bright'],
            relief='flat', bd=0, padx=12, pady=4,
            cursor='hand2',
            height=height,
        )
        if width:
            btn.configure(width=width)

        # Hover effect
        hover_bg = COLORS['accent_hover'] if bg == COLORS['accent'] else COLORS['border']
        btn.bind('<Enter>', lambda e: btn.configure(bg=hover_bg))
        btn.bind('<Leave>', lambda e: btn.configure(bg=bg))

        return btn

    def _add_combo(self, parent, label, var, values, row, col):
        """Aggiunge una coppia label + combobox alla griglia."""
        ttk.Label(parent, text=label, style='Field.TLabel').grid(
            row=row, column=col, sticky='w', padx=(0, 4), pady=3)
        cb = ttk.Combobox(parent, textvariable=var, values=values,
                          state='readonly', width=14,
                          style='Card.TCombobox')
        cb.grid(row=row, column=col + 1, sticky='w', padx=(0, 16), pady=3)

    def _center_window(self):
        """Centra la finestra sullo schermo."""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'+{x}+{y}')

    # =========================================================================
    # AZIONI
    # =========================================================================

    def _enable_drag_and_drop(self):
        """Abilita drag&drop sul campo input, se supportato."""
        if not self.dragdrop_enabled:
            return

        self.input_entry.drop_target_register(DND_FILES)
        self.input_entry.dnd_bind('<<Drop>>', self._on_drop_file)

    def _on_drop_file(self, event):
        """Gestisce il rilascio di file drag&drop nel campo input."""
        try:
            files = self.root.tk.splitlist(event.data)
        except Exception:
            files = [event.data]

        if not files:
            return

        selected = files[0].strip()
        if selected.startswith('{') and selected.endswith('}'):
            selected = selected[1:-1]

        self._select_input_file(selected, source='drag&drop')

    def _select_input_file(self, path: str, source: str = 'sfoglia'):
        """Imposta input/output a partire dal file selezionato."""
        if not path:
            return

        path = os.path.abspath(path)
        if not os.path.exists(path):
            messagebox.showerror("Errore", f"File non trovato:\n{path}")
            return

        ext = Path(path).suffix.lower()
        if ext not in DocumentParser.SUPPORTED_EXTENSIONS:
            messagebox.showerror(
                "Formato non supportato",
                f"Formato '{ext}' non supportato.\n\n"
                f"Formati supportati:\n{', '.join(sorted(DocumentParser.SUPPORTED_EXTENSIONS))}",
            )
            return

        self.input_path.set(path)

        p = Path(path)
        if self.tex_only.get():
            out = p.with_suffix('.tex')
        else:
            out = p.with_suffix('.pdf')
        self.output_path.set(str(out))

        if source == 'drag&drop':
            self.status_var.set(f"Input impostato via drag&drop: {p.name}")
            self._log(f"📥 File caricato via drag&drop: {path}", 'accent')

    def _browse_input(self):
        """Apre il dialogo per selezionare il file di input."""
        path = filedialog.askopenfilename(
            title="Seleziona documento",
            filetypes=self.SUPPORTED_FILETYPES,
        )
        if path:
            self._select_input_file(path, source='sfoglia')

    def _browse_output(self):
        """Apre il dialogo per selezionare il file di output."""
        if self.tex_only.get():
            filetypes = [("LaTeX", "*.tex"), ("Tutti", "*.*")]
            default_ext = '.tex'
        else:
            filetypes = [("PDF", "*.pdf"), ("Tutti", "*.*")]
            default_ext = '.pdf'

        path = filedialog.asksaveasfilename(
            title="Salva come",
            filetypes=filetypes,
            defaultextension=default_ext,
        )
        if path:
            self.output_path.set(path)

    def _log(self, message: str, tag: str = 'info'):
        """Aggiunge un messaggio al log."""
        self.log_text.configure(state='normal')
        self.log_text.insert('end', message + '\n', tag)
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _clear_log(self):
        """Pulisce il log."""
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

    def _set_converting(self, state: bool):
        """Abilita/disabilita l'UI durante la conversione."""
        self.is_converting = state
        if state:
            self.convert_btn.configure(
                state='disabled', text="⏳ Conversione...",
                bg=COLORS['btn_secondary'])
            self.progress.start(15)
            self.status_var.set("Conversione in corso…")
        else:
            self.convert_btn.configure(
                state='normal', text="▶  Converti",
                bg=COLORS['accent'])
            self.progress.stop()

    def _start_conversion(self):
        """Avvia la conversione in un thread separato."""
        if self.is_converting:
            return

        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()

        if not input_path:
            messagebox.showwarning("Attenzione", "Seleziona un file di input.")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Errore", f"File non trovato:\n{input_path}")
            return

        ext = Path(input_path).suffix.lower()
        if ext not in DocumentParser.SUPPORTED_EXTENSIONS:
            messagebox.showerror(
                "Errore",
                f"Formato '{ext}' non supportato.\n\n"
                f"Formati supportati:\n{', '.join(sorted(DocumentParser.SUPPORTED_EXTENSIONS))}"
            )
            return

        self._clear_log()
        self._set_converting(True)
        self._log("═" * 56, 'accent')
        self._log("  WordToLaTeX Converter", 'accent')
        self._log("═" * 56, 'accent')
        self._log(f"  Input:  {input_path}")
        self._log(f"  Output: {output_path or '(auto)'}")
        self._log("")

        # Avvia conversione in thread
        thread = threading.Thread(
            target=self._convert_thread,
            args=(input_path, output_path),
            daemon=True,
        )
        thread.start()

    def _convert_thread(self, input_path: str, output_path: str):
        """Esegue la conversione in background."""
        try:
            engine = self.engine.get()
            if engine == 'auto':
                engine = None

            converter = WordToLatexConverter(
                document_class=self.doc_class.get(),
                font_size=int(self.font_size.get()),
                paper_size=self.paper_size.get(),
                language=self.language.get(),
                latex_engine=engine,
                keep_tex=self.keep_tex.get(),
                keep_images=self.keep_images.get(),
            )

            # Redirect print output
            old_stdout = sys.stdout
            sys.stdout = _LogCapture(self._log_threadsafe)

            try:
                if self.tex_only.get():
                    result = converter.convert_to_tex(
                        input_path,
                        output_path or None,
                    )
                else:
                    result = converter.convert(
                        input_path,
                        output_path or None,
                    )
            finally:
                sys.stdout = old_stdout

            self.root.after(0, self._conversion_done, result, None)

        except Exception as e:
            self.root.after(0, self._conversion_done, None, str(e))

    def _log_threadsafe(self, message: str, tag: str = 'info'):
        """Log thread-safe (schedula nel thread principale)."""
        self.root.after(0, self._log, message, tag)

    def _conversion_done(self, result: str, error: str):
        """Callback al completamento della conversione."""
        self._set_converting(False)

        if error:
            self._log(f"\n❌ ERRORE: {error}", 'error')
            self.status_var.set("Errore nella conversione")
            messagebox.showerror("Errore", f"Conversione fallita:\n\n{error}")
        else:
            self._log(f"\n✅ Conversione completata!", 'success')
            self._log(f"   Output: {result}", 'success')
            self.status_var.set(f"Completato: {Path(result).name}")

            # Chiedi se aprire il file
            if messagebox.askyesno(
                "Completato",
                f"Conversione completata!\n\n"
                f"File: {Path(result).name}\n\n"
                f"Vuoi aprire il file?",
            ):
                self._open_file(result)

    def _open_file(self, path: str):
        """Apre un file con l'applicazione predefinita."""
        import subprocess
        try:
            if sys.platform == 'linux':
                subprocess.Popen(['xdg-open', path])
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            elif sys.platform == 'win32':
                os.startfile(path)
        except Exception:
            pass

    def _run_check(self):
        """Esegue la verifica del sistema."""
        self._clear_log()
        self._log("═" * 56, 'accent')
        self._log("  Verifica Installazione", 'accent')
        self._log("═" * 56, 'accent')
        self._log("")

        # Python deps
        self._log("📦 Dipendenze Python:", 'accent')
        python_deps = {
            'python-docx': 'docx',
            'odfpy': 'odf',
            'python-pptx (opz.)': 'pptx',
        }
        for pkg_name, import_name in python_deps.items():
            try:
                __import__(import_name)
                self._log(f"  ✅ {pkg_name}: installato", 'success')
            except ImportError:
                self._log(f"  ❌ {pkg_name}: NON installato", 'error')

        self._log("")

        # LaTeX
        self._log("📄 Distribuzione LaTeX:", 'accent')
        status = check_latex_installation()
        if not status['available']:
            self._log("  ❌ Nessun motore LaTeX trovato!", 'error')
        else:
            for eng, info in status['engines'].items():
                if info['installed']:
                    self._log(f"  ✅ {eng}: {info['path']}", 'success')
                else:
                    self._log(f"  ⬜ {eng}: non trovato")

        if status.get('packages'):
            self._log("")
            self._log("📦 Pacchetti LaTeX:", 'accent')
            missing = []
            for pkg, info in status['packages'].items():
                if info['installed']:
                    self._log(f"  ✅ {pkg}", 'success')
                else:
                    self._log(f"  ❌ {pkg}", 'error')
                    missing.append(pkg)
            if missing:
                self._log(f"\n  ⚠️  Mancanti: {', '.join(missing)}", 'warning')

        self._log("")

        # LibreOffice
        import shutil
        lo_path = shutil.which('libreoffice') or shutil.which('soffice')
        self._log("📎 LibreOffice (per .doc/.rtf):", 'accent')
        if lo_path:
            self._log(f"  ✅ {lo_path}", 'success')
        else:
            self._log("  ⚠️  Non trovato (necessario solo per .doc e .rtf)", 'warning')

        self._log("")
        self._log("═" * 56, 'accent')

    def run(self):
        """Avvia l'applicazione."""
        self.root.mainloop()


class _LogCapture:
    """Cattura l'output di print() e lo invia al log della GUI."""

    def __init__(self, log_func):
        self.log_func = log_func
        self.buffer = ''

    def write(self, text):
        if text == '\n':
            if self.buffer:
                self.log_func(self.buffer)
                self.buffer = ''
            return
        self.buffer += text

    def flush(self):
        if self.buffer:
            self.log_func(self.buffer)
            self.buffer = ''


def main():
    """Entry point per la GUI."""
    app = WordToLatexGUI()
    app.run()


if __name__ == '__main__':
    main()
