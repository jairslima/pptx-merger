"""
PPTX Merger by Jair Lima
Aplicativo para mesclar múltiplos arquivos PPTX em uma única apresentação.
"""

import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from merger import get_slide_count, merge_pptx, unique_output_path

APP_NAME = "PPTX Merger by Jair Lima"

# ──────────────────────────────────────────────────────────
# Detecção de ordem pelas partes no nome do arquivo
# ──────────────────────────────────────────────────────────

PART_PATTERNS = [
    r'parte\s*0*(\d+)',
    r'part\s*0*(\d+)',
    r'\bp0*(\d+)\b',
    r'vol(?:ume)?\s*0*(\d+)',
    r'cap(?:itulo)?\s*0*(\d+)',
    r'#\s*0*(\d+)',
]


def detect_part_number(filename: str):
    """Retorna o número de parte encontrado no nome do arquivo, ou None."""
    stem = os.path.splitext(os.path.basename(filename))[0].lower()
    for pattern in PART_PATTERNS:
        m = re.search(pattern, stem)
        if m:
            return int(m.group(1))
    return None


def auto_detect_pairs(file_paths: list) -> list:
    """
    Tenta detectar e ordenar automaticamente arquivos por número de parte.
    Retorna a lista reordenada; se não detectar padrão retorna lista original.
    """
    numbered = [(detect_part_number(f), f) for f in file_paths]
    detected = [(n, f) for n, f in numbered if n is not None]
    if len(detected) >= 2:
        detected.sort(key=lambda x: x[0])
        return [f for _, f in detected]
    return file_paths


# ──────────────────────────────────────────────────────────
# Janela principal
# ──────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.resizable(True, True)
        self.minsize(700, 550)

        # Estado interno
        self._folder_path = tk.StringVar()
        self._output_name = tk.StringVar(value="apresentacao_completa.pptx")
        self._all_files: list[dict] = []   # {path, name, slides, checked, order_var}
        self._merge_order: list[dict] = [] # subconjunto selecionado, em ordem
        self._busy = False

        self._build_ui()
        self._set_status("Selecione uma pasta para começar.")

    # ── Construção da UI ──────────────────────────────────

    def _build_ui(self):
        self.configure(bg="#1e1e2e")

        # Cabeçalho
        header = tk.Frame(self, bg="#313244", pady=10)
        header.pack(fill=tk.X)
        tk.Label(
            header, text=APP_NAME,
            font=("Segoe UI", 15, "bold"),
            fg="#cdd6f4", bg="#313244"
        ).pack()

        # Linha de seleção de pasta
        folder_frame = tk.Frame(self, bg="#1e1e2e", pady=6, padx=12)
        folder_frame.pack(fill=tk.X)
        tk.Label(folder_frame, text="Pasta:", fg="#cdd6f4", bg="#1e1e2e",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=(0, 6))
        tk.Entry(folder_frame, textvariable=self._folder_path, width=55,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 relief=tk.FLAT, font=("Segoe UI", 10)).grid(row=0, column=1, sticky="ew", padx=4)
        tk.Button(folder_frame, text="Selecionar Pasta",
                  command=self._browse_folder,
                  bg="#89b4fa", fg="#1e1e2e", relief=tk.FLAT,
                  font=("Segoe UI", 10, "bold"), cursor="hand2",
                  padx=8).grid(row=0, column=2, padx=(6, 0))
        folder_frame.columnconfigure(1, weight=1)

        # Área principal: lista de arquivos + painel de ordem
        main_area = tk.Frame(self, bg="#1e1e2e", padx=12)
        main_area.pack(fill=tk.BOTH, expand=True, pady=4)
        main_area.columnconfigure(0, weight=3)
        main_area.columnconfigure(1, weight=2)

        # Painel esquerdo: arquivos encontrados
        left = tk.LabelFrame(main_area, text=" Arquivos encontrados ",
                             bg="#1e1e2e", fg="#a6adc8",
                             font=("Segoe UI", 9), relief=tk.GROOVE)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=4)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        self._file_frame = tk.Frame(left, bg="#1e1e2e")
        self._file_frame.pack(fill=tk.BOTH, expand=True)

        # Canvas com scrollbar para a lista de arquivos
        self._canvas = tk.Canvas(self._file_frame, bg="#1e1e2e",
                                 highlightthickness=0)
        scrollbar = ttk.Scrollbar(self._file_frame, orient="vertical",
                                  command=self._canvas.yview)
        self._scrollable = tk.Frame(self._canvas, bg="#1e1e2e")
        self._scrollable.bind(
            "<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all"))
        )
        self._canvas.create_window((0, 0), window=self._scrollable, anchor="nw")
        self._canvas.configure(yscrollcommand=scrollbar.set)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Botão selecionar todos / nenhum
        sel_row = tk.Frame(left, bg="#1e1e2e", pady=3)
        sel_row.pack(fill=tk.X)
        tk.Button(sel_row, text="Marcar todos", command=self._select_all,
                  bg="#45475a", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2").pack(side=tk.LEFT, padx=4)
        tk.Button(sel_row, text="Desmarcar todos", command=self._deselect_all,
                  bg="#45475a", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2").pack(side=tk.LEFT)

        # Painel direito: ordem de mesclagem
        right = tk.LabelFrame(main_area, text=" Ordem de mesclagem ",
                              bg="#1e1e2e", fg="#a6adc8",
                              font=("Segoe UI", 9), relief=tk.GROOVE)
        right.grid(row=0, column=1, sticky="nsew", pady=4)
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        self._order_list = tk.Listbox(
            right, bg="#313244", fg="#cdd6f4",
            selectbackground="#89b4fa", selectforeground="#1e1e2e",
            font=("Segoe UI", 9), relief=tk.FLAT,
            activestyle="none", height=12
        )
        self._order_list.pack(fill=tk.BOTH, expand=True, padx=4, pady=(4, 2))

        btn_row = tk.Frame(right, bg="#1e1e2e", pady=3)
        btn_row.pack(fill=tk.X)
        tk.Button(btn_row, text="▲ Subir", command=self._move_up,
                  bg="#45475a", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2").pack(side=tk.LEFT, padx=4)
        tk.Button(btn_row, text="▼ Descer", command=self._move_down,
                  bg="#45475a", fg="#cdd6f4", relief=tk.FLAT,
                  font=("Segoe UI", 9), cursor="hand2").pack(side=tk.LEFT)

        # Linha de saída + botão mesclar
        bottom = tk.Frame(self, bg="#1e1e2e", padx=12, pady=6)
        bottom.pack(fill=tk.X)
        tk.Label(bottom, text="Arquivo de saída:", fg="#cdd6f4", bg="#1e1e2e",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=(0, 6))
        tk.Entry(bottom, textvariable=self._output_name, width=38,
                 bg="#313244", fg="#cdd6f4", insertbackground="#cdd6f4",
                 relief=tk.FLAT, font=("Segoe UI", 10)).grid(row=0, column=1, sticky="ew")
        self._merge_btn = tk.Button(
            bottom, text="Mesclar Arquivos",
            command=self._start_merge,
            bg="#a6e3a1", fg="#1e1e2e", relief=tk.FLAT,
            font=("Segoe UI", 11, "bold"), cursor="hand2", padx=12
        )
        self._merge_btn.grid(row=0, column=2, padx=(10, 0))
        bottom.columnconfigure(1, weight=1)

        # Barra de progresso
        self._progress = ttk.Progressbar(self, mode="determinate",
                                         maximum=100, value=0)
        self._progress.pack(fill=tk.X, padx=12, pady=(0, 2))

        # Status bar
        self._status_var = tk.StringVar()
        status_bar = tk.Frame(self, bg="#181825", pady=4)
        status_bar.pack(fill=tk.X)
        tk.Label(status_bar, textvariable=self._status_var,
                 fg="#a6adc8", bg="#181825",
                 font=("Segoe UI", 9), anchor="w", padx=8).pack(fill=tk.X)

    # ── Navegação de pasta ─────────────────────────────────

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Selecionar pasta com arquivos PPTX")
        if folder:
            self._folder_path.set(folder)
            self._scan_folder(folder)

    def _scan_folder(self, folder: str):
        """Varre a pasta em busca de .pptx e popula a lista."""
        self._all_files.clear()
        self._merge_order.clear()
        self._set_status("Varrendo pasta...")

        paths = [
            os.path.join(folder, f)
            for f in sorted(os.listdir(folder))
            if f.lower().endswith(".pptx") and not f.startswith("~")
        ]

        if not paths:
            self._set_status("Nenhum arquivo .pptx encontrado nesta pasta.")
            self._rebuild_file_list()
            self._rebuild_order_list()
            return

        # Detectar ordem automática
        ordered = auto_detect_pairs(paths)
        auto_selected = set(ordered) if len(ordered) >= 2 else set()

        # Contar slides de cada arquivo
        for path in paths:
            count = get_slide_count(path)
            name = os.path.basename(path)
            part_n = detect_part_number(path)

            entry = {
                "path": path,
                "name": name,
                "slides": count,
                "checked": tk.BooleanVar(value=(path in auto_selected)),
                "part_n": part_n,
            }
            self._all_files.append(entry)

        # Montar ordem de mesclagem inicial direto da lista auto-detectada
        # (não usar _rebuild_merge_order aqui: ela usaria a ordem alfabética de _all_files)
        if auto_selected:
            for path in ordered:          # 'ordered' já está na ordem certa (parte 1, 2...)
                entry = next((e for e in self._all_files if e["path"] == path), None)
                if entry:
                    self._merge_order.append(entry)
        self._rebuild_file_list()
        self._rebuild_order_list()
        self._set_status(
            f"{len(paths)} arquivo(s) encontrado(s). "
            + (f"Detecção automática: {len(auto_selected)} arquivo(s) selecionado(s)."
               if auto_selected else "Selecione os arquivos a mesclar.")
        )

    # ── Construção das listas na UI ───────────────────────

    def _rebuild_file_list(self):
        """Reconstrói os widgets de checkbox na lista de arquivos."""
        for widget in self._scrollable.winfo_children():
            widget.destroy()

        for entry in self._all_files:
            row = tk.Frame(self._scrollable, bg="#1e1e2e", pady=1)
            row.pack(fill=tk.X)

            cb = tk.Checkbutton(
                row, variable=entry["checked"],
                bg="#1e1e2e", fg="#cdd6f4",
                selectcolor="#313244", activebackground="#1e1e2e",
                command=self._on_check_changed,
                cursor="hand2"
            )
            cb.pack(side=tk.LEFT)

            info = f"{entry['name']}  [{entry['slides']} slide(s)]"
            if entry["part_n"] is not None:
                info += f"  [Parte {entry['part_n']}]"

            tk.Label(row, text=info, fg="#cdd6f4", bg="#1e1e2e",
                     font=("Segoe UI", 9), anchor="w").pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _rebuild_merge_order(self):
        """Reconstrói _merge_order a partir dos arquivos marcados."""
        checked = [e for e in self._all_files if e["checked"].get()]

        # Manter ordem anterior para itens já presentes
        prev_paths = [e["path"] for e in self._merge_order]
        new_order = []
        for path in prev_paths:
            match = next((e for e in checked if e["path"] == path), None)
            if match:
                new_order.append(match)
        for entry in checked:
            if entry not in new_order:
                new_order.append(entry)
        self._merge_order = new_order

        # Se houve detecção automática e a ordem está vazia, usar ordem detectada
        if not self._merge_order and checked:
            detected = sorted(
                [e for e in checked if e["part_n"] is not None],
                key=lambda e: e["part_n"]
            )
            others = [e for e in checked if e["part_n"] is None]
            self._merge_order = detected + others

    def _rebuild_order_list(self):
        """Atualiza o Listbox da ordem de mesclagem."""
        self._order_list.delete(0, tk.END)
        total = sum(e["slides"] for e in self._merge_order)
        for i, entry in enumerate(self._merge_order, 1):
            self._order_list.insert(tk.END, f"{i}. {entry['name']}  [{entry['slides']} sl]")
        if self._merge_order:
            self._order_list.insert(tk.END, "")
            self._order_list.insert(tk.END, f"   Total: {total} slide(s)")

    # ── Eventos ───────────────────────────────────────────

    def _on_check_changed(self):
        self._rebuild_merge_order()
        self._rebuild_order_list()

    def _select_all(self):
        for e in self._all_files:
            e["checked"].set(True)
        self._on_check_changed()

    def _deselect_all(self):
        for e in self._all_files:
            e["checked"].set(False)
        self._merge_order.clear()
        self._rebuild_order_list()

    def _move_up(self):
        sel = self._order_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx <= 0 or idx >= len(self._merge_order):
            return
        self._merge_order[idx], self._merge_order[idx - 1] = (
            self._merge_order[idx - 1], self._merge_order[idx]
        )
        self._rebuild_order_list()
        self._order_list.selection_set(idx - 1)

    def _move_down(self):
        sel = self._order_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self._merge_order) - 1:
            return
        self._merge_order[idx], self._merge_order[idx + 1] = (
            self._merge_order[idx + 1], self._merge_order[idx]
        )
        self._rebuild_order_list()
        self._order_list.selection_set(idx + 1)

    # ── Mesclagem ─────────────────────────────────────────

    def _start_merge(self):
        if self._busy:
            return

        if not self._merge_order:
            messagebox.showwarning(APP_NAME,
                                   "Nenhum arquivo selecionado para mesclar.")
            return
        if len(self._merge_order) < 2:
            messagebox.showwarning(APP_NAME,
                                   "Selecione pelo menos 2 arquivos para mesclar.")
            return

        folder = self._folder_path.get() or os.path.dirname(
            self._merge_order[0]["path"]
        )
        out_name = self._output_name.get().strip() or "apresentacao_completa.pptx"
        if not out_name.lower().endswith(".pptx"):
            out_name += ".pptx"

        out_path = unique_output_path(os.path.join(folder, out_name))

        self._busy = True
        self._merge_btn.configure(state=tk.DISABLED, bg="#585b70")
        self._progress["value"] = 0
        self._set_status("Mesclando arquivos...")

        files = [e["path"] for e in self._merge_order]

        def run():
            total_extra = sum(e["slides"] for e in self._merge_order[1:])

            def on_progress(done, total):
                pct = int(done / total * 100) if total else 100
                self.after(0, lambda: self._progress.configure(value=pct))

            try:
                n = merge_pptx(files, out_path, progress_callback=on_progress)
                self.after(0, lambda: self._merge_done(out_path, n))
            except Exception as exc:
                self.after(0, lambda: self._merge_error(str(exc)))

        threading.Thread(target=run, daemon=True).start()

    def _merge_done(self, out_path: str, total_slides: int):
        self._busy = False
        self._merge_btn.configure(state=tk.NORMAL, bg="#a6e3a1")
        self._progress["value"] = 100
        self._set_status(
            f"Concluído. {total_slides} slide(s) em: {os.path.basename(out_path)}"
        )
        answer = messagebox.askyesno(
            APP_NAME,
            f"Mesclagem concluída!\n\n"
            f"{total_slides} slide(s) salvos em:\n{out_path}\n\n"
            f"Abrir a pasta de saída?"
        )
        if answer:
            _open_folder(os.path.dirname(out_path))

    def _merge_error(self, msg: str):
        self._busy = False
        self._merge_btn.configure(state=tk.NORMAL, bg="#f38ba8")
        self._progress["value"] = 0
        self._set_status(f"Erro: {msg}")
        messagebox.showerror(APP_NAME, f"Erro durante a mesclagem:\n\n{msg}")

    # ── Utilitários ───────────────────────────────────────

    def _set_status(self, msg: str):
        self._status_var.set(msg)


# ──────────────────────────────────────────────────────────
# Utilitários externos
# ──────────────────────────────────────────────────────────

def _open_folder(path: str):
    """Abre o explorador de arquivos na pasta especificada."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ──────────────────────────────────────────────────────────
# Ponto de entrada
# ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
