# -*- coding: utf-8 -*-
"""Inicializador empacotável do Caderno Faculdade.

Mantém a aplicação principal existente e melhora a seleção da pasta do Drive:
o seletor começa na pasta sincronizada detectada, quando possível.
"""
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import caderno_faculdade as app
from auto_updater import AutoUpdater
from build_info import BUILD_NUMBER

# Paleta refinada: contraste mais consistente e tons esmeralda menos saturados.
app.THEMES["light"].update({
    "bg": "#F3F6F4", "card": "#FFFFFF", "fg": "#18231D", "muted": "#6F7F75",
    "line": "#E0E8E2", "field": "#F7F9F7", "accent": "#238A4A",
    "accent_hover": "#1D773F", "accent_fg": "#FFFFFF", "accent_text": "#1E7A40",
    "soft": "#E5F3E9", "soft_fg": "#143B22", "dark": "#14231A",
    "dark_fg": "#FFFFFF", "dark_muted": "#A7B9AC", "dark_hover": "#20372A",
    "bar": "#D9E7DC", "danger": "#C74646", "warn": "#D99A22", "select": "#D8EBDD",
    "side": "#101A14", "side2": "#1B2A20", "side_hover": "#26392C",
    "side_fg": "#EDF5EF", "side_muted": "#91A397", "side_line": "#2C4031",
})
app.THEMES["dark"].update({
    "bg": "#0C1510", "card": "#142019", "fg": "#EAF2EC", "muted": "#91A397",
    "line": "#26372C", "field": "#101A14", "accent": "#43D477",
    "accent_hover": "#32C468", "accent_fg": "#07120B", "accent_text": "#5AE48A",
    "soft": "#1B3524", "soft_fg": "#E7F6EB", "dark": "#08110B",
    "dark_fg": "#FFFFFF", "dark_muted": "#879B8D", "dark_hover": "#17281D",
    "bar": "#2A3E31", "danger": "#FF7777", "warn": "#F2BD4A", "select": "#214B30",
    "side": "#080F0B", "side2": "#142019", "side_hover": "#1B2D21",
    "side_fg": "#EAF2EC", "side_muted": "#84988B", "side_line": "#24362A",
})


def drive_candidates(current):
    home = Path.home()
    current_path = Path(current)
    candidates = [
        home / "Google Drive",
        home / "Meu Drive",
        home / "My Drive",
        home / "Google Drive" / "Meu Drive",
        home / "Google Drive" / "My Drive",
        home / "Google Drive" / "My Drive" / "Caderno Faculdade",
        home / "Google Drive" / "Meu Drive" / "Caderno Faculdade",
    ]
    # Google Drive para computador pode montar o Meu Drive como uma unidade,
    # por exemplo G:\Meu Drive ou G:\My Drive.
    if os.name == "nt":
        for letter in "DEFGHIJKLMNOPQRSTUVWXYZ":
            root = Path(f"{letter}:\\")
            candidates.extend([
                root / "Meu Drive",
                root / "My Drive",
                root / "Google Drive",
                root / "Meu Drive" / "Caderno Faculdade",
                root / "My Drive" / "Caderno Faculdade",
            ])
    seen = set()
    result = []
    for item in candidates:
        try:
            key = os.path.normcase(os.path.abspath(str(item)))
            if key not in seen and item.exists() and item.is_dir():
                seen.add(key)
                result.append(item)
        except OSError:
            continue
    return result


def choose_root(self):
    if not self.leave_ok():
        return
    candidates = drive_candidates(self.root_dir)
    initial = candidates[0] if candidates else Path.home()
    chosen = filedialog.askdirectory(
        parent=self.root,
        title="Escolha uma pasta para salvar suas matérias",
        initialdir=str(initial),
        mustexist=False,
    )
    if not chosen:
        return
    self.root_dir = Path(chosen).expanduser()
    try:
        self.root_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        messagebox.showerror(
            "Não foi possível usar esta pasta",
            f"O Windows não conseguiu acessar a pasta selecionada.\\n\\n{exc}",
            parent=self.root,
        )
        return
    self.config["root_dir"] = str(self.root_dir)
    app.write_config(self.config)
    self.cancel_autosave()
    self.editing = self.dirty = self.meta_dirty = False
    self.current_file = None
    self.show_empty()
    self.update_sync_label()
    self.set_status("idle", "Pasta atualizada")
    self.rescan()
    self.highlight_current()


# Substitui o seletor de pasta e adiciona a verificação de atualizações no rodapé.
app.NotebookApp.choose_root = choose_root

_original_build_sidebar = app.NotebookApp.build_sidebar


def build_sidebar(self):
    _original_build_sidebar(self)
    bottom = self.sidebar.winfo_children()[-1]
    app.Btn(
        bottom, self, "Verificar atualizações",
        lambda: getattr(self, "updater", None) and self.updater.check_now(manual=True),
        variant="nav", on="side", size=10, bold=False, height=40,
        anchor="w", expand=True, icon="↻",
    ).pack(fill="x", padx=14, pady=2)


app.NotebookApp.build_sidebar = build_sidebar


def main():
    # A DPI awareness precisa ser ativada antes de criar a janela Tk.
    app._enable_windows_dpi_awareness()
    root = tk.Tk()
    notebook = app.NotebookApp(root)
    notebook.updater = AutoUpdater(root, notebook, BUILD_NUMBER)
    notebook.updater.start()
    root.mainloop()


if __name__ == "__main__":
    main()
