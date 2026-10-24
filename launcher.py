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


# Substitui apenas o seletor de pasta; todo o restante continua vindo do app original.
app.NotebookApp.choose_root = choose_root


if __name__ == "__main__":
    app.main()
