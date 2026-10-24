# -*- coding: utf-8 -*-
"""Verificação e atualização automática da versão portátil do Caderno Faculdade."""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
import zipfile
from pathlib import Path
from tkinter import messagebox


REPO_API = "https://api.github.com/repos/CaioBarretor/Caderno-Digital/releases/latest"
APP_EXE_NAME = "Caderno Faculdade.exe"


class AutoUpdater:
    def __init__(self, root, app, build_number):
        self.root = root
        self.app = app
        self.build_number = int(build_number)
        self._checking = False
        self._lock = threading.Lock()
        self._pending = None

    def start(self):
        self.root.after(3500, lambda: self.check_now(manual=False))
        self.root.after(6 * 60 * 60 * 1000, self._periodic_check)

    def _periodic_check(self):
        self.check_now(manual=False)
        try:
            self.root.after(6 * 60 * 60 * 1000, self._periodic_check)
        except Exception:
            pass

    def check_now(self, manual=True):
        if self._pending:
            self.root.after(0, self._install_ready)
            return
        with self._lock:
            if self._checking:
                if manual:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Atualizações", "Já estou verificando se há uma versão nova.", parent=self.root))
                return
            self._checking = True
        threading.Thread(target=self._check_worker, args=(manual,), daemon=True).start()

    def _check_worker(self, manual):
        try:
            req = urllib.request.Request(
                REPO_API,
                headers={"Accept": "application/vnd.github+json", "User-Agent": "Caderno-Faculdade-Updater"},
            )
            with urllib.request.urlopen(req, timeout=15) as response:
                release = json.loads(response.read().decode("utf-8"))
            tag = str(release.get("tag_name", ""))
            match = re.search(r"(?:portable-|build[- ]?)(\d+)", tag, re.IGNORECASE)
            if not match:
                match = re.search(r"(\d+)", str(release.get("name", "")))
            latest = int(match.group(1)) if match else 0
            asset = next((a for a in release.get("assets", [])
                          if a.get("name") == "Caderno-Faculdade-Windows.zip"), None)
            if latest > self.build_number and asset and asset.get("browser_download_url"):
                self.root.after(0, lambda: self._update_found(release, asset))
            elif manual:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Atualizações", "Você já está usando a versão mais recente do Caderno Faculdade.",
                    parent=self.root))
        except Exception as exc:
            if manual:
                msg = str(exc)
                self.root.after(0, lambda msg=msg: messagebox.showwarning(
                    "Não foi possível verificar", "Não consegui consultar as atualizações agora.\n\n"
                    + msg[:500], parent=self.root))
        finally:
            with self._lock:
                self._checking = False

    def _update_found(self, release, asset):
        version = release.get("name") or release.get("tag_name") or "nova versão"
        messagebox.showinfo(
            "Atualização disponível",
            f"Foi encontrada a {version}. O download começará agora. "
            "Quando estiver pronto, o aplicativo será atualizado e reiniciado automaticamente.",
            parent=self.root,
        )
        threading.Thread(target=self._download_worker, args=(release, asset), daemon=True).start()

    def _download_worker(self, release, asset):
        workdir = None
        try:
            workdir = Path(tempfile.mkdtemp(prefix="caderno-update-"))
            archive = workdir / "update.zip"
            req = urllib.request.Request(
                asset["browser_download_url"],
                headers={"User-Agent": "Caderno-Faculdade-Updater"},
            )
            with urllib.request.urlopen(req, timeout=60) as response, archive.open("wb") as output:
                shutil.copyfileobj(response, output)
            extract_dir = workdir / "extracted"
            extract_dir.mkdir()
            with zipfile.ZipFile(archive, "r") as zf:
                member = next((m for m in zf.infolist()
                               if Path(m.filename.replace("\\", "/")).name.lower() == APP_EXE_NAME.lower()), None)
                if member is None:
                    raise RuntimeError("O pacote baixado não contém o executável esperado.")
                with zf.open(member) as source, (extract_dir / APP_EXE_NAME).open("wb") as output:
                    shutil.copyfileobj(source, output)
            self._pending = (workdir, extract_dir / APP_EXE_NAME, release)
            self.root.after(0, self._install_ready)
        except Exception as exc:
            if workdir:
                shutil.rmtree(workdir, ignore_errors=True)
            msg = str(exc)
            self.root.after(0, lambda msg=msg: messagebox.showerror(
                "Falha na atualização", "Não foi possível baixar ou preparar a atualização. "
                "O aplicativo atual continua funcionando.\n\n" + msg[:700], parent=self.root))

    def _install_ready(self):
        if not self._pending:
            return
        if not getattr(sys, "frozen", False):
            messagebox.showinfo(
                "Atualização baixada",
                "A nova versão foi baixada, mas a atualização automática só funciona na versão instalada "
                "pelo executável portátil do Windows.", parent=self.root)
            return
        workdir, new_exe, release = self._pending
        if not self.app.leave_ok():
            messagebox.showinfo(
                "Atualização pronta",
                "O download terminou. Salve ou resolva as alterações abertas e clique em "
                "'Verificar atualizações' para concluir a instalação.", parent=self.root)
            return
        target = Path(sys.executable).resolve()
        if target.name.lower() != APP_EXE_NAME.lower():
            messagebox.showwarning(
                "Atualização não aplicada",
                "Não consegui identificar o executável atual com segurança. Baixe a versão nova manualmente "
                "pela página de versões do GitHub.", parent=self.root)
            return
        script = workdir / "install-update.ps1"
        script.write_text(
            "$ErrorActionPreference = 'Stop'\n"
            f"$processId = {os.getpid()}\n"
            f"$source = '{str(new_exe).replace(chr(39), chr(39)+chr(39))}'\n"
            f"$target = '{str(target).replace(chr(39), chr(39)+chr(39))}'\n"
            "try {\n"
            "  try { Wait-Process -Id $processId -Timeout 120 -ErrorAction SilentlyContinue } catch {}\n"
            "  Start-Sleep -Milliseconds 700\n"
            "  Copy-Item -LiteralPath $source -Destination $target -Force\n"
            "  Start-Process -FilePath $target\n"
            "} catch {\n"
            "  Add-Type -AssemblyName PresentationFramework\n"
            "  [System.Windows.MessageBox]::Show('A atualização não pôde substituir o aplicativo. "
            "Feche-o e tente novamente, ou baixe a versão no GitHub.','Caderno Faculdade') | Out-Null\n"
            "  if (Test-Path -LiteralPath $target) { Start-Process -FilePath $target }\n"
            "}\n"
            "Start-Sleep -Seconds 3\n"
            f"Remove-Item -LiteralPath '{str(workdir).replace(chr(39), chr(39)+chr(39))}' -Recurse -Force -ErrorAction SilentlyContinue\n",
            encoding="utf-8",
        )
        try:
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "DETACHED_PROCESS", 0)
            subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script)],
                cwd=str(workdir), creationflags=flags, close_fds=True,
            )
        except Exception as exc:
            messagebox.showerror("Falha na atualização", str(exc), parent=self.root)
            return
        self._pending = None
        self.root.destroy()
