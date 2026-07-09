#!/usr/bin/env python3
from __future__ import annotations

import argparse
import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

BG = "#0F0F13"
SURFACE2 = "#23232E"
ACCENT = "#5B6BFF"
TEXT = "#F0F0F8"
MUTED = "#7878A0"


def is_running_as_admin() -> bool:
    if not sys.platform.startswith("win"):
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    """Relaunch this updater through the Windows UAC prompt."""
    if not sys.platform.startswith("win"):
        return False

    if getattr(sys, "frozen", False):
        executable = str(Path(sys.executable).resolve())
        arguments = sys.argv[1:]
    else:
        executable = str(Path(sys.executable).resolve())
        arguments = [str(Path(__file__).resolve()), *sys.argv[1:]]

    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        executable,
        subprocess.list2cmdline(arguments),
        str(Path(executable).parent),
        1,
    )
    return int(result) > 32


def launch_app_normally(executable: Path) -> None:
    """
    Ask the normal Windows Explorer process to start the app.

    The updater may be elevated, but the main CopyPro app should not stay
    elevated because Windows blocks drag-and-drop from normal Explorer windows
    into an administrator process.
    """
    if sys.platform.startswith("win"):
        try:
            subprocess.Popen(
                ["explorer.exe", str(executable)],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            return
        except Exception:
            pass

    subprocess.Popen([str(executable)], cwd=str(executable.parent))


def wait_for_process(pid: int, timeout: int = 45) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if str(pid) not in result.stdout:
            return
        time.sleep(0.4)


def find_payload_root(extract_dir: Path, exe_name: str) -> Path:
    if (extract_dir / exe_name).exists():
        return extract_dir
    matches = list(extract_dir.rglob(exe_name))
    if len(matches) == 1:
        return matches[0].parent
    raise FileNotFoundError(f"{exe_name} was not found in the update package.")


class UpdaterWindow(tk.Tk):
    def __init__(self, args):
        super().__init__()
        self.args = args
        self.title("CopyPro Tools update")
        self.geometry("470x180")
        self.resizable(False, False)
        self.configure(bg=BG)
        tk.Label(self, text=f"Updating CopyPro Tools to {args.version}", bg=BG, fg=TEXT,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=20, pady=(20, 5))
        self.status = tk.Label(self, text="Preparing update…", bg=BG, fg=MUTED,
                               font=("Segoe UI", 9))
        self.status.pack(anchor="w", padx=20, pady=(0, 12))
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Horizontal.TProgressbar", troughcolor=SURFACE2,
                        background=ACCENT, borderwidth=0, thickness=9)
        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=20)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        self.after(200, self._start)

    def set_status(self, text: str, value: int | None = None) -> None:
        self.after(0, lambda: self.status.configure(text=text))
        if value is not None:
            self.after(0, lambda: self.progress.configure(value=value))

    def _start(self):
        threading.Thread(target=self._run, daemon=True).start()

    def _download(self, destination: Path) -> None:
        request = urllib.request.Request(self.args.url, headers={"User-Agent": "CopyPro-Tools-Updater"})
        with urllib.request.urlopen(request, timeout=60) as response:
            total = int(response.headers.get("Content-Length", "0") or "0")
            received = 0
            with destination.open("wb") as handle:
                while True:
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    handle.write(chunk)
                    received += len(chunk)
                    if total:
                        value = max(2, min(55, int(received / total * 53) + 2))
                        self.set_status("Downloading update…", value)

    def _run(self):
        app_dir = Path(self.args.app_dir)
        executable = app_dir / self.args.exe_name
        backup_dir = app_dir.with_name(app_dir.name + ".backup")
        new_dir = app_dir.with_name(app_dir.name + ".new")
        try:
            with tempfile.TemporaryDirectory(prefix="copypro_update_") as temp_value:
                temp = Path(temp_value)
                archive = temp / "update.zip"
                extracted = temp / "extracted"
                self.set_status("Downloading update…", 2)
                self._download(archive)
                self.set_status("Extracting files…", 60)
                extracted.mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(archive, "r") as package:
                    package.extractall(extracted)
                payload = find_payload_root(extracted, self.args.exe_name)
                self.set_status("Waiting for CopyPro Tools to close…", 68)
                wait_for_process(self.args.pid)
                if new_dir.exists():
                    shutil.rmtree(new_dir, ignore_errors=True)
                shutil.copytree(payload, new_dir)
                if not (new_dir / self.args.exe_name).exists():
                    raise FileNotFoundError("The new executable is missing.")
                self.set_status("Installing update…", 82)
                if backup_dir.exists():
                    shutil.rmtree(backup_dir, ignore_errors=True)
                if app_dir.exists():
                    app_dir.replace(backup_dir)
                new_dir.replace(app_dir)
                self.set_status("Starting CopyPro Tools…", 97)
                launch_app_normally(executable)
                shutil.rmtree(backup_dir, ignore_errors=True)
                self.set_status("Update complete.", 100)
                self.after(900, self.destroy)
        except Exception as exc:
            try:
                if not app_dir.exists() and backup_dir.exists():
                    backup_dir.replace(app_dir)
            except Exception:
                pass
            self.after(0, lambda: messagebox.showerror(
                "Update failed", f"CopyPro Tools could not be updated.\n\n{exc}", parent=self))
            self.set_status("Update failed.", 0)
            self.after(500, lambda: self.protocol("WM_DELETE_WINDOW", self.destroy))


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--url", required=True)
    parser.add_argument("--app-dir", required=True)
    parser.add_argument("--exe-name", required=True)
    parser.add_argument("--version", required=True)
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if sys.platform.startswith("win") and not is_running_as_admin():
        try:
            if relaunch_as_admin():
                raise SystemExit(0)
        except Exception:
            pass

        ctypes.windll.user32.MessageBoxW(
            None,
            "Administrator permission is required to install this CopyPro Tools update.",
            "CopyPro Tools update",
            0x10,
        )
        raise SystemExit(1)

    UpdaterWindow(args).mainloop()
