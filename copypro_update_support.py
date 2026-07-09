from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import threading
import urllib.request
from pathlib import Path
from typing import Any

APP_DISPLAY_NAME = "CopyPro Tools"
RELEASE_ASSET_NAME = "CopyPro-Tools-Windows.zip"
GITHUB_OWNER = "CHANGE_ME"
GITHUB_REPOSITORY = "copypro-tools"

IS_WINDOWS = sys.platform.startswith("win")
IS_FROZEN = bool(getattr(sys, "frozen", False))

ROAMING_ROOT = Path(os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming")))
LOCAL_ROOT = Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local")))

DATA_DIR = ROAMING_ROOT / "CopyPro"
INSTALL_DIR = LOCAL_ROOT / APP_DISPLAY_NAME
UPDATER_DIR = LOCAL_ROOT / "CopyPro Updater"
UPDATE_LOG = DATA_DIR / "update.log"

DATA_DIR.mkdir(parents=True, exist_ok=True)
UPDATER_DIR.mkdir(parents=True, exist_ok=True)


def _log(message: str) -> None:
    try:
        with UPDATE_LOG.open("a", encoding="utf-8") as handle:
            handle.write(message.rstrip() + "\n")
    except Exception:
        pass


def _runtime_dir() -> Path:
    if IS_FROZEN:
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _resource_path(name: str) -> Path:
    bundle = Path(getattr(sys, "_MEIPASS", str(_runtime_dir())))
    return bundle / name


def load_update_config() -> dict[str, Any]:
    result: dict[str, Any] = {
        "github_owner": GITHUB_OWNER,
        "github_repository": GITHUB_REPOSITORY,
        "asset_name": RELEASE_ASSET_NAME,
    }
    for path in (_runtime_dir() / "update_config.json", _resource_path("update_config.json")):
        try:
            if path.exists():
                value = json.loads(path.read_text(encoding="utf-8-sig"))
                if isinstance(value, dict):
                    result.update(value)
                    break
        except Exception:
            continue
    return result


def current_version() -> str:
    for path in (_runtime_dir() / "version.json", _resource_path("version.json")):
        try:
            value = json.loads(path.read_text(encoding="utf-8-sig"))
            version = str(value.get("version", "")).strip().lstrip("v")
            if version:
                return version
        except Exception:
            continue
    return "0.0.0"


def _version_tuple(value: str) -> tuple[int, ...]:
    result = []
    for part in value.strip().lstrip("v").split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        result.append(int(digits or 0))
    return tuple(result or [0])


def is_newer_version(remote: str, local: str | None = None) -> bool:
    return _version_tuple(remote) > _version_tuple(local or current_version())


def create_desktop_shortcut(target_exe: Path) -> None:
    if not IS_WINDOWS:
        return
    target = str(target_exe).replace("'", "''")
    working = str(target_exe.parent).replace("'", "''")
    shortcut_name = f"{APP_DISPLAY_NAME}.lnk".replace("'", "''")
    ps_script = (
        "$w = New-Object -ComObject WScript.Shell;"
        "$desktop = $w.SpecialFolders.Item('Desktop');"
        f"$s = $w.CreateShortcut((Join-Path $desktop '{shortcut_name}'));"
        f"$s.TargetPath = '{target}';"
        f"$s.WorkingDirectory = '{working}';"
        f"$s.IconLocation = '{target},0';"
        "$s.Save();"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
        check=False,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def _copy_runtime_to(destination: Path) -> None:
    source = _runtime_dir()
    temporary = destination.with_name(destination.name + ".installing")
    if temporary.exists():
        shutil.rmtree(temporary, ignore_errors=True)
    shutil.copytree(source, temporary)
    if destination.exists():
        shutil.rmtree(destination, ignore_errors=True)
    temporary.replace(destination)


def ensure_external_updater() -> Path | None:
    if not IS_WINDOWS:
        return None
    updater_name = "CopyPro Updater.exe"
    source = None
    for candidate in (_runtime_dir() / updater_name, _resource_path(updater_name)):
        if candidate.exists():
            source = candidate
            break
    if source is None:
        _log("Updater executable was not found.")
        return None
    destination = UPDATER_DIR / updater_name
    try:
        if (
            not destination.exists()
            or destination.stat().st_size != source.stat().st_size
            or destination.stat().st_mtime < source.stat().st_mtime
        ):
            shutil.copy2(source, destination)
        return destination
    except Exception as exc:
        _log(f"Could not copy updater: {exc}")
        return None


def ensure_installed_in_appdata() -> bool:
    # A frozen onedir build launched from Desktop/Downloads/USB is copied to
    # Local AppData. The installed copy is launched and this process exits.
    if not (IS_WINDOWS and IS_FROZEN):
        return False
    current_dir = _runtime_dir()
    try:
        if current_dir.samefile(INSTALL_DIR):
            ensure_external_updater()
            return False
    except Exception:
        if str(current_dir).lower() == str(INSTALL_DIR).lower():
            ensure_external_updater()
            return False
    try:
        INSTALL_DIR.parent.mkdir(parents=True, exist_ok=True)
        _copy_runtime_to(INSTALL_DIR)
        installed_exe = INSTALL_DIR / Path(sys.executable).name
        bundled_updater = INSTALL_DIR / "CopyPro Updater.exe"
        if bundled_updater.exists():
            shutil.copy2(bundled_updater, UPDATER_DIR / bundled_updater.name)
        create_desktop_shortcut(installed_exe)
        subprocess.Popen([str(installed_exe)], cwd=str(INSTALL_DIR))
        return True
    except Exception as exc:
        _log(f"AppData installation failed: {exc}")
        return False


def fetch_latest_release(timeout: int = 12) -> dict[str, Any] | None:
    config = load_update_config()
    owner = str(config.get("github_owner", "")).strip()
    repository = str(config.get("github_repository", "")).strip()
    asset_name = str(config.get("asset_name", RELEASE_ASSET_NAME)).strip()
    if not owner or owner == "CHANGE_ME" or not repository:
        return None
    url = f"https://api.github.com/repos/{owner}/{repository}/releases/latest"
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"{APP_DISPLAY_NAME}/{current_version()}",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        release = json.loads(response.read().decode("utf-8"))
    version = str(release.get("tag_name", "")).strip().lstrip("v")
    if not version:
        return None
    download_url = None
    for asset in release.get("assets", []):
        if asset.get("name") == asset_name:
            download_url = asset.get("browser_download_url")
            break
    if not download_url:
        return None
    return {
        "version": version,
        "download_url": download_url,
        "notes": str(release.get("body") or "").strip(),
        "name": str(release.get("name") or release.get("tag_name") or version),
        "published_at": release.get("published_at"),
    }


def check_for_update_async(callback) -> None:
    def worker() -> None:
        try:
            release = fetch_latest_release()
            if release and is_newer_version(release["version"]):
                callback(release, None)
            else:
                callback(None, None)
        except Exception as exc:
            callback(None, exc)
    threading.Thread(target=worker, daemon=True).start()


def launch_updater(release: dict[str, Any]) -> None:
    """
    Start only the external updater with administrator permission.

    CopyPro Tools itself remains non-elevated so Windows Explorer drag-and-drop
    continues to work normally.
    """
    if not (IS_WINDOWS and IS_FROZEN):
        raise RuntimeError("Automatic updating is available in the compiled Windows build only.")

    updater = ensure_external_updater()
    if updater is None:
        raise FileNotFoundError(
            "CopyPro Updater.exe was not found. Build the release with the updater included."
        )

    arguments = [
        "--pid", str(os.getpid()),
        "--url", str(release["download_url"]),
        "--app-dir", str(INSTALL_DIR),
        "--exe-name", Path(sys.executable).name,
        "--version", str(release["version"]),
    ]

    # ShellExecuteW with the "runas" verb displays the Windows UAC prompt.
    import ctypes

    parameters = subprocess.list2cmdline(arguments)
    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        str(updater),
        parameters,
        str(UPDATER_DIR),
        1,
    )
    if int(result) <= 32:
        raise PermissionError(
            "Administrator permission was not granted, so the update could not start."
        )
