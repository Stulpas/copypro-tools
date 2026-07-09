import tkinter as tk, time
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
import os, sys, threading, subprocess, io, math, tempfile, shutil, unicodedata, difflib, json, csv, zipfile, webbrowser, base64, mimetypes, html, gc, ctypes, queue
from collections import OrderedDict
from pathlib import Path
from copypro_update_support import (
    current_version,
    ensure_installed_in_appdata,
    check_for_update_async,
    launch_updater,
)

# ── Dependency bootstrap ──────────────────────────────────────────────────────
def _pip(*pkgs):
    subprocess.check_call([sys.executable, "-m", "pip", "install", *pkgs, "--quiet"])

try: from PIL import Image, ImageTk, ImageDraw, ImageFilter, ImageOps, ImageFont, ImageEnhance
except ImportError: _pip("Pillow"); from PIL import Image, ImageTk, ImageDraw, ImageFilter, ImageOps, ImageFont, ImageEnhance

try: import pillow_heif; pillow_heif.register_heif_opener()
except ImportError: _pip("pillow-heif"); import pillow_heif; pillow_heif.register_heif_opener()

try: import fitz  # PyMuPDF
except ImportError: _pip("PyMuPDF"); import fitz

try: from openpyxl import load_workbook, Workbook
except ImportError: _pip("openpyxl"); from openpyxl import load_workbook, Workbook

try:
    import numpy as np
    import cv2
except ImportError:
    _pip("numpy", "opencv-python-headless")
    import numpy as np
    import cv2

try: import tkinterdnd2 as dnd; HAS_DND = True
except ImportError:
    try: _pip("tkinterdnd2"); import tkinterdnd2 as dnd; HAS_DND = True
    except Exception: HAS_DND = False

# ── Palette ───────────────────────────────────────────────────────────────────
BG       = "#0F0F13"
SURFACE  = "#1A1A22"
SURFACE2 = "#23232E"
BORDER   = "#2E2E3D"
ACCENT   = "#5B6BFF"
ACCENT2  = "#8B5CF6"
SUCCESS  = "#22C55E"
WARNING  = "#F59E0B"
DANGER   = "#EF4444"
TEXT     = "#F0F0F8"
MUTED    = "#7878A0"
DROP_HL  = "#3B4BCC"

PAPER_SIZES_MM = {}
DPI_OPTIONS = [72, 96, 150, 300, 600]

def _copypro_data_dir():
    """Durable per-user settings folder, independent of Downloads/Desktop."""
    if sys.platform.startswith("win"):
        root = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
        path = os.path.join(root, "CopyPro")
    elif sys.platform == "darwin":
        path = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "CopyPro")
    else:
        root = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
        path = os.path.join(root, "copypro")
    os.makedirs(path, exist_ok=True)
    return path

COPYPRO_DATA_DIR = _copypro_data_dir()
CUSTOM_SIZES_FILE = os.path.join(COPYPRO_DATA_DIR, "custom_sizes.json")
PRINT_LAYOUTS_FILE = os.path.join(COPYPRO_DATA_DIR, "print_layouts.json")
APP_SETTINGS_FILE = os.path.join(COPYPRO_DATA_DIR, "settings.json")
DEFAULT_CODES_FILE = os.path.join(COPYPRO_DATA_DIR, "copypro_kodai.xlsx")
DEFAULT_PAPER_SIZES_FILE = os.path.join(COPYPRO_DATA_DIR, "copypro_popieriaus_dydziai.xlsx")
DEFAULT_WIDE_FORMAT_FILE = DEFAULT_CODES_FILE
BACKUPS_DIR = os.path.join(COPYPRO_DATA_DIR, "backups")
os.makedirs(BACKUPS_DIR, exist_ok=True)

def resource_path(filename):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)

def _migrate_legacy_settings():
    legacy = {
        os.path.join(os.path.expanduser("~"), ".copypro_custom_sizes.json"): CUSTOM_SIZES_FILE,
        os.path.join(os.path.expanduser("~"), ".copypro_print_layouts.json"): PRINT_LAYOUTS_FILE,
    }
    for old_path, new_path in legacy.items():
        if os.path.exists(old_path) and not os.path.exists(new_path):
            try:
                import shutil
                shutil.copy2(old_path, new_path)
            except Exception:
                pass

_migrate_legacy_settings()

def load_app_settings():
    import json
    try:
        with open(APP_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_app_settings(data):
    import json
    try:
        with open(APP_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

def load_custom_sizes():
    import json
    if os.path.exists(CUSTOM_SIZES_FILE):
        try:
            with open(CUSTOM_SIZES_FILE, "r") as f:
                data = json.load(f)
            return {k: tuple(v) for k, v in data.items()}
        except Exception:
            return {}
    return {}

def save_custom_sizes(sizes_dict):
    import json
    try:
        with open(CUSTOM_SIZES_FILE, "w") as f:
            json.dump(sizes_dict, f, indent=2)
    except Exception:
        pass

IMAGE_EXTS = {".jpg",".jpeg",".png",".webp",".bmp",".tiff",".tif",
              ".gif",".heic",".heif",".avif",".ico",".ppm",".tga",
              ".raw",".cr2",".nef",".arw",".dng",".orf",".rw2"}
ALL_EXTS   = IMAGE_EXTS | {".pdf",".svg",".eps",".ai",".ps",".pages"}

CONV_FORMATS = ["PNG","JPEG","WEBP","BMP","TIFF","GIF","ICO","PDF","PPM","TGA"]

def mm_to_px(mm, dpi): return int(round(mm / 25.4 * dpi))
def px_to_mm(px, dpi): return px / dpi * 25.4

# ── Shared helpers ────────────────────────────────────────────────────────────
def styled_btn(parent, text, cmd, style="primary", **kw):
    cols = {"primary":(ACCENT,TEXT),"secondary":(SURFACE2,TEXT),
            "success":(SUCCESS,"#000"),"danger":(DANGER,TEXT)}
    bg, fg = cols.get(style, cols["primary"])
    options = {
        "text": text,
        "command": cmd,
        "bg": bg,
        "fg": fg,
        "relief": "flat",
        "cursor": "hand2",
        "font": ("Segoe UI",9,"bold"),
        "padx": 14,
        "pady": 7,
        "activebackground": ACCENT2,
        "activeforeground": TEXT,
    }
    options.update(kw)
    return tk.Button(parent, **options)

def sep(parent): tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=16, pady=6)

def lbl(parent, text, size=9, color=TEXT, bold=False, bg=None):
    return tk.Label(parent, text=text, bg=bg or SURFACE, fg=color,
                    font=("Segoe UI", size, "bold" if bold else "normal"))

def entry(parent, var=None, w=8):
    e = tk.Entry(parent, textvariable=var, bg=SURFACE2, fg=TEXT,
                 insertbackground=TEXT, relief="flat",
                 font=("Segoe UI",9), width=w)
    return e

def build_output_location(parent, owner, default_subfolder):
    """Add the shared output-location controls used by all processing tabs."""
    lbl(parent,"OUTPUT LOCATION",8,MUTED,True).pack(pady=(2,2),padx=16,anchor="w")
    owner.dest_var = tk.StringVar(value="choose")

    for text, value in (
        ("Choose folder…", "choose"),
        ("Same folder as originals", "same"),
        ("New subfolder here", "subfolder"),
    ):
        tk.Radiobutton(parent, text=text, variable=owner.dest_var, value=value,
            bg=SURFACE, fg=TEXT, selectcolor=SURFACE2, activebackground=SURFACE,
            font=("Segoe UI",9), command=owner._toggle_dest_opts
        ).pack(padx=16, anchor="w")

    owner._dest_sub = tk.Frame(parent, bg=SURFACE)
    owner._dest_sub.pack(fill="x", pady=(1,1))

    owner.overwrite_frame = tk.Frame(owner._dest_sub, bg=SURFACE)
    owner.overwrite_var = tk.BooleanVar(value=False)
    tk.Checkbutton(owner.overwrite_frame, text="Replace original files",
        variable=owner.overwrite_var, bg=SURFACE, fg=WARNING, selectcolor=SURFACE2,
        activebackground=SURFACE, font=("Segoe UI",8,"bold")
    ).pack(padx=22, anchor="w")

    owner.subfolder_frame = tk.Frame(owner._dest_sub, bg=SURFACE)
    row = tk.Frame(owner.subfolder_frame, bg=SURFACE)
    row.pack(fill="x", padx=22, pady=1)
    lbl(row,"Name:",8).pack(side="left", padx=(0,4))
    owner.subfolder_name = entry(row, w=14)
    owner.subfolder_name.insert(0, default_subfolder)
    owner.subfolder_name.pack(side="left")

def toggle_output_location(owner):
    owner.overwrite_frame.pack_forget()
    owner.subfolder_frame.pack_forget()
    if owner.dest_var.get() == "same":
        owner.overwrite_frame.pack(fill="x")
    elif owner.dest_var.get() == "subfolder":
        owner.subfolder_frame.pack(fill="x")

def get_output_resolver(owner, dialog_title):
    """Return (directory resolver, replace_originals), or (None, False) on cancel."""
    mode = owner.dest_var.get()
    replace = mode == "same" and owner.overwrite_var.get()
    if mode == "choose":
        out = filedialog.askdirectory(title=dialog_title)
        if not out:
            return None, False
        return (lambda src_path, d=out: d), False
    if mode == "same":
        if replace and not messagebox.askyesno(
            "Replace originals?",
            "This will replace the original files. This cannot be undone.\n\nContinue?"
        ):
            return None, False
        return (lambda src_path: os.path.dirname(src_path)), replace
    name = owner.subfolder_name.get().strip() or "output"
    def resolver(src_path, folder_name=name):
        out = os.path.join(os.path.dirname(src_path), folder_name)
        os.makedirs(out, exist_ok=True)
        return out
    return resolver, False

def open_pdf_pages(path, dpi=150):
    """Rasterise every page of a PDF, return list of PIL Images."""
    doc = fitz.open(path)
    imgs = []
    mat = fitz.Matrix(dpi/72, dpi/72)
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        imgs.append(img)
    doc.close()
    return imgs

def parse_dropped_paths(data):
    """Parse tkinterdnd2 drop data into a list of file paths."""
    paths = []
    # DnD data can be space-separated, possibly with braces for paths with spaces
    raw = data.strip()
    import re
    # Extract {path with spaces} or plain tokens
    tokens = re.findall(r'\{([^}]+)\}|(\S+)', raw)
    for t in tokens:
        p = t[0] or t[1]
        if p:
            paths.append(p)
    return paths

# ── Drop Zone mixin ───────────────────────────────────────────────────────────
class DropMixin:
    """Reliable whole-tab file dropping with a non-flickering in-tab overlay."""

    def setup_drop(self, widget, callback, extensions=None, root=None):
        if not HAS_DND:
            return
        try:
            drop_root = root or getattr(widget, "_copypro_drop_root", widget)
            widget._copypro_drop_root = drop_root
            widget._copypro_drop_callback = callback
            widget._copypro_drop_extensions = extensions
            if getattr(widget, "_copypro_dnd_registered", False):
                return
            widget.drop_target_register(dnd.DND_FILES)
            widget.dnd_bind("<<DropEnter>>", lambda e, w=widget: self._drop_enter(w))
            widget.dnd_bind("<<DropPosition>>", lambda e, w=widget: self._drop_enter(w))
            widget.dnd_bind("<<DropLeave>>", lambda e, w=widget: self._drop_leave(w))
            widget.dnd_bind(
                "<<Drop>>",
                lambda e, w=widget: self._on_drop(
                    e,
                    w,
                    getattr(w, "_copypro_drop_callback", callback),
                    getattr(w, "_copypro_drop_extensions", extensions),
                ),
            )
            widget._copypro_dnd_registered = True
        except Exception:
            pass

    def setup_drop_everywhere(self, root, callback, extensions=None):
        root._copypro_drop_root = root
        root._copypro_drop_callback = callback
        root._copypro_drop_extensions = extensions

        def register_tree(widget):
            self.setup_drop(widget, callback, extensions, root=root)
            try:
                for child in widget.winfo_children():
                    register_tree(child)
            except Exception:
                pass

        register_tree(root)

        # Pick up dynamically rebuilt cards without repeatedly rebinding old widgets.
        def refresh():
            try:
                register_tree(root)
                root._copypro_drop_refresh_job = root.after(3000, refresh)
            except Exception:
                pass

        try:
            old = getattr(root, "_copypro_drop_refresh_job", None)
            if old:
                root.after_cancel(old)
            root._copypro_drop_refresh_job = root.after(3000, refresh)
        except Exception:
            pass

    def _drop_root_for(self, widget):
        return getattr(widget, "_copypro_drop_root", widget)

    def _show_drop_overlay(self, widget):
        """Show one childless in-tab canvas so drag enter/leave cannot flicker."""
        try:
            root=self._drop_root_for(widget)
            hide_job=getattr(root,"_copypro_drop_hide_job",None)
            if hide_job:
                try:root.after_cancel(hide_job)
                except Exception:pass
                root._copypro_drop_hide_job=None

            overlay=getattr(root,"_copypro_drop_overlay",None)
            if overlay is None or not overlay.winfo_exists():
                overlay=tk.Canvas(
                    root,bg=SURFACE,highlightbackground=ACCENT,
                    highlightthickness=3,bd=0,cursor="arrow"
                )
                root._copypro_drop_overlay=overlay
                callback=getattr(root,"_copypro_drop_callback",None)
                extensions=getattr(root,"_copypro_drop_extensions",None)
                if callback:self.setup_drop(overlay,callback,extensions,root=root)
                def redraw(_event=None,o=overlay):
                    try:
                        o.delete("all"); w=max(1,o.winfo_width()); h=max(1,o.winfo_height())
                        o.create_text(w/2,h/2-12,text="DROP FILES HERE",fill=TEXT,
                                      font=("Segoe UI",14,"bold"))
                        o.create_text(w/2,h/2+16,text="Release to add the selected files",
                                      fill=MUTED,font=("Segoe UI",9))
                    except Exception:pass
                overlay.bind("<Configure>",redraw)

            overlay.place(relx=.5,rely=.5,anchor="center",width=330,height=100)
            overlay.lift()
        except Exception:pass

    def _hide_drop_overlay_now(self, widget):
        try:
            root = self._drop_root_for(widget)
            overlay = getattr(root, "_copypro_drop_overlay", None)
            if overlay is not None and overlay.winfo_exists():
                overlay.place_forget()
            root._copypro_drop_hide_job = None
        except Exception:
            pass

    def _drop_enter(self, widget):
        self._show_drop_overlay(widget)
        return getattr(dnd, "COPY", "copy") if HAS_DND else None

    def _drop_leave(self, widget):
        # Child controls and the overlay can generate leave/enter pairs. Delay the
        # hide briefly so the following enter cancels it instead of flickering.
        try:
            root = self._drop_root_for(widget)
            old = getattr(root, "_copypro_drop_hide_job", None)
            if old:
                root.after_cancel(old)
            root._copypro_drop_hide_job = root.after(
                350,
                lambda w=widget: self._hide_drop_overlay_now(w),
            )
        except Exception:
            pass
        return getattr(dnd, "COPY", "copy") if HAS_DND else None

    def _on_drop(self, event, widget, callback, extensions):
        self._hide_drop_overlay_now(widget)
        paths = parse_dropped_paths(event.data)
        filtered = []
        for path in paths:
            if os.path.isfile(path):
                if extensions is None or os.path.splitext(path)[1].lower() in extensions:
                    filtered.append(path)
            elif os.path.isdir(path):
                for name in os.listdir(path):
                    candidate = os.path.join(path, name)
                    if os.path.isfile(candidate):
                        ext = os.path.splitext(candidate)[1].lower()
                        if extensions is None or ext in extensions:
                            filtered.append(candidate)
        if filtered:
            callback(filtered)
        return getattr(dnd, "COPY", "copy") if HAS_DND else None

class ImageMemoryCache:
    """Byte-limited LRU cache that closes PIL images when evicted."""
    def __init__(self, max_bytes):
        self.max_bytes = int(max_bytes)
        self.current_bytes = 0
        self._items = OrderedDict()

    @staticmethod
    def _size(image):
        try:
            return image.width * image.height * max(1, len(image.getbands()))
        except Exception:
            return 0

    def get(self, key, default=None):
        value = self._items.get(key)
        if value is None:
            return default
        self._items.move_to_end(key)
        return value[0]

    def __setitem__(self, key, image):
        old = self._items.pop(key, None)
        if old:
            self.current_bytes -= old[1]
            if old[0] is not image:
                try: old[0].close()
                except Exception: pass
        size = self._size(image)
        self._items[key] = (image, size)
        self.current_bytes += size
        while self.current_bytes > self.max_bytes and len(self._items) > 1:
            _, (victim, victim_size) = self._items.popitem(last=False)
            self.current_bytes -= victim_size
            try: victim.close()
            except Exception: pass

    def clear(self):
        for image, _size in self._items.values():
            try: image.close()
            except Exception: pass
        self._items.clear()
        self.current_bytes = 0

    def __len__(self):
        return len(self._items)


def process_memory_mb():
    try:
        if sys.platform.startswith("win"):
            class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
                _fields_ = [
                    ("cb", ctypes.c_ulong),
                    ("PageFaultCount", ctypes.c_ulong),
                    ("PeakWorkingSetSize", ctypes.c_size_t),
                    ("WorkingSetSize", ctypes.c_size_t),
                    ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                    ("PagefileUsage", ctypes.c_size_t),
                    ("PeakPagefileUsage", ctypes.c_size_t),
                ]
            counters = PROCESS_MEMORY_COUNTERS()
            counters.cb = ctypes.sizeof(counters)
            handle = ctypes.windll.kernel32.GetCurrentProcess()
            ctypes.windll.psapi.GetProcessMemoryInfo(handle, ctypes.byref(counters), counters.cb)
            return counters.WorkingSetSize / (1024 * 1024)
        statm = Path("/proc/self/statm")
        if statm.exists():
            resident_pages = int(statm.read_text().split()[1])
            return resident_pages * os.sysconf("SC_PAGE_SIZE") / (1024 * 1024)
    except Exception:
        pass
    return 0.0

# ── Crop Canvas ───────────────────────────────────────────────────────────────
PREVIEW_MAX_DIM = 900  # cap preview image's long side in px — keeps UI fast
HANDLE_SIZE = 9         # corner handle hit-box half-size, in canvas px
MIN_CROP_FRAC = 0.15    # don't allow shrinking the crop box below 15% of the fitted size

class CropCanvas(tk.Canvas):
    def __init__(self, parent, img_path, target_w, target_h, size=(300,300), on_change=None, **kw):
        super().__init__(parent, width=size[0], height=size[1],
                         bg=SURFACE2, highlightthickness=0, **kw)
        self.on_change = on_change
        self.target_w, self.target_h = target_w, target_h
        self.canvas_w, self.canvas_h = size
        self.src_path = img_path
        self.total_rotation = 0  # cumulative rotation (deg, multiple of 90) vs. EXIF-corrected original
        self._load(img_path)
        self._init_crop()
        self._draw()
        self.bind("<ButtonPress-1>",   self._press)
        self.bind("<B1-Motion>",       self._drag)
        self.bind("<ButtonRelease-1>", self._release_crop)
        self.bind("<Motion>",          self._on_hover)
        self._ds = None
        self._drag_mode = None  # "move" or one of "nw","ne","sw","se"

    def _notify_change(self):
        if callable(self.on_change):
            try: self.on_change()
            except Exception: pass

    def _release_crop(self, _event=None):
        self._ds = None
        self._notify_change()

    @property
    def rotated(self):
        """Kept for backward compatibility: True if a 90/270 rotation is in effect."""
        return self.total_rotation in (90, 270)

    def _load(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            doc = fitz.open(path)
            try:
                if len(doc):
                    pix = doc[0].get_pixmap(matrix=fitz.Matrix(1,1), alpha=False)
                    full = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                else:
                    full = Image.new("RGB", (100,100))
            finally:
                doc.close()
        else:
            with Image.open(path) as opened:
                corrected = ImageOps.exif_transpose(opened)
                full = corrected.convert("RGB")
                if corrected is not opened:
                    try: corrected.close()
                    except Exception: pass

        iw, ih = full.size
        self.source_aspect_ratio = iw / max(1, ih)
        # Auto-rotate to match target orientation
        target_portrait = self.target_h >= self.target_w
        img_portrait    = ih >= iw
        if target_portrait != img_portrait:
            full = full.rotate(90, expand=True)
            self.total_rotation = 90
            iw, ih = full.size

        # Downscale for a fast, low-memory preview — crop math uses fractions,
        # so resolution here doesn't affect export quality.
        if max(iw, ih) > PREVIEW_MAX_DIM:
            pscale = PREVIEW_MAX_DIM / max(iw, ih)
            pw, ph = max(1,int(iw*pscale)), max(1,int(ih*pscale))
            self.pil_img = full.resize((pw, ph), Image.LANCZOS)
        else:
            self.pil_img = full
        del full

        self._fit_to_canvas()

    def _fit_to_canvas(self):
        """(Re)compute display scale/offset and the Tk image from self.pil_img.
        Call this after self.pil_img changes (e.g. manual rotation) without
        needing to reload the file from disk."""
        iw, ih = self.pil_img.size
        scale = min(self.canvas_w/iw, self.canvas_h/ih)
        dw, dh = int(iw*scale), int(ih*scale)
        self.offset_x = (self.canvas_w-dw)//2
        self.offset_y = (self.canvas_h-dh)//2
        self.disp_w, self.disp_h, self.scale = dw, dh, scale
        resized = self.pil_img.resize((dw,dh), Image.Resampling.BILINEAR)
        try:
            self.tk_img = ImageTk.PhotoImage(resized)
        finally:
            resized.close()

    def rotate_manual(self, delta):
        """Rotate the preview (and the eventual export) by delta degrees (±90).
        Lossless since it's always a multiple of 90."""
        self.pil_img = self.pil_img.rotate(delta, expand=True)
        self.total_rotation = (self.total_rotation + delta) % 360
        self._fit_to_canvas()
        self._init_crop()
        self._draw()
        self._notify_change()

    def _init_crop(self):
        """Default crop = FIT the whole image (nothing cropped away).
        The box matches the smallest aspect-locked rectangle that contains
        the entire displayed image — i.e. it's exactly the full image extent
        when the image's own aspect equals the target, otherwise it's the
        full image with the box sized to its own bounds (export pads the rest)."""
        asp = self.target_w / self.target_h
        img_asp = self.disp_w / self.disp_h
        if img_asp >= asp:
            # image is relatively wider than target -> fit by width (full width shown)
            cw = self.disp_w
            ch = int(cw / asp)
            if ch > self.disp_h:
                ch = self.disp_h
                cw = int(ch * asp)
        else:
            # image is relatively taller than target -> fit by height (full height shown)
            ch = self.disp_h
            cw = int(ch * asp)
            if cw > self.disp_w:
                cw = self.disp_w
                ch = int(cw / asp)
        # Track the "fit" box size as the reference max for resize limits
        self._fit_w, self._fit_h = cw, ch
        self.crop_x = self.offset_x + (self.disp_w-cw)//2
        self.crop_y = self.offset_y + (self.disp_h-ch)//2
        self.crop_w, self.crop_h = cw, ch

    def _handle_positions(self):
        x1,y1 = self.crop_x, self.crop_y
        x2,y2 = x1+self.crop_w, y1+self.crop_h
        return {"nw":(x1,y1), "ne":(x2,y1), "sw":(x1,y2), "se":(x2,y2)}

    def _hit_handle(self, x, y):
        for name, (hx,hy) in self._handle_positions().items():
            if abs(x-hx) <= HANDLE_SIZE and abs(y-hy) <= HANDLE_SIZE:
                return name
        return None

    def _draw(self):
        self.delete("all")
        self.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.tk_img)
        x1,y1 = self.crop_x, self.crop_y
        x2,y2 = x1+self.crop_w, y1+self.crop_h
        ox,oy = self.offset_x, self.offset_y
        for r in [(ox,oy,x1,oy+self.disp_h),(x2,oy,ox+self.disp_w,oy+self.disp_h),
                  (x1,oy,x2,y1),(x1,y2,x2,oy+self.disp_h)]:
            self.create_rectangle(*r, fill="#000", stipple="gray50", outline="")
        self.create_rectangle(x1,y1,x2,y2, outline=ACCENT, width=2)
        for i in (1,2):
            self.create_line(x1+self.crop_w*i//3,y1,x1+self.crop_w*i//3,y2,fill=ACCENT,dash=(3,3))
            self.create_line(x1,y1+self.crop_h*i//3,x2,y1+self.crop_h*i//3,fill=ACCENT,dash=(3,3))
        for cx,cy in [(x1,y1),(x2,y1),(x1,y2),(x2,y2)]:
            self.create_rectangle(cx-HANDLE_SIZE,cy-HANDLE_SIZE,cx+HANDLE_SIZE,cy+HANDLE_SIZE,
                                  fill=ACCENT, outline=SURFACE2, width=1)

    def _on_hover(self, e):
        if self._ds:  # don't change cursor mid-drag
            return
        h = self._hit_handle(e.x, e.y)
        cursors = {"nw":"size_nw_se", "se":"size_nw_se", "ne":"size_ne_sw", "sw":"size_ne_sw"}
        if h:
            self.config(cursor=cursors[h])
        else:
            self.config(cursor="fleur")

    def _press(self, e):
        h = self._hit_handle(e.x, e.y)
        if h:
            self._drag_mode = h
            self._ds = (e.x, e.y, self.crop_x, self.crop_y, self.crop_w, self.crop_h)
        else:
            self._drag_mode = "move"
            self._ds = (e.x, e.y, self.crop_x, self.crop_y)

    def _drag(self, e):
        if not self._ds: return
        if self._drag_mode == "move":
            sx,sy,ox,oy = self._ds
            nx = max(self.offset_x, min(ox+e.x-sx, self.offset_x+self.disp_w-self.crop_w))
            ny = max(self.offset_y, min(oy+e.y-sy, self.offset_y+self.disp_h-self.crop_h))
            self.crop_x, self.crop_y = nx, ny
        else:
            self._resize_from_handle(e)
        self._draw()
        self._notify_change()

    def _resize_from_handle(self, e):
        """Resize the crop box from a corner, keeping the target aspect ratio locked.
        The opposite corner stays anchored in place."""
        sx, sy, ox, oy, ow, oh = self._ds
        asp = self.target_w / self.target_h
        min_w = max(20, self._fit_w * MIN_CROP_FRAC)
        min_h = max(20, self._fit_h * MIN_CROP_FRAC)

        # Anchor = the corner opposite to the one being dragged (stays fixed)
        anchors = {
            "se": (ox, oy),                 # top-left stays fixed
            "nw": (ox+ow, oy+oh),           # bottom-right stays fixed
            "ne": (ox, oy+oh),              # bottom-left stays fixed
            "sw": (ox+ow, oy),              # top-right stays fixed
        }
        ax, ay = anchors[self._drag_mode]

        # Use horizontal mouse movement to drive the resize (consistent + simple),
        # then derive height from the locked aspect ratio.
        dx = e.x - ax
        dy = e.y - ay
        # Determine desired width from whichever drag axis moved more,
        # for a natural feel when dragging diagonally.
        if abs(dx) >= abs(dy) * asp:
            new_w = abs(dx)
        else:
            new_w = abs(dy) * asp
        new_w = max(min_w, new_w)
        new_h = new_w / asp
        if new_h < min_h:
            new_h = min_h
            new_w = new_h * asp

        # Compute new crop_x/crop_y/crop_w/crop_h based on which corner is anchored
        if self._drag_mode == "se":
            nx, ny = ax, ay
        elif self._drag_mode == "nw":
            nx, ny = ax-new_w, ay-new_h
        elif self._drag_mode == "ne":
            nx, ny = ax, ay-new_h
        else:  # sw
            nx, ny = ax-new_w, ay

        # Clamp to the visible image bounds
        nx = max(self.offset_x, min(nx, self.offset_x+self.disp_w-new_w))
        ny = max(self.offset_y, min(ny, self.offset_y+self.disp_h-new_h))
        if nx + new_w > self.offset_x + self.disp_w:
            new_w = self.offset_x + self.disp_w - nx
            new_h = new_w / asp
        if ny + new_h > self.offset_y + self.disp_h:
            new_h = self.offset_y + self.disp_h - ny
            new_w = new_h * asp

        self.crop_x, self.crop_y = nx, ny
        self.crop_w, self.crop_h = new_w, new_h

    def get_crop_box(self):
        rx = (self.crop_x-self.offset_x)/self.scale
        ry = (self.crop_y-self.offset_y)/self.scale
        rw = self.crop_w/self.scale; rh = self.crop_h/self.scale
        iw,ih = self.pil_img.size
        return (max(0,int(rx)), max(0,int(ry)),
                min(iw,int(rx+rw)), min(ih,int(ry+rh)))

    def get_crop_fractions(self):
        """Crop box as (left, top, right, bottom) fractions of the (rotated) image, 0..1.
        Resolution-independent — use this to apply the same crop to a full-res reload."""
        iw, ih = self.pil_img.size
        l, t, r, b = self.get_crop_box()
        return (l/iw, t/ih, r/iw, b/ih)

# ── Bleed helper ─────────────────────────────────────────────────────────────
def add_bleed_and_marks(img,bleed_mm,dpi,crop_marks=True,mark_len_mm=5,
                        mark_gap_mm=2,bleed_mode="reflected",
                        bleed_color=(255,255,255)):
    """Add reflected or solid-colour bleed and optional crop marks."""
    bp=max(0,mm_to_px(bleed_mm,dpi)); iw,ih=img.size
    extra=mm_to_px(mark_len_mm+mark_gap_mm+2,dpi) if crop_marks else 0
    canvas=Image.new("RGB",(iw+bp*2+extra*2,ih+bp*2+extra*2),"white")
    ox,oy=extra+bp,extra+bp; source=img.convert("RGB")
    if bleed_mode=="solid":
        canvas.paste(Image.new("RGB",(iw+bp*2,ih+bp*2),tuple(bleed_color)),(extra,extra))
        canvas.paste(source,(ox,oy))
    else:
        canvas.paste(source,(ox,oy))
        if bp:
            left=source.crop((0,0,min(bp,iw),ih)).resize((bp,ih)).transpose(Image.FLIP_LEFT_RIGHT)
            right=source.crop((max(0,iw-bp),0,iw,ih)).resize((bp,ih)).transpose(Image.FLIP_LEFT_RIGHT)
            top=source.crop((0,0,iw,min(bp,ih))).resize((iw,bp)).transpose(Image.FLIP_TOP_BOTTOM)
            bottom=source.crop((0,max(0,ih-bp),iw,ih)).resize((iw,bp)).transpose(Image.FLIP_TOP_BOTTOM)
            canvas.paste(left,(ox-bp,oy)); canvas.paste(right,(ox+iw,oy))
            canvas.paste(top,(ox,oy-bp)); canvas.paste(bottom,(ox,oy+ih))
            corners=[
                ((0,0,min(bp,iw),min(bp,ih)),(ox-bp,oy-bp)),
                ((max(0,iw-bp),0,iw,min(bp,ih)),(ox+iw,oy-bp)),
                ((0,max(0,ih-bp),min(bp,iw),ih),(ox-bp,oy+ih)),
                ((max(0,iw-bp),max(0,ih-bp),iw,ih),(ox+iw,oy+ih)),
            ]
            for box,pos in corners:
                canvas.paste(source.crop(box).resize((bp,bp)).transpose(Image.ROTATE_180),pos)
    if crop_marks:
        draw=ImageDraw.Draw(canvas); gap=mm_to_px(mark_gap_mm,dpi)
        mlen=mm_to_px(mark_len_mm,dpi); lw=max(1,mm_to_px(.25,dpi))
        for cx,cy,dx,dy in ((ox-bp,oy-bp,-1,-1),(ox+iw+bp,oy-bp,1,-1),
                            (ox-bp,oy+ih+bp,-1,1),(ox+iw+bp,oy+ih+bp,1,1)):
            draw.line((cx+dx*gap,cy,cx+dx*(gap+mlen),cy),fill="black",width=lw)
            draw.line((cx,cy+dy*gap,cx,cy+dy*(gap+mlen)),fill="black",width=lw)
    return canvas

def add_crop_marks_only(img, dpi, mark_len_mm=5, mark_gap_mm=2):
    """Add crop marks around an image without adding bleed."""
    gap = mm_to_px(mark_gap_mm, dpi)
    mlen = mm_to_px(mark_len_mm, dpi)
    margin = gap + mlen + mm_to_px(2, dpi)
    iw, ih = img.size
    canvas = Image.new("RGB", (iw + margin*2, ih + margin*2), "white")
    canvas.paste(img.convert("RGB"), (margin, margin))
    draw = ImageDraw.Draw(canvas)
    lw = max(1, mm_to_px(0.25, dpi))
    x1, y1, x2, y2 = margin, margin, margin+iw, margin+ih
    for x, dx in ((x1,-1),(x2,1)):
        for y, dy in ((y1,-1),(y2,1)):
            draw.line((x+dx*gap,y,x+dx*(gap+mlen),y),fill="black",width=lw)
            draw.line((x,y+dy*gap,x,y+dy*(gap+mlen)),fill="black",width=lw)
    return canvas

def apply_resizer_finishing(img,dpi,add_bleed,bleed_mm,crop_marks,
                              bleed_mode="reflected",bleed_color=(255,255,255)):
    if add_bleed and bleed_mm>0:
        return add_bleed_and_marks(img,bleed_mm,dpi,crop_marks,5,2,
                                   bleed_mode,bleed_color)
    if crop_marks:return add_crop_marks_only(img,dpi,5,2)
    return img

# ── Resizer Tab ───────────────────────────────────────────────────────────────
class ResizerTab(tk.Frame, DropMixin):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.files = []; self.canvases = []
        self.custom_sizes = load_custom_sizes()  # name -> (w_mm, h_mm)
        self._build()

    def _all_size_names(self):
        return list(PAPER_SIZES_MM.keys()) + list(self.custom_sizes.keys()) + ["Custom…"]

    def _build(self):
        left_outer = tk.Frame(self, bg=SURFACE, width=250)
        left_outer.pack(side="left", fill="y", padx=(0,1))
        left_outer.pack_propagate(False)

        # Fixed, top-aligned settings panel. All controls fit, so this area should not scroll.
        left = tk.Frame(left_outer, bg=SURFACE)
        left.pack(fill="both", expand=True, anchor="n")

        lbl(left,"RESIZE SETTINGS",8,MUTED,True).pack(pady=(14,3),padx=16,anchor="w")

        lbl(left,"Paper size").pack(padx=16,anchor="w")
        self.size_var = tk.StringVar(value="A4")
        self.size_cb = ttk.Combobox(left, textvariable=self.size_var,
            values=self._all_size_names(),state="readonly",width=20)
        self.size_cb.pack(fill="x",padx=16,pady=(2,2))
        self.size_cb.bind("<<ComboboxSelected>>", self._on_size)
        # Label shown below dropdown when Custom is active
        self.custom_size_lbl = tk.Label(left, text="", bg=SURFACE, fg=ACCENT,
            font=("Segoe UI", 10, "bold"), anchor="w")
        self.custom_w_mm = 100.0
        self.custom_h_mm = 100.0

        self.manage_lbl = tk.Label(left, text="Manage saved sizes…", bg=SURFACE,
            fg=MUTED, font=("Segoe UI", 8, "underline"), cursor="hand2", anchor="w")
        self.manage_lbl.pack(padx=16, anchor="w", pady=(0,4))
        self.manage_lbl.bind("<Button-1>", lambda e: self._open_manage_sizes())

        # Orientation + DPI side by side to save vertical space
        od_row = tk.Frame(left, bg=SURFACE)
        od_row.pack(fill="x", padx=16, pady=(2,2))
        od_left = tk.Frame(od_row, bg=SURFACE); od_left.pack(side="left", fill="x", expand=True)
        od_right = tk.Frame(od_row, bg=SURFACE); od_right.pack(side="left", fill="x", expand=True, padx=(6,0))

        lbl(od_left,"Orientation").pack(anchor="w")
        self.orient_var = tk.StringVar(value="Portrait")
        orow = tk.Frame(od_left,bg=SURFACE); orow.pack(fill="x", pady=(2,0))
        for o,short in (("Portrait","Port."),("Landscape","Land.")):
            tk.Radiobutton(orow,text=short,variable=self.orient_var,value=o,
                bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,
                font=("Segoe UI",8), command=self._refresh).pack(side="left")

        lbl(od_right,"DPI").pack(anchor="w")
        self.dpi_var = tk.IntVar(value=300)
        ttk.Combobox(od_right,textvariable=self.dpi_var,values=DPI_OPTIONS,
                     state="readonly",width=8).pack(fill="x", pady=(2,0))

        lbl(left,"Output format").pack(padx=16,anchor="w",pady=(4,0))
        self.fmt_var = tk.StringVar(value="JPEG")
        self.fmt_cb = ttk.Combobox(left,textvariable=self.fmt_var,
            values=["JPEG","PNG","TIFF","PDF"],state="readonly",width=20)
        self.fmt_cb.pack(fill="x",padx=16,pady=(2,4))
        self.fmt_cb.bind("<<ComboboxSelected>>", self._toggle_quality)

        self.q_lbl = lbl(left,"JPEG Quality: 90")
        self.q_lbl.pack(padx=16,anchor="w")
        self.q_var = tk.IntVar(value=90)
        self.q_slider = ttk.Scale(left,from_=10,to=100,variable=self.q_var,
            orient="horizontal",
            command=lambda v: self.q_lbl.config(text=f"JPEG Quality: {int(float(v))}"))
        self.q_slider.pack(fill="x",padx=16,pady=(0,4))

        sep(left)

        # Bleed & crop marks section
        lbl(left,"BLEED & CROP MARKS",8,MUTED,True).pack(pady=(2,2),padx=16,anchor="w")

        self.bleed_row = tk.Frame(left, bg=SURFACE)
        self.bleed_row.pack(fill="x", padx=16, pady=(0,2))
        bleed_row = self.bleed_row
        self.bleed_var = tk.BooleanVar(value=False)
        tk.Checkbutton(bleed_row, text="Add bleed", variable=self.bleed_var,
            bg=SURFACE, fg=TEXT, selectcolor=SURFACE2, activebackground=SURFACE,
            font=("Segoe UI",9), command=self._toggle_bleed).pack(side="left")

        # Bleed amount and mode are shown only when bleed is enabled.
        self.bleed_opts=tk.Frame(left,bg=SURFACE)
        br=tk.Frame(self.bleed_opts,bg=SURFACE); br.pack(fill="x",padx=4,pady=1)
        lbl(br,"Bleed (mm)",8).pack(side="left",padx=(0,4))
        self.bleed_mm=entry(br,w=5); self.bleed_mm.insert(0,"3"); self.bleed_mm.pack(side="left")
        self.bleed_mode_var=tk.StringVar(value="reflected")
        mode=tk.Frame(self.bleed_opts,bg=SURFACE); mode.pack(fill="x",padx=4,pady=(2,1))
        for caption,value in (("Reflected","reflected"),("Solid colour","solid")):
            tk.Radiobutton(mode,text=caption,variable=self.bleed_mode_var,value=value,
                bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,
                font=("Segoe UI",8),command=self._toggle_bleed_colour).pack(side="left",padx=(0,6))
        self.bleed_color=(255,255,255)
        self.bleed_color_btn=tk.Button(self.bleed_opts,text="Choose bleed colour",
            command=self._choose_bleed_colour,bg="#FFFFFF",fg="#111111",
            relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2",padx=7,pady=3)

        marks_row=tk.Frame(left,bg=SURFACE); marks_row.pack(fill="x",padx=16,pady=(1,2))
        self.marks_var=tk.BooleanVar(value=False)
        tk.Checkbutton(marks_row,text="Add crop marks",variable=self.marks_var,
            bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,
            font=("Segoe UI",9)).pack(side="left")

        self.gray_var = tk.BooleanVar(value=False)
        tk.Checkbutton(left, text="Grayscale", variable=self.gray_var,
            bg=SURFACE, fg=TEXT, selectcolor=SURFACE2, activebackground=SURFACE,
            font=("Segoe UI",9)).pack(padx=16,anchor="w",pady=(1,0))

        self.stretch_var = tk.BooleanVar(value=False)
        tk.Checkbutton(left, text="Stretch to fit", variable=self.stretch_var,
            bg=SURFACE, fg=TEXT, selectcolor=SURFACE2, activebackground=SURFACE,
            font=("Segoe UI",9)).pack(padx=16,anchor="w",pady=(1,0))

        sep(left)

        # Output location section
        lbl(left,"OUTPUT LOCATION",8,MUTED,True).pack(pady=(2,2),padx=16,anchor="w")
        self.dest_var = tk.StringVar(value="choose")

        opt1 = tk.Radiobutton(left, text="Choose folder…", variable=self.dest_var,
            value="choose", bg=SURFACE, fg=TEXT, selectcolor=SURFACE2,
            activebackground=SURFACE, font=("Segoe UI",9), command=self._toggle_dest_opts)
        opt1.pack(padx=16, anchor="w")

        opt2 = tk.Radiobutton(left, text="Same folder as originals", variable=self.dest_var,
            value="same", bg=SURFACE, fg=TEXT, selectcolor=SURFACE2,
            activebackground=SURFACE, font=("Segoe UI",9), command=self._toggle_dest_opts)
        opt2.pack(padx=16, anchor="w")

        opt3 = tk.Radiobutton(left, text="New subfolder here", variable=self.dest_var,
            value="subfolder", bg=SURFACE, fg=TEXT, selectcolor=SURFACE2,
            activebackground=SURFACE, font=("Segoe UI",9), command=self._toggle_dest_opts)
        opt3.pack(padx=16, anchor="w")

        # Fixed-position container for sub-options — always packed here,
        # only its INNER contents toggle, so order never shifts.
        self._dest_sub = tk.Frame(left, bg=SURFACE)
        self._dest_sub.pack(fill="x", padx=0, pady=(1,1))

        self.overwrite_frame = tk.Frame(self._dest_sub, bg=SURFACE)
        self.overwrite_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.overwrite_frame, text="Replace original files",
            variable=self.overwrite_var, bg=SURFACE, fg=WARNING, selectcolor=SURFACE2,
            activebackground=SURFACE, font=("Segoe UI",8,"bold")).pack(padx=22, anchor="w")

        self.subfolder_frame = tk.Frame(self._dest_sub, bg=SURFACE)
        sfr = tk.Frame(self.subfolder_frame, bg=SURFACE); sfr.pack(fill="x", padx=22, pady=1)
        lbl(sfr,"Name:",8).pack(side="left", padx=(0,4))
        self.subfolder_name = entry(sfr, w=14)
        self.subfolder_name.insert(0, "resized")
        self.subfolder_name.pack(side="left")

        sep(left)

        btn_row1 = tk.Frame(left, bg=SURFACE); btn_row1.pack(fill="x", padx=16, pady=2)
        styled_btn(btn_row1,"➕ Add",self._add_files).pack(side="left", fill="x", expand=True, padx=(0,3))
        styled_btn(btn_row1,"🗑 Clear",self._clear,style="secondary").pack(side="left", fill="x", expand=True, padx=(3,0))
        self.exp_btn = styled_btn(left,"💾  Export All",self._export,style="success")
        self.exp_btn.pack(fill="x",padx=16,pady=(4,2))
        self.status = lbl(left,"",8,MUTED)
        self.status.pack(padx=16,pady=(0,10),anchor="w")

        # Right scrollable area
        right = tk.Frame(self, bg=BG)
        right.pack(side="left",fill="both",expand=True)

        # Drop zone hint at top
        self.drop_zone = tk.Label(right,
            text="⬇  Drop images here, or use 'Add Images'",
            bg=SURFACE2, fg=MUTED, font=("Segoe UI",10),
            pady=10, cursor="hand2")
        self.drop_zone.pack(fill="x",padx=12,pady=(8,0))
        self.drop_zone.bind("<Button-1>", lambda e: self._add_files())
        self.setup_drop(self.drop_zone, self._drop_files, IMAGE_EXTS)

        self.scroll_c = tk.Canvas(right,bg=BG,highlightthickness=0)
        self.scroll_c.pack(fill="both",expand=True)
        self.grid_f = tk.Frame(self.scroll_c,bg=BG)
        self.grid_win = self.scroll_c.create_window((0,0),window=self.grid_f,anchor="nw")
        self.grid_f.bind("<Configure>",
            lambda e: self.scroll_c.configure(scrollregion=self.scroll_c.bbox("all")))
        self.COLS_4_THRESHOLD = 760  # canvas width (px) above which we switch from 3 to 4 columns
        self._current_cols = None
        self._resize_job = None
        self.scroll_c.bind("<Configure>", self._on_canvas_resize)
        # Mouse-wheel scroll without a visible scrollbar
        self.scroll_c.bind("<Enter>", lambda e: self.scroll_c.bind_all("<MouseWheel>",
            lambda ev: self.scroll_c.yview_scroll(int(-1*(ev.delta/120)),"units")))
        self.scroll_c.bind("<Leave>", lambda e: self.scroll_c.unbind_all("<MouseWheel>"))
        self.setup_drop(self.scroll_c, self._drop_files, IMAGE_EXTS)

        self._show_empty()

    def _on_canvas_resize(self, e):
        # Always keep the inner frame's width synced to the canvas (cheap, no rebuild)
        self.scroll_c.itemconfig(self.grid_win, width=e.width)
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(150, lambda w=e.width: self._maybe_switch_columns(w))

    def _maybe_switch_columns(self, width):
        self._resize_job = None
        target_cols = 4 if width >= self.COLS_4_THRESHOLD else 3
        if target_cols != self._current_cols and self.files:
            self._rebuild()

    def _toggle_quality(self, e=None):
        if self.fmt_var.get()=="JPEG":
            self.q_lbl.pack(padx=16,anchor="w"); self.q_slider.pack(fill="x",padx=16,pady=(0,8))
        else:
            self.q_lbl.pack_forget(); self.q_slider.pack_forget()

    def _choose_bleed_colour(self):
        result=colorchooser.askcolor(color="#%02X%02X%02X"%self.bleed_color,
                                     title="Choose bleed colour",parent=self)
        if result and result[0]:
            self.bleed_color=tuple(int(round(v)) for v in result[0])
            self.bleed_color_btn.config(bg=result[1],
                fg="#111111" if sum(self.bleed_color)/3>150 else "#FFFFFF")

    def _toggle_bleed_colour(self):
        if self.bleed_mode_var.get()=="solid":
            self.bleed_color_btn.pack(anchor="w",padx=4,pady=(2,2))
        else:self.bleed_color_btn.pack_forget()

    def _toggle_bleed(self):
        if self.bleed_var.get():
            self.bleed_opts.pack(fill="x",padx=16,pady=(0,6),after=self.bleed_row)
            self._toggle_bleed_colour()
        else:self.bleed_opts.pack_forget()

    def _on_size(self, e=None):
        sel = self.size_var.get()
        if sel == "Custom…":
            self._open_custom_size_dialog()
        elif sel in self.custom_sizes:
            w, h = self.custom_sizes[sel]
            self.custom_size_lbl.config(text=f"  {w:g} × {h:g} mm  (saved)")
            self.custom_size_lbl.pack(fill="x", padx=16, pady=(0, 6))
            self._refresh()
        else:
            self.custom_size_lbl.pack_forget()
            self._refresh()

    def _open_custom_size_dialog(self):
        popup = tk.Toplevel(self)
        popup.title("Custom paper size")
        popup.configure(bg=BG)
        popup.resizable(False, False)
        popup.grab_set()
        self.update_idletasks()
        px = self.winfo_rootx() + self.winfo_width()//2 - 170
        py = self.winfo_rooty() + self.winfo_height()//2 - 130
        popup.geometry(f"340x260+{px}+{py}")

        tk.Label(popup, text="Custom paper size", bg=BG, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).pack(pady=(18, 10))

        fields = tk.Frame(popup, bg=BG)
        fields.pack(pady=4)
        tk.Label(fields, text="Width (mm)", bg=BG, fg=MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=0, padx=10, sticky="w")
        tk.Label(fields, text="Height (mm)", bg=BG, fg=MUTED,
                 font=("Segoe UI", 9)).grid(row=0, column=1, padx=10, sticky="w")
        e_w = tk.Entry(fields, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                       relief="flat", font=("Segoe UI", 11), width=8)
        e_w.insert(0, str(int(getattr(self, "custom_w_mm", 100))))
        e_w.grid(row=1, column=0, padx=10, pady=4)
        e_h = tk.Entry(fields, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                       relief="flat", font=("Segoe UI", 11), width=8)
        e_h.insert(0, str(int(getattr(self, "custom_h_mm", 100))))
        e_h.grid(row=1, column=1, padx=10, pady=4)
        e_w.focus_set(); e_w.select_range(0, "end")

        save_var = tk.BooleanVar(value=False)
        save_chk = tk.Checkbutton(popup, text="Save this size for later", variable=save_var,
            bg=BG, fg=TEXT, selectcolor=SURFACE2, activebackground=BG,
            font=("Segoe UI", 9), command=lambda: toggle_name())
        save_chk.pack(pady=(8,2))

        name_frame = tk.Frame(popup, bg=BG)
        e_name = tk.Entry(name_frame, bg=SURFACE2, fg=TEXT, insertbackground=TEXT,
                          relief="flat", font=("Segoe UI", 10), width=22)
        e_name.pack(padx=10)

        def toggle_name():
            if save_var.get():
                name_frame.pack(pady=(0,6))
                e_name.focus_set()
            else:
                name_frame.pack_forget()

        def confirm(event=None):
            try:
                w = float(e_w.get().replace(",", "."))
                h = float(e_h.get().replace(",", "."))
                if w <= 0 or h <= 0: raise ValueError
            except ValueError:
                e_w.config(bg=DANGER); e_h.config(bg=DANGER)
                return
            self.custom_w_mm = w
            self.custom_h_mm = h

            if save_var.get():
                name = e_name.get().strip() or f"{w:g}×{h:g} mm"
                self.custom_sizes[name] = (w, h)
                save_custom_sizes(self.custom_sizes)
                self.size_cb.config(values=self._all_size_names())
                self.size_var.set(name)
                self.custom_size_lbl.config(text=f"  {w:g} × {h:g} mm  (saved)")
            else:
                self.size_var.set("Custom…")
                self.custom_size_lbl.config(text=f"  {w:g} × {h:g} mm")

            self.custom_size_lbl.pack(fill="x", padx=16, pady=(0, 6))
            popup.destroy()
            self._refresh()

        def cancel():
            if self.size_var.get() == "Custom…" and not self.custom_size_lbl.cget("text"):
                self.size_var.set("A4")
            popup.destroy()

        btn_row = tk.Frame(popup, bg=BG)
        btn_row.pack(pady=10)
        styled_btn(btn_row, "OK", confirm).pack(side="left", padx=6)
        styled_btn(btn_row, "Cancel", cancel, style="secondary").pack(side="left", padx=6)
        popup.bind("<Return>", confirm)
        popup.bind("<Escape>", lambda e: cancel())
        popup.protocol("WM_DELETE_WINDOW", cancel)

    def _open_manage_sizes(self):
        popup = tk.Toplevel(self)
        popup.title("Manage saved sizes")
        popup.configure(bg=BG)
        popup.resizable(False, False)
        popup.grab_set()
        self.update_idletasks()
        px = self.winfo_rootx() + self.winfo_width()//2 - 170
        py = self.winfo_rooty() + self.winfo_height()//2 - 150
        popup.geometry(f"340x320+{px}+{py}")

        tk.Label(popup, text="Saved custom sizes", bg=BG, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).pack(pady=(18, 10))

        list_frame = tk.Frame(popup, bg=SURFACE2)
        list_frame.pack(fill="both", expand=True, padx=16, pady=(0,10))

        def render_list():
            for w in list_frame.winfo_children(): w.destroy()
            if not self.custom_sizes:
                tk.Label(list_frame, text="No saved sizes yet", bg=SURFACE2, fg=MUTED,
                         font=("Segoe UI", 9)).pack(pady=20)
                return
            for name, (w_mm, h_mm) in sorted(self.custom_sizes.items()):
                row = tk.Frame(list_frame, bg=SURFACE2)
                row.pack(fill="x", padx=8, pady=3)
                tk.Label(row, text=name, bg=SURFACE2, fg=TEXT,
                         font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left")
                tk.Label(row, text=f"{w_mm:g}×{h_mm:g} mm", bg=SURFACE2, fg=MUTED,
                         font=("Segoe UI", 8), anchor="w").pack(side="left", padx=8)
                tk.Button(row, text="✕", bg=SURFACE2, fg=MUTED, relief="flat",
                         cursor="hand2", font=("Segoe UI", 9),
                         activebackground=DANGER, activeforeground=TEXT,
                         command=lambda n=name: delete_size(n)).pack(side="right")

        def delete_size(name):
            del self.custom_sizes[name]
            save_custom_sizes(self.custom_sizes)
            self.size_cb.config(values=self._all_size_names())
            if self.size_var.get() == name:
                self.size_var.set("A4")
                self.custom_size_lbl.pack_forget()
                self._refresh()
            render_list()

        render_list()
        styled_btn(popup, "Close", popup.destroy, style="secondary").pack(pady=(0,16))

    def _get_target_px(self):
        dpi = self.dpi_var.get()
        key = self.size_var.get()
        if key == "Custom…":
            w, h = self.custom_w_mm, self.custom_h_mm
        elif key in self.custom_sizes:
            w, h = self.custom_sizes[key]
        else:
            w, h = PAPER_SIZES_MM[key]
        if self.orient_var.get()=="Landscape": w,h = max(w,h),min(w,h)
        else: w,h = min(w,h),max(w,h)
        return mm_to_px(w,dpi), mm_to_px(h,dpi)

    def _show_empty(self):
        for w in self.grid_f.winfo_children(): w.destroy()
        tk.Label(self.grid_f, text="No images loaded\n\nDrop files here or click 'Add Images'",
                 bg=BG, fg=MUTED, font=("Segoe UI",13)).pack(pady=60)

    def _add_files(self):
        exts = " ".join(f"*{e}" for e in IMAGE_EXTS)
        paths = filedialog.askopenfilenames(
            filetypes=[("Images",exts),("All","*.*")])
        self._drop_files(list(paths))

    def _drop_files(self, paths):
        for p in paths:
            if p not in self.files: self.files.append(p)
        self._rebuild()

    def _clear(self):
        self.files.clear(); self.canvases.clear(); self._show_empty()

    def _refresh(self): self._rebuild()

    def _remove(self, path):
        if path in self.files: self.files.remove(path)
        self._rebuild()

    def _rebuild(self):
        for w in self.grid_f.winfo_children(): w.destroy()
        self.canvases.clear()
        if not self.files: self._show_empty(); return
        tw, th = self._get_target_px()

        # Fixed column tiers: 3 columns on smaller windows, 4 on larger ones.
        avail_w = max(self.scroll_c.winfo_width(), 300)
        cols = 4 if avail_w >= self.COLS_4_THRESHOLD else 3
        cell_pad = 32  # cell internal padx*2 + grid padx*2 (approx)
        thumb = max(160, min(320, (avail_w // cols) - cell_pad))
        thumb_size = (thumb, thumb)
        self._current_cols = cols

        for i,path in enumerate(self.files):
            r,c = divmod(i,cols)
            cell = tk.Frame(self.grid_f, bg=SURFACE2, padx=8, pady=8)
            cell.grid(row=r,column=c,padx=8,pady=8,sticky="n")
            try:
                cc = CropCanvas(cell, path, tw, th, size=thumb_size)
                cc.pack()
                self.canvases.append((path,cc))
                name = os.path.basename(path)
                if len(name)>24: name=name[:11]+"…"+name[-11:]
                tk.Label(cell,text=name,bg=SURFACE2,fg=MUTED,font=("Segoe UI",8)).pack(pady=(4,0))

                action_row = tk.Frame(cell, bg=SURFACE2)
                action_row.pack(pady=(4,0), fill="x")
                for text, command in (
                    ("⟲ Left", lambda c=cc: c.rotate_manual(90)),
                    ("Right ⟳", lambda c=cc: c.rotate_manual(-90)),
                    ("✕ Remove", lambda p=path: self._remove(p)),
                ):
                    tk.Button(action_row, text=text, command=command,
                        bg=SURFACE, fg=TEXT, relief="flat", cursor="hand2",
                        font=("Segoe UI",7,"bold"), padx=4, pady=3,
                        activebackground=ACCENT2, activeforeground=TEXT
                    ).pack(side="left", fill="x", expand=True, padx=1)
            except Exception as ex:
                tk.Label(cell,text=f"⚠ {os.path.basename(path)}\n{str(ex)[:50]}",
                         bg=SURFACE2,fg=DANGER,font=("Segoe UI",8)).pack()

    def _toggle_dest_opts(self):
        mode = self.dest_var.get()
        self.overwrite_frame.pack_forget()
        self.subfolder_frame.pack_forget()
        if mode == "same":
            self.overwrite_frame.pack(fill="x")
        elif mode == "subfolder":
            self.subfolder_frame.pack(fill="x")

    def _export(self):
        if not self.canvases: messagebox.showwarning("Nothing to export","Add some images first."); return
        mode = self.dest_var.get()

        if mode == "choose":
            out = filedialog.askdirectory(title="Output folder")
            if not out: return
            dest_resolver = lambda src_path: out

        elif mode == "same":
            if self.overwrite_var.get():
                if not messagebox.askyesno("Replace originals?",
                    "This will overwrite your original image files with the resized "
                    "versions. This cannot be undone.\n\nContinue?"):
                    return
            dest_resolver = lambda src_path: os.path.dirname(src_path)

        elif mode == "subfolder":
            name = self.subfolder_name.get().strip() or "resized"
            def dest_resolver(src_path, _name=name):
                d = os.path.join(os.path.dirname(src_path), _name)
                os.makedirs(d, exist_ok=True)
                return d
        else:
            dest_resolver = lambda src_path: os.path.dirname(src_path)

        self.exp_btn.config(state="disabled")
        threading.Thread(target=self._do_export, args=(dest_resolver,), daemon=True).start()

    def _fit_or_stretch(self, image, size):
        """Resize while preserving proportions unless Stretch to fit is enabled."""
        if self.stretch_var.get():
            return image.resize(size, Image.Resampling.LANCZOS)
        return ImageOps.fit(
            image,
            size,
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5),
        )

    def _do_export(self, dest_resolver):
        dpi = self.dpi_var.get()
        tw, th = self._get_target_px()
        fmt = self.fmt_var.get()
        quality = int(self.q_var.get())
        ext = fmt.lower().replace("jpeg","jpg")
        bleed = self.bleed_var.get()
        try: bleed_mm = float(self.bleed_mm.get())
        except: bleed_mm = 3.0
        marks=self.marks_var.get()
        bleed_mode=self.bleed_mode_var.get(); bleed_color=self.bleed_color
        replace_originals = (self.dest_var.get() == "same" and self.overwrite_var.get())
        done = errors = 0
        last_out_dir = ""

        if fmt == "PDF":
            pages = []
            for path, cc in self.canvases:
                try:
                    src_ext = os.path.splitext(path)[1].lower()
                    if src_ext == ".pdf":
                        src_pages = open_pdf_pages(path, dpi=dpi)
                        img = src_pages[0].convert("RGB") if src_pages else Image.new("RGB", (tw, th))
                    else:
                        img = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
                    if cc.total_rotation:
                        img = img.rotate(cc.total_rotation, expand=True)
                    iw, ih = img.size
                    fl, ft, fr, fb = cc.get_crop_fractions()
                    cropped = self._fit_or_stretch(
                        img.crop((int(fl*iw), int(ft*ih), int(fr*iw), int(fb*ih))),
                        (tw, th),
                    )
                    if self.gray_var.get(): cropped = ImageOps.grayscale(cropped).convert("RGB")
                    cropped = apply_resizer_finishing(cropped,dpi,bleed,bleed_mm,marks,bleed_mode,bleed_color)
                    pages.append(cropped.convert("RGB"))
                    done += 1
                except Exception:
                    errors += 1
            if pages:
                out_dir = dest_resolver(self.canvases[0][0])
                last_out_dir = out_dir
                out_path = _unique(os.path.join(out_dir, "resized_images.pdf"))
                pages[0].save(out_path, "PDF", save_all=True, append_images=pages[1:], resolution=dpi)
            msg = f"✓ {done} exported" + (f"  ⚠ {errors} failed" if errors else "")
            self.after(0, lambda: self.status.config(text=msg, fg=SUCCESS if not errors else WARNING))
            self.after(0, lambda: self.exp_btn.config(state="normal"))
            if pages:
                self.after(0, lambda: messagebox.showinfo("Done", f"Exported {done} image(s) to one PDF:\n{out_path}"))
            return

        for path, cc in self.canvases:
            try:
                # Reload the original at full resolution (preview was downscaled
                # for performance) and apply the exact same rotation + crop fractions.
                src_ext = os.path.splitext(path)[1].lower()
                if src_ext == ".pdf":
                    pages = open_pdf_pages(path, dpi=dpi)
                    img = pages[0].convert("RGB") if pages else Image.new("RGB",(tw,th))
                else:
                    raw = Image.open(path)
                    raw = ImageOps.exif_transpose(raw)
                    img = raw.convert("RGB")
                if cc.total_rotation:
                    img = img.rotate(cc.total_rotation, expand=True)

                iw, ih = img.size
                fl, ft, fr, fb = cc.get_crop_fractions()
                box = (int(fl*iw), int(ft*ih), int(fr*iw), int(fb*ih))
                cropped = self._fit_or_stretch(img.crop(box), (tw, th))
                if self.gray_var.get(): cropped = ImageOps.grayscale(cropped).convert("RGB")
                cropped = apply_resizer_finishing(cropped,dpi,bleed,bleed_mm,marks,bleed_mode,bleed_color)

                out_dir = dest_resolver(path)
                last_out_dir = out_dir
                base = os.path.splitext(os.path.basename(path))[0]

                if replace_originals:
                    # Overwrite the original file in place, keeping its exact name + extension
                    out_path = path
                    save_fmt = fmt
                    # If output format differs from the original extension, keep the new
                    # extension instead of silently mismatching file content vs. name.
                    if src_ext.lstrip(".").replace("jpg","jpeg") != fmt.lower():
                        out_path = os.path.join(out_dir, f"{base}.{ext}")
                else:
                    sfx = "_bleed" if bleed else "_resized"
                    out_path = os.path.join(out_dir, f"{base}{sfx}.{ext}")
                    n=1
                    while os.path.exists(out_path):
                        out_path = os.path.join(out_dir, f"{base}{sfx}_{n}.{ext}"); n+=1

                kw = {"dpi":(dpi,dpi)}
                if fmt=="JPEG": kw["quality"]=quality
                cropped.save(out_path, fmt, **kw)

                # If we replaced in place but had to change extension, remove the old file
                if replace_originals and out_path != path and os.path.exists(path):
                    try: os.remove(path)
                    except Exception: pass

                done+=1
            except Exception as ex:
                errors+=1
        msg = f"✓ {done} exported" + (f"  ⚠ {errors} failed" if errors else "")
        col = SUCCESS if not errors else WARNING
        self.after(0, lambda: self.status.config(text=msg, fg=col))
        self.after(0, lambda: self.exp_btn.config(state="normal"))
        if done:
            dest_desc = "original location(s)" if self.dest_var.get() in ("same","subfolder") else last_out_dir
            self.after(0, lambda: messagebox.showinfo("Done",
                f"Exported {done} image(s) to:\n{dest_desc}"))

def _pages_member_candidates(names):
    normalised=[name.replace("\\","/") for name in names]
    preferred=[
        "QuickLook/Preview.pdf","quicklook/preview.pdf","Preview.pdf","preview.pdf",
        "QuickLook/Preview.jpg","QuickLook/Preview.jpeg","QuickLook/Preview.png",
        "preview.jpg","preview.jpeg","preview.png",
    ]
    result=[]
    lower_map={name.lower():name for name in normalised}
    for wanted in preferred:
        actual=lower_map.get(wanted.lower())
        if actual and actual not in result: result.append(actual)
    # Some Pages versions use differently named preview files.
    for name in normalised:
        lower=name.lower()
        if (lower.endswith(".pdf") or lower.endswith((".jpg",".jpeg",".png"))) and ("preview" in lower or "quicklook" in lower):
            if name not in result: result.append(name)
    return result

def extract_pages_preview(path):
    """Return (list[PIL.Image], warning). Uses the best preview embedded by Pages."""
    if not zipfile.is_zipfile(path):
        raise ValueError("Šis .pages failas neturi nuskaitomo peržiūros paketo.")
    with zipfile.ZipFile(path,"r") as package:
        candidates=_pages_member_candidates(package.namelist())
        if not candidates:
            raise ValueError("Faile nerasta įterpto PDF ar vaizdo peržiūros.")
        selected=candidates[0]
        data=package.read(selected)
        if selected.lower().endswith(".pdf"):
            doc=fitz.open(stream=data,filetype="pdf")
            images=[]
            for page in doc:
                pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
                images.append(Image.frombytes("RGB",[pix.width,pix.height],pix.samples))
            doc.close()
            return images,"Iš .pages failo panaudota įterpta PDF peržiūra. Patikrinkite, ar joje yra visas dokumentas."
        image=Image.open(io.BytesIO(data)); image.load()
        return [image.convert("RGB")],"Iš .pages failo panaudota vaizdo peržiūra. Ji gali būti mažesnės raiškos arba nepilna."

# ── Converter Tab ─────────────────────────────────────────────────────────────
class ConverterTab(tk.Frame, DropMixin):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.files = []
        self._build()

    def _build(self):
        left = tk.Frame(self, bg=SURFACE, width=250)
        left.pack(side="left", fill="y", padx=(0,1))
        left.pack_propagate(False)

        lbl(left,"CONVERT SETTINGS",8,MUTED,True).pack(pady=(18,4),padx=16,anchor="w")

        lbl(left,"Output format").pack(padx=16,anchor="w")
        self.fmt_var = tk.StringVar(value="PNG")
        self.fmt_cb = ttk.Combobox(left, textvariable=self.fmt_var,
            values=CONV_FORMATS, state="readonly", width=20)
        self.fmt_cb.pack(fill="x",padx=16,pady=(2,8))
        self.fmt_cb.bind("<<ComboboxSelected>>", self._toggle_opts)

        # JPEG quality
        self.jq_lbl = lbl(left,"JPEG Quality: 90"); self.jq_lbl.pack_forget()
        self.jq_var = tk.IntVar(value=90)
        self.jq_slider = ttk.Scale(left, from_=10,to=100, variable=self.jq_var,
            orient="horizontal",
            command=lambda v: self.jq_lbl.config(text=f"JPEG Quality: {int(float(v))}"))
        self.jq_slider.pack_forget()

        # WebP lossless
        self.webp_frame = tk.Frame(left,bg=SURFACE)
        self.webp_lossless = tk.BooleanVar(value=False)
        tk.Checkbutton(self.webp_frame, text="Lossless WebP",
            variable=self.webp_lossless, bg=SURFACE, fg=TEXT, selectcolor=SURFACE2,
            activebackground=SURFACE, font=("Segoe UI",9)).pack(anchor="w",padx=4)
        self.webp_frame.pack_forget()

        # PDF raster DPI (for PDF→image)
        lbl(left,"Raster DPI (PDF input)").pack(padx=16,anchor="w",pady=(6,0))
        self.rdpi_var = tk.IntVar(value=150)
        ttk.Combobox(left, textvariable=self.rdpi_var, values=DPI_OPTIONS,
                     state="readonly", width=20).pack(fill="x",padx=16,pady=(2,8))

        # BG colour for transparency
        lbl(left,"BG colour (flatten transparency)").pack(padx=16,anchor="w",pady=(4,0))
        bgr = tk.Frame(left,bg=SURFACE); bgr.pack(fill="x",padx=16,pady=(2,8))
        self.bg_var = tk.StringVar(value="#FFFFFF")
        entry(bgr, var=self.bg_var, w=10).pack(side="left")
        lbl(bgr,"(hex)  e.g. #FFFFFF",8,MUTED).pack(side="left",padx=4)

        sep(left)
        build_output_location(left, self, "converted")
        sep(left)

        styled_btn(left,"➕  Add Files",self._add_files).pack(fill="x",padx=16,pady=3)
        styled_btn(left,"Atidaryti Pages for iCloud",lambda:webbrowser.open("https://www.icloud.com/pages"),style="secondary").pack(fill="x",padx=16,pady=3)
        styled_btn(left,"🗑   Clear All",self._clear,style="secondary").pack(fill="x",padx=16,pady=3)
        sep(left)
        self.conv_btn = styled_btn(left,"🔄  Convert All",self._convert_all,style="success")
        self.conv_btn.pack(fill="x",padx=16,pady=3)
        self.status = lbl(left,"",8,MUTED); self.status.pack(padx=16,pady=6,anchor="w")

        # Right
        right = tk.Frame(self, bg=BG)
        right.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        self.drop_zone = tk.Label(right,
            text="⬇  Drop files here  (images, PDFs, SVGs, Pages…)",
            bg=SURFACE2, fg=MUTED, font=("Segoe UI",10), pady=10, cursor="hand2")
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Button-1>", lambda e: self._add_files())
        self.setup_drop(self.drop_zone, self._drop_files)

        lbl_f = tk.Frame(right,bg=SURFACE2)
        lbl_f.pack(fill="both",expand=True,pady=6)

        self.listbox = tk.Listbox(lbl_f, bg=SURFACE2, fg=TEXT, relief="flat",
            selectbackground=ACCENT, selectforeground=TEXT,
            font=("Segoe UI",10), activestyle="none",
            highlightthickness=0, borderwidth=0)
        vsb = ttk.Scrollbar(lbl_f,orient="vertical",command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right",fill="y")
        self.listbox.pack(fill="both",expand=True,padx=2,pady=2)
        self.setup_drop(self.listbox, self._drop_files)

        btns = tk.Frame(right,bg=BG); btns.pack(fill="x",pady=4)
        styled_btn(btns,"Remove selected",self._remove_sel,style="secondary").pack(side="left")
        self.info_lbl = tk.Label(right,text="0 files queued",bg=BG,fg=MUTED,
                                  font=("Segoe UI",9))
        self.info_lbl.pack(anchor="w",pady=(0,4))
        self.progress = ttk.Progressbar(right,orient="horizontal",mode="determinate")
        self.progress.pack(fill="x",pady=4)

    def _toggle_dest_opts(self):
        toggle_output_location(self)

    def _toggle_opts(self, e=None):
        fmt = self.fmt_var.get()
        if fmt=="JPEG":
            self.jq_lbl.pack(padx=16,anchor="w"); self.jq_slider.pack(fill="x",padx=16,pady=(0,8))
        else:
            self.jq_lbl.pack_forget(); self.jq_slider.pack_forget()
        if fmt=="WEBP": self.webp_frame.pack(fill="x",padx=16,pady=(0,8))
        else: self.webp_frame.pack_forget()

    def _add_files(self):
        exts = " ".join(f"*{e}" for e in ALL_EXTS)
        paths = filedialog.askopenfilenames(
            filetypes=[("Supported files",exts),("All","*.*")])
        self._drop_files(list(paths))

    def _drop_files(self, paths):
        for p in paths:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end","  "+os.path.basename(p))
        self._update_info()

    def _clear(self):
        self.files.clear(); self.listbox.delete(0,"end"); self._update_info()

    def _remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            self.files.pop(i); self.listbox.delete(i)
        self._update_info()

    def _update_info(self):
        n=len(self.files)
        self.info_lbl.config(text=f"{n} file{'s' if n!=1 else ''} queued")

    def _convert_all(self):
        if not self.files: messagebox.showwarning("Nothing","Add files first."); return
        dest_resolver, replace_originals = get_output_resolver(self, "Output folder")
        if dest_resolver is None: return
        self.conv_btn.config(state="disabled")
        self.progress["value"]=0; self.progress["maximum"]=len(self.files)
        threading.Thread(target=self._do_convert,
            args=(dest_resolver, replace_originals), daemon=True).start()

    def _do_convert(self, dest_resolver, replace_originals):
        fmt = self.fmt_var.get()
        quality = int(self.jq_var.get())
        lossless = self.webp_lossless.get()
        rdpi = self.rdpi_var.get()
        bg_color = self.bg_var.get().strip() or "#FFFFFF"
        ext = fmt.lower().replace("jpeg","jpg")
        done=errors=0; msgs=[]

        def bg_rgb():
            try: return tuple(int(bg_color.lstrip("#")[j:j+2],16) for j in (0,2,4))
            except: return (255,255,255)

        if fmt == "PDF":
            pages = []
            for i, path in enumerate(self.files):
                try:
                    source_ext=os.path.splitext(path)[1].lower()
                    if source_ext == ".pdf":
                        pages.extend(img.convert("RGB") for img in open_pdf_pages(path, dpi=rdpi))
                    elif source_ext == ".pages":
                        extracted, warning=extract_pages_preview(path)
                        pages.extend(img.convert("RGB") for img in extracted)
                        msgs.append(f"{os.path.basename(path)}: {warning}")
                    else:
                        pages.append(_prepare_img(Image.open(path), "PDF", bg_rgb()).convert("RGB"))
                    done += 1
                except Exception as ex:
                    errors += 1; msgs.append(f"{os.path.basename(path)}: {ex}")
                self.after(0, lambda v=i+1: self.progress.config(value=v))
            if pages:
                out_dir = dest_resolver(self.files[0])
                out_path = _unique(os.path.join(out_dir, "converted_files.pdf"))
                pages[0].save(out_path, "PDF", save_all=True, append_images=pages[1:], resolution=rdpi)
            msg = f"✓ {done} converted" + (f"  ⚠ {errors} failed" if errors else "")
            self.after(0, lambda: self.status.config(text=msg, fg=SUCCESS if not errors else WARNING))
            self.after(0, lambda: self.conv_btn.config(state="normal"))
            if pages:
                self.after(0, lambda: messagebox.showinfo("Done", f"Combined {done} file(s) into:\n{out_path}"))
            return

        for i,path in enumerate(self.files):
            src_ext = os.path.splitext(path)[1].lower()
            base = os.path.splitext(os.path.basename(path))[0]
            try:
                # ── PDF / Pages input → rasterise each page ──
                if src_ext in (".pdf",".pages"):
                    if src_ext == ".pdf": pages = open_pdf_pages(path, dpi=rdpi)
                    else:
                        pages, warning = extract_pages_preview(path)
                        msgs.append(f"{os.path.basename(path)}: {warning}")
                    for p_idx, page_img in enumerate(pages):
                        out_base = f"{base}_p{p_idx+1:03d}" if len(pages)>1 else base
                        out_dir = dest_resolver(path)
                        candidate = os.path.join(out_dir, f"{out_base}.{ext}")
                        out_path = candidate if replace_originals else _unique(candidate)
                        page_img = _prepare_img(page_img, fmt, bg_rgb())
                        _save(page_img, out_path, fmt, quality, lossless)
                    if replace_originals:
                        try: os.remove(path)
                        except OSError: pass
                    done+=1
                else:
                    img = Image.open(path)
                    img = _prepare_img(img, fmt, bg_rgb())
                    out_dir = dest_resolver(path)
                    candidate = os.path.join(out_dir, f"{base}.{ext}")
                    out_path = candidate if replace_originals else _unique(candidate)
                    _save(img, out_path, fmt, quality, lossless)
                    if replace_originals and os.path.abspath(out_path) != os.path.abspath(path):
                        try: os.remove(path)
                        except OSError: pass
                    done+=1
            except Exception as ex:
                errors+=1; msgs.append(f"{os.path.basename(path)}: {ex}")
            self.after(0, lambda v=i+1: setattr(self.progress,"value",v) or
                       self.progress.config(value=v))

        msg = f"✓ {done} converted" + (f"  ⚠ {errors} failed" if errors else "")
        col = SUCCESS if not errors else WARNING
        self.after(0, lambda: self.status.config(text=msg,fg=col))
        self.after(0, lambda: self.conv_btn.config(state="normal"))
        if done:
            detail = f"Converted {done} file(s) to:\n{out_dir}"
            if msgs: detail += "\n\nErrors:\n"+"\n".join(msgs[:5])
            self.after(0, lambda: messagebox.showinfo("Done",detail))

def _prepare_img(img, fmt, bg_rgb):
    if fmt in ("JPEG","BMP") and img.mode in ("RGBA","LA","P"):
        if img.mode=="P": img=img.convert("RGBA")
        bg = Image.new("RGB", img.size, bg_rgb)
        bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA","LA") else None)
        return bg
    if fmt=="ICO": return img.convert("RGBA")
    if fmt in ("JPEG","PDF","BMP","PPM","TGA"): return img.convert("RGB")
    return img

def _save(img, path, fmt, quality, lossless):
    kw = {}
    if fmt=="JPEG": kw["quality"]=quality
    elif fmt=="WEBP": kw["lossless"]=lossless; kw["quality"]=quality
    elif fmt=="ICO": kw["sizes"]=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)]
    elif fmt=="GIF": img=img.convert("P",palette=Image.ADAPTIVE)
    img.save(path, fmt, **kw)

def _unique(path):
    if not os.path.exists(path): return path
    base,ext=os.path.splitext(path); n=1
    while os.path.exists(f"{base}_{n}{ext}"): n+=1
    return f"{base}_{n}{ext}"

# ── PDF Tools Tab ─────────────────────────────────────────────────────────────
class PdfToolsTab(tk.Frame, DropMixin):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.files = []
        self._build()

    def _build(self):
        left = tk.Frame(self, bg=SURFACE, width=250)
        left.pack(side="left", fill="y", padx=(0,1))
        left.pack_propagate(False)

        lbl(left,"PDF OPTIMISER",8,MUTED,True).pack(pady=(18,4),padx=16,anchor="w")

        lbl(left,"Rasterise DPI").pack(padx=16,anchor="w")
        self.dpi_var = tk.IntVar(value=150)
        ttk.Combobox(left, textvariable=self.dpi_var, values=DPI_OPTIONS,
                     state="readonly", width=20).pack(fill="x",padx=16,pady=(2,8))

        lbl(left,"JPEG quality (re-compress)").pack(padx=16,anchor="w")
        self.q_var = tk.IntVar(value=80)
        self.q_lbl = lbl(left,"Quality: 80")
        self.q_lbl.pack(padx=16,anchor="w")
        ttk.Scale(left, from_=10, to=100, variable=self.q_var, orient="horizontal",
            command=lambda v: self.q_lbl.config(text=f"Quality: {int(float(v))}")
        ).pack(fill="x",padx=16,pady=(0,8))

        sep(left)
        lbl(left,"OPERATIONS",8,MUTED,True).pack(pady=(4,4),padx=16,anchor="w")

        self.op_flatten   = tk.BooleanVar(value=True)
        self.op_compress  = tk.BooleanVar(value=True)
        self.op_rasterise = tk.BooleanVar(value=False)

        for var, text, tip in [
            (self.op_flatten,   "Flatten",   "Merge all layers into one"),
            (self.op_compress,  "Compress",  "Re-save with deflate compression"),
            (self.op_rasterise, "Rasterise", "Convert every page to image (max compatibility)"),
        ]:
            row = tk.Frame(left,bg=SURFACE); row.pack(fill="x",padx=16,pady=2)
            tk.Checkbutton(row,text=text,variable=var,bg=SURFACE,fg=TEXT,
                selectcolor=SURFACE2,activebackground=SURFACE,
                font=("Segoe UI",9,"bold")).pack(side="left")
            lbl(row,f"  {tip}",8,MUTED).pack(side="left")

        sep(left)
        build_output_location(left, self, "optimised")
        sep(left)
        styled_btn(left,"➕  Add PDFs",self._add_files).pack(fill="x",padx=16,pady=3)
        styled_btn(left,"🗑   Clear All",self._clear,style="secondary").pack(fill="x",padx=16,pady=3)
        sep(left)
        self.run_btn = styled_btn(left,"⚙  Process All",self._process,style="success")
        self.run_btn.pack(fill="x",padx=16,pady=3)
        self.status = lbl(left,"",8,MUTED); self.status.pack(padx=16,pady=6,anchor="w")

        # Right
        right = tk.Frame(self,bg=BG)
        right.pack(side="left",fill="both",expand=True,padx=12,pady=12)

        self.drop_zone = tk.Label(right,
            text="⬇  Drop PDF files here",
            bg=SURFACE2, fg=MUTED, font=("Segoe UI",10), pady=10, cursor="hand2")
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Button-1>", lambda e: self._add_files())
        self.setup_drop(self.drop_zone, self._drop_files, {".pdf"})

        lf = tk.Frame(right,bg=SURFACE2); lf.pack(fill="both",expand=True,pady=6)
        self.listbox = tk.Listbox(lf, bg=SURFACE2, fg=TEXT, relief="flat",
            selectbackground=ACCENT, font=("Segoe UI",10), activestyle="none",
            highlightthickness=0, borderwidth=0)
        vsb = ttk.Scrollbar(lf,orient="vertical",command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right",fill="y"); self.listbox.pack(fill="both",expand=True,padx=2,pady=2)
        self.setup_drop(self.listbox, self._drop_files, {".pdf"})

        btns = tk.Frame(right,bg=BG); btns.pack(fill="x",pady=4)
        styled_btn(btns,"Remove selected",self._remove_sel,style="secondary").pack(side="left")
        self.info_lbl = tk.Label(right,text="0 PDFs queued",bg=BG,fg=MUTED,font=("Segoe UI",9))
        self.info_lbl.pack(anchor="w",pady=(0,4))
        self.progress = ttk.Progressbar(right,orient="horizontal",mode="determinate")
        self.progress.pack(fill="x",pady=4)

        # Size summary
        self.size_lbl = tk.Label(right,text="",bg=BG,fg=MUTED,font=("Segoe UI",8))
        self.size_lbl.pack(anchor="w")

    def _toggle_dest_opts(self):
        toggle_output_location(self)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("PDF files","*.pdf"),("All","*.*")])
        self._drop_files(list(paths))

    def _drop_files(self, paths):
        for p in paths:
            if p not in self.files and p.lower().endswith(".pdf"):
                self.files.append(p)
                self.listbox.insert("end","  "+os.path.basename(p))
        self._update_info()

    def _clear(self):
        self.files.clear(); self.listbox.delete(0,"end"); self._update_info()

    def _remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            self.files.pop(i); self.listbox.delete(i)
        self._update_info()

    def _update_info(self):
        n=len(self.files)
        total = sum(os.path.getsize(f) for f in self.files if os.path.exists(f))
        size_str = f"  ({total/1024/1024:.1f} MB total)" if total else ""
        self.info_lbl.config(text=f"{n} PDF{'s' if n!=1 else ''} queued{size_str}")

    def _process(self):
        if not self.files: messagebox.showwarning("Nothing","Add PDF files first."); return
        dest_resolver, replace_originals = get_output_resolver(self, "Output folder")
        if dest_resolver is None: return
        self.run_btn.config(state="disabled")
        self.progress["value"]=0; self.progress["maximum"]=len(self.files)
        threading.Thread(target=self._do_process,
            args=(dest_resolver, replace_originals), daemon=True).start()

    def _do_process(self, dest_resolver, replace_originals):
        dpi = self.dpi_var.get()
        quality = int(self.q_var.get())
        flatten   = self.op_flatten.get()
        compress  = self.op_compress.get()
        rasterise = self.op_rasterise.get()
        done=errors=0; saved_bytes=0

        for i,path in enumerate(self.files):
            orig_size = os.path.getsize(path)
            base = os.path.splitext(os.path.basename(path))[0]
            try:
                doc = fitz.open(path)
                out_dir = dest_resolver(path)
                if replace_originals:
                    out_path = _unique(os.path.join(out_dir, f".{base}_optimised_tmp.pdf"))
                else:
                    out_path = _unique(os.path.join(out_dir, f"{base}_optimised.pdf"))

                if rasterise:
                    # Convert each page to image, rebuild PDF
                    new_doc = fitz.open()
                    mat = fitz.Matrix(dpi/72, dpi/72)
                    for page in doc:
                        pix = page.get_pixmap(matrix=mat, alpha=False)
                        # Save as JPEG bytes then insert
                        img_pil = Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                        buf = io.BytesIO()
                        img_pil.save(buf,"JPEG",quality=quality)
                        buf.seek(0)
                        img_doc = fitz.open("pdf", fitz.open(stream=buf.read(),filetype="jpeg")
                                            .convert_to_pdf())
                        new_doc.insert_pdf(img_doc)
                    if compress:
                        new_doc.save(out_path, garbage=4, deflate=True, clean=True)
                    else:
                        new_doc.save(out_path)
                    new_doc.close()
                else:
                    # Flatten annotations + form fields
                    if flatten:
                        for page in doc:
                            page.clean_contents()
                            for annot in page.annots():
                                annot.set_flags(fitz.ANNOT_HIDDEN)

                    if compress:
                        doc.save(out_path, garbage=4, deflate=True,
                                 deflate_images=True, deflate_fonts=True, clean=True)
                    else:
                        doc.save(out_path, garbage=1, clean=True)

                doc.close()
                if replace_originals:
                    os.replace(out_path, path)
                    out_path = path
                new_size = os.path.getsize(out_path)
                saved_bytes += max(0, orig_size - new_size)
                done+=1
            except Exception as ex:
                errors+=1
            self.after(0, lambda v=i+1: self.progress.config(value=v))

        msg = f"✓ {done} processed" + (f"  ⚠ {errors} failed" if errors else "")
        col = SUCCESS if not errors else WARNING
        saved_mb = saved_bytes/1024/1024
        self.after(0, lambda: self.status.config(text=msg,fg=col))
        self.after(0, lambda: self.size_lbl.config(
            text=f"Space saved: {saved_mb:.1f} MB" if saved_mb>0 else ""))
        self.after(0, lambda: self.run_btn.config(state="normal"))
        if done:
            self.after(0, lambda: messagebox.showinfo("Done",
                f"Processed {done} PDF(s) → {out_dir}\nSaved ~{saved_mb:.1f} MB"))


# ── Sticker Cut-line Tab ─────────────────────────────────────────────────────
class StickerCutlineTab(tk.Frame, DropMixin):
    """Create editable vector cut paths from artwork on a light background."""

    SUPPORTED = IMAGE_EXTS | {".pdf"}

    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.source_path = None
        self.source_image = None
        self.analysis_image = None
        self.analysis_contours = []
        self.source_dpi = 300.0
        self.source_page_points = None
        self._preview_photo = None
        self._analysis_job = None
        self._build()

    def _build(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)

        # LEFT: controls
        controls = tk.Frame(self, bg=SURFACE, width=280)
        controls.grid(row=0, column=0, sticky="nsew")
        controls.grid_propagate(False)

        lbl(controls, "VECTOR CUT LINE", 12, TEXT, True).pack(anchor="w", padx=16, pady=(16, 4))
        lbl(
            controls,
            "Finds the outer edge of artwork on a white background\nand exports a real red vector path.",
            8,
            MUTED,
        ).pack(anchor="w", padx=16, pady=(0, 10))

        sep(controls)
        lbl(controls, "EDGE DETECTION", 8, MUTED, True).pack(anchor="w", padx=16, pady=(4, 2))

        self.tolerance_var = tk.IntVar(value=15)
        self.min_area_var = tk.DoubleVar(value=0.05)
        self.offset_var = tk.DoubleVar(value=0.0)
        self.include_holes_var = tk.BooleanVar(value=False)
        self.dpi_var = tk.DoubleVar(value=300.0)
        self.line_width_var = tk.DoubleVar(value=0.25)

        self._add_scale(controls, "White tolerance", self.tolerance_var, 0, 80, 1)
        self._add_scale(controls, "Minimum object area (%)", self.min_area_var, 0, 5, 0.05)
        self._add_scale(controls, "Outline offset (mm)", self.offset_var, -3, 10, 0.1)

        row = tk.Frame(controls, bg=SURFACE)
        row.pack(fill="x", padx=16, pady=(5, 3))
        tk.Checkbutton(
            row,
            text="Include internal white areas",
            variable=self.include_holes_var,
            bg=SURFACE,
            fg=TEXT,
            selectcolor=SURFACE2,
            activebackground=SURFACE,
            activeforeground=TEXT,
            font=("Segoe UI", 9),
            command=self._schedule_analysis,
        ).pack(anchor="w")

        row = tk.Frame(controls, bg=SURFACE)
        row.pack(fill="x", padx=16, pady=3)
        lbl(row, "Source DPI", 8, MUTED).pack(side="left")
        dpi_entry = entry(row, self.dpi_var, 7)
        dpi_entry.pack(side="right")
        dpi_entry.bind("<FocusOut>", lambda _e: self._schedule_analysis())
        dpi_entry.bind("<Return>", lambda _e: self._schedule_analysis())

        sep(controls)
        lbl(controls, "EXPORT", 8, MUTED, True).pack(anchor="w", padx=16, pady=(4, 2))

        self.export_mode_var = tk.StringVar(value="outline")
        for title, value in (
            ("Outline only", "outline"),
            ("Artwork + outline (separate layers)", "artwork"),
        ):
            tk.Radiobutton(
                controls,
                text=title,
                variable=self.export_mode_var,
                value=value,
                bg=SURFACE,
                fg=TEXT,
                selectcolor=SURFACE2,
                activebackground=SURFACE,
                activeforeground=TEXT,
                font=("Segoe UI", 9),
            ).pack(anchor="w", padx=16)

        row = tk.Frame(controls, bg=SURFACE)
        row.pack(fill="x", padx=16, pady=(7, 3))
        lbl(row, "Format", 8, MUTED).pack(side="left")
        self.export_format_var = tk.StringVar(value="PDF")
        self.export_format_cb = ttk.Combobox(
            row,
            textvariable=self.export_format_var,
            values=["PDF", "SVG"],
            state="readonly",
            width=8,
        )
        self.export_format_cb.pack(side="right")

        row = tk.Frame(controls, bg=SURFACE)
        row.pack(fill="x", padx=16, pady=3)
        lbl(row, "Red line width (mm)", 8, MUTED).pack(side="left")
        entry(row, self.line_width_var, 7).pack(side="right")

        self.export_btn = styled_btn(
            controls,
            "Export vector cut line",
            self._export,
            style="success",
        )
        self.export_btn.pack(fill="x", padx=16, pady=(12, 4))
        self.export_btn.config(state="disabled")

        styled_btn(
            controls,
            "Reset settings",
            self._reset_settings,
            style="secondary",
        ).pack(fill="x", padx=16, pady=4)

        self.status_lbl = lbl(controls, "Drop or choose an image/PDF.", 8, MUTED)
        self.status_lbl.pack(anchor="w", padx=16, pady=(8, 12))

        # MIDDLE: large preview
        middle = tk.Frame(self, bg=BG)
        middle.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        middle.grid_rowconfigure(1, weight=1)
        middle.grid_columnconfigure(0, weight=1)
        lbl(middle, "PREVIEW", 9, MUTED, True, bg=BG).grid(row=0, column=0, sticky="w", pady=(0, 5))
        preview_wrap = tk.Frame(middle, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        preview_wrap.grid(row=1, column=0, sticky="nsew")
        preview_wrap.grid_rowconfigure(0, weight=1)
        preview_wrap.grid_columnconfigure(0, weight=1)
        self.preview_canvas = tk.Canvas(preview_wrap, bg="#D8D8D8", highlightthickness=0)
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        self.preview_canvas.bind("<Configure>", lambda _e: self._draw_preview())

        # RIGHT: compact source panel
        files = tk.Frame(self, bg=SURFACE, width=230)
        files.grid(row=0, column=2, sticky="nsew")
        files.grid_propagate(False)
        lbl(files, "SOURCE FILE", 9, MUTED, True).pack(anchor="w", padx=14, pady=(16, 6))
        self.file_name_lbl = lbl(files, "No file selected", 9, TEXT)
        self.file_name_lbl.pack(anchor="w", padx=14, pady=(0, 8))
        self.file_info_lbl = lbl(files, "", 8, MUTED)
        self.file_info_lbl.pack(anchor="w", padx=14, pady=(0, 8))
        styled_btn(files, "Choose file…", self._choose_file, style="secondary").pack(fill="x", padx=14, pady=4)
        styled_btn(files, "Remove", self._clear_file, style="secondary").pack(fill="x", padx=14, pady=4)

        self.drop_zone = tk.Label(
            files,
            text="DROP IMAGE OR PDF HERE",
            bg=SURFACE2,
            fg=MUTED,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            bd=0,
            height=5,
        )
        self.drop_zone.pack(side="bottom", fill="x", padx=14, pady=14)
        self.setup_drop(self.drop_zone, self._add_files, self.SUPPORTED)

    def _add_scale(self, parent, title, variable, start, end, resolution):
        row = tk.Frame(parent, bg=SURFACE)
        row.pack(fill="x", padx=16, pady=(3, 0))
        lbl(row, title, 8, MUTED).pack(anchor="w")
        scale = tk.Scale(
            row,
            from_=start,
            to=end,
            resolution=resolution,
            orient="horizontal",
            variable=variable,
            bg=SURFACE,
            fg=TEXT,
            troughcolor=SURFACE2,
            highlightthickness=0,
            bd=0,
            font=("Segoe UI", 8),
            command=lambda _v: self._schedule_analysis(),
        )
        scale.pack(fill="x")

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Choose artwork",
            filetypes=[
                ("Artwork", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp *.heic *.heif *.pdf"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._load_file(path)

    def _add_files(self, paths):
        if paths:
            self._load_file(paths[0])

    def _clear_file(self):
        self.source_path = None
        self.source_image = None
        self.analysis_image = None
        self.analysis_contours = []
        self.source_page_points = None
        self.file_name_lbl.config(text="No file selected")
        self.file_info_lbl.config(text="")
        self.status_lbl.config(text="Drop or choose an image/PDF.", fg=MUTED)
        self.export_btn.config(state="disabled")
        self.preview_canvas.delete("all")

    def _load_file(self, path):
        try:
            ext = os.path.splitext(path)[1].lower()
            self.source_path = path
            self.source_page_points = None

            if ext == ".pdf":
                doc = fitz.open(path)
                if doc.page_count < 1:
                    doc.close()
                    raise ValueError("PDF has no pages.")
                page = doc[0]
                page_width_pt, page_height_pt = page.rect.width, page.rect.height
                self.source_page_points = (page_width_pt, page_height_pt)
                render_dpi = 180
                pix = page.get_pixmap(matrix=fitz.Matrix(render_dpi / 72, render_dpi / 72), alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                doc.close()
                self.source_dpi = float(render_dpi)
                self.dpi_var.set(render_dpi)
                info = f"PDF page 1 • {page_width_pt:.0f} × {page_height_pt:.0f} pt"
            else:
                raw = ImageOps.exif_transpose(Image.open(path))
                dpi = raw.info.get("dpi", (300, 300))
                try:
                    detected_dpi = float(dpi[0]) if isinstance(dpi, (tuple, list)) else float(dpi)
                    if not math.isfinite(detected_dpi) or detected_dpi < 20:
                        detected_dpi = 300.0
                except Exception:
                    detected_dpi = 300.0
                self.source_dpi = detected_dpi
                self.dpi_var.set(round(detected_dpi, 2))
                if raw.mode in ("RGBA", "LA") or "transparency" in raw.info:
                    rgba = raw.convert("RGBA")
                    bg = Image.new("RGBA", rgba.size, "white")
                    bg.alpha_composite(rgba)
                    img = bg.convert("RGB")
                else:
                    img = raw.convert("RGB")
                info = f"{img.width} × {img.height} px • {detected_dpi:.0f} DPI"

            max_side = 2600
            if max(img.size) > max_side:
                scale = max_side / max(img.size)
                img = img.resize(
                    (max(1, round(img.width * scale)), max(1, round(img.height * scale))),
                    Image.LANCZOS,
                )

            self.source_image = img
            self.analysis_image = img.copy()
            self.file_name_lbl.config(text=os.path.basename(path), wraplength=195, justify="left")
            self.file_info_lbl.config(text=info, wraplength=195, justify="left")
            self.export_btn.config(state="normal")
            self._analyse()
        except Exception as exc:
            self._clear_file()
            messagebox.showerror("Could not open artwork", str(exc), parent=self)

    def _schedule_analysis(self):
        if not self.source_path:
            return
        if self._analysis_job:
            try:
                self.after_cancel(self._analysis_job)
            except Exception:
                pass
        self._analysis_job = self.after(180, self._analyse)

    def _analyse(self):
        self._analysis_job = None
        if self.analysis_image is None:
            return
        try:
            image = self.analysis_image
            array = np.asarray(image.convert("RGB"), dtype=np.uint8)
            tolerance = max(0, min(254, int(self.tolerance_var.get())))
            white_floor = 255 - tolerance
            mask = np.any(array < white_floor, axis=2).astype(np.uint8) * 255

            # Remove isolated noise and close tiny antialiasing gaps.
            kernel = np.ones((3, 3), np.uint8)
            mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel)
            mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)

            try:
                dpi = max(20.0, float(self.dpi_var.get()))
            except Exception:
                dpi = max(20.0, self.source_dpi)
            offset_px = int(round(float(self.offset_var.get()) / 25.4 * dpi * (image.width / max(1, self._source_pixel_width()))))
            if offset_px != 0:
                radius = min(250, abs(offset_px))
                size = radius * 2 + 1
                offset_kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (size, size))
                if offset_px > 0:
                    mask = cv2.dilate(mask, offset_kernel)
                else:
                    mask = cv2.erode(mask, offset_kernel)

            retrieval = cv2.RETR_TREE if self.include_holes_var.get() else cv2.RETR_EXTERNAL
            contours, hierarchy = cv2.findContours(mask, retrieval, cv2.CHAIN_APPROX_NONE)
            min_area = image.width * image.height * max(0.0, float(self.min_area_var.get())) / 100.0
            smooth = 0.0

            result = []
            for index, contour in enumerate(contours):
                area = abs(cv2.contourArea(contour))
                if area < min_area:
                    continue
                if self.include_holes_var.get() and hierarchy is not None:
                    # Keep outer paths and meaningful inner holes. Tiny holes are
                    # already removed by the area threshold above.
                    pass
                perimeter = cv2.arcLength(contour, True)
                epsilon = perimeter * smooth
                approx = cv2.approxPolyDP(contour, epsilon, True) if epsilon > 0 else contour
                points = [(float(p[0][0]), float(p[0][1])) for p in approx]
                if len(points) >= 3:
                    result.append(points)

            self.analysis_contours = result
            self.status_lbl.config(
                text=f"Detected {len(result)} vector path{'s' if len(result) != 1 else ''}.",
                fg=SUCCESS if result else WARNING,
            )
            self._draw_preview()
        except Exception as exc:
            self.analysis_contours = []
            self.status_lbl.config(text=f"Detection failed: {exc}", fg=DANGER)
            self._draw_preview()

    def _source_pixel_width(self):
        if self.source_page_points:
            return self.source_page_points[0] / 72.0 * max(20.0, float(self.dpi_var.get()))
        try:
            with Image.open(self.source_path) as image:
                return image.width
        except Exception:
            return self.analysis_image.width if self.analysis_image else 1

    def _draw_preview(self):
        canvas = self.preview_canvas
        canvas.delete("all")
        if self.analysis_image is None:
            canvas.create_text(
                max(1, canvas.winfo_width()) / 2,
                max(1, canvas.winfo_height()) / 2,
                text="Artwork preview",
                fill="#777777",
                font=("Segoe UI", 12, "bold"),
            )
            return

        cw = max(80, canvas.winfo_width() - 30)
        ch = max(80, canvas.winfo_height() - 30)
        iw, ih = self.analysis_image.size
        scale = min(cw / iw, ch / ih)
        dw, dh = max(1, int(iw * scale)), max(1, int(ih * scale))
        x0 = (canvas.winfo_width() - dw) / 2
        y0 = (canvas.winfo_height() - dh) / 2
        shown = self.analysis_image.resize((dw, dh), Image.LANCZOS)
        self._preview_photo = ImageTk.PhotoImage(shown)
        canvas.create_image(x0, y0, anchor="nw", image=self._preview_photo)

        for contour in self.analysis_contours:
            coords = []
            for x, y in contour:
                coords.extend([x0 + x * scale, y0 + y * scale])
            if len(coords) >= 6:
                coords.extend(coords[:2])
                canvas.create_line(*coords, fill="#FF0000", width=2, smooth=False)

    def _reset_settings(self):
        self.tolerance_var.set(15)
        self.min_area_var.set(0.05)
        self.offset_var.set(0.0)
        self.include_holes_var.set(False)
        self.line_width_var.set(0.25)
        self._schedule_analysis()

    def _source_size_mm(self):
        if self.source_page_points:
            return (
                self.source_page_points[0] * 25.4 / 72.0,
                self.source_page_points[1] * 25.4 / 72.0,
            )
        dpi = max(20.0, float(self.dpi_var.get()))
        try:
            with Image.open(self.source_path) as img:
                width, height = ImageOps.exif_transpose(img).size
        except Exception:
            width, height = self.analysis_image.size
        return width * 25.4 / dpi, height * 25.4 / dpi

    def _export(self):
        if not self.source_path or not self.analysis_contours:
            messagebox.showwarning("Nothing to export", "No vector outline has been detected.", parent=self)
            return

        fmt = self.export_format_var.get().upper()
        extension = ".pdf" if fmt == "PDF" else ".svg"
        base_name = os.path.splitext(os.path.basename(self.source_path))[0] + "_cutline"
        output = filedialog.asksaveasfilename(
            title="Export vector cut line",
            defaultextension=extension,
            initialfile=base_name + extension,
            filetypes=[(fmt, "*" + extension)],
        )
        if not output:
            return

        try:
            if fmt == "PDF":
                self._export_pdf(output)
            else:
                self._export_svg(output)
            messagebox.showinfo("Export complete", f"Saved:\n{output}", parent=self)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc), parent=self)

    def _mapped_contours(self, width, height):
        aw, ah = self.analysis_image.size
        sx, sy = width / aw, height / ah
        return [[(x * sx, y * sy) for x, y in contour] for contour in self.analysis_contours]

    def _export_pdf(self, output):
        width_mm, height_mm = self._source_size_mm()
        width_pt, height_pt = width_mm * 72 / 25.4, height_mm * 72 / 25.4
        doc = fitz.open()
        page = doc.new_page(width=width_pt, height=height_pt)
        artwork_ocg = doc.add_ocg("Artwork") if self.export_mode_var.get() == "artwork" else 0
        cut_ocg = doc.add_ocg("Cut line")
        rect = fitz.Rect(0, 0, width_pt, height_pt)

        source_doc = None
        if self.export_mode_var.get() == "artwork":
            if os.path.splitext(self.source_path)[1].lower() == ".pdf":
                source_doc = fitz.open(self.source_path)
                page.show_pdf_page(rect, source_doc, 0, keep_proportion=False, oc=artwork_ocg)
            else:
                try:
                    page.insert_image(rect, filename=self.source_path, keep_proportion=False, oc=artwork_ocg)
                except Exception:
                    # Unsupported image formats are converted losslessly without
                    # resizing. PNG compression level 0 avoids additional work and
                    # never introduces quality loss.
                    with Image.open(self.source_path) as raw:
                        raw = ImageOps.exif_transpose(raw)
                        buffer = io.BytesIO()
                        raw.save(buffer, format="PNG", compress_level=0)
                    page.insert_image(rect, stream=buffer.getvalue(), keep_proportion=False, oc=artwork_ocg)

        width = max(0.01, float(self.line_width_var.get())) * 72 / 25.4
        for contour in self._mapped_contours(width_pt, height_pt):
            page.draw_polyline(
                [fitz.Point(x, y) for x, y in contour],
                color=(1, 0, 0),
                width=width,
                closePath=True,
                lineJoin=1,
                lineCap=1,
                oc=cut_ocg,
            )

        doc.save(output, garbage=4)
        if source_doc is not None:
            source_doc.close()
        doc.close()

    def _artwork_data_uri(self):
        ext = os.path.splitext(self.source_path)[1].lower()
        direct_types = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".webp": "image/webp",
            ".gif": "image/gif",
        }
        if ext in direct_types:
            data = open(self.source_path, "rb").read()
            return direct_types[ext], base64.b64encode(data).decode("ascii")

        # SVG cannot directly contain a PDF page or every specialist image
        # format. Convert losslessly to an unscaled PNG for the SVG Artwork group.
        if ext == ".pdf":
            doc = fitz.open(self.source_path)
            page = doc[0]
            dpi = max(72.0, float(self.dpi_var.get()))
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=True)
            image = Image.frombytes("RGBA", (pix.width, pix.height), pix.samples)
            doc.close()
        else:
            image = ImageOps.exif_transpose(Image.open(self.source_path))
        buffer = io.BytesIO()
        image.save(buffer, format="PNG", compress_level=0)
        return "image/png", base64.b64encode(buffer.getvalue()).decode("ascii")

    def _export_svg(self, output):
        aw, ah = self.analysis_image.size
        width_mm, height_mm = self._source_size_mm()
        line_px = max(0.01, float(self.line_width_var.get())) / max(width_mm / aw, height_mm / ah)

        parts = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
            f'width="{width_mm:.6f}mm" height="{height_mm:.6f}mm" viewBox="0 0 {aw} {ah}">',
        ]

        if self.export_mode_var.get() == "artwork":
            mime, encoded = self._artwork_data_uri()
            parts.extend([
                '<g id="Artwork" inkscape:label="Artwork" xmlns:inkscape="http://www.inkscape.org/namespaces/inkscape">',
                f'<image x="0" y="0" width="{aw}" height="{ah}" preserveAspectRatio="none" '
                f'xlink:href="data:{mime};base64,{encoded}"/>',
                '</g>',
            ])

        parts.append('<g id="Cut_line" inkscape:label="Cut line" xmlns:inkscape="http://www.inkscape.org/namespaces/inkscape">')
        for contour in self.analysis_contours:
            commands = [f'M {contour[0][0]:.3f} {contour[0][1]:.3f}']
            commands.extend(f'L {x:.3f} {y:.3f}' for x, y in contour[1:])
            commands.append('Z')
            parts.append(
                f'<path d="{" ".join(commands)}" fill="none" stroke="#FF0000" '
                f'stroke-width="{line_px:.4f}" stroke-linejoin="round" stroke-linecap="round"/>'
            )
        parts.extend(['</g>', '</svg>'])
        Path(output).write_text("\n".join(parts), encoding="utf-8")


# ── Print Layout Tab ──────────────────────────────────────────────────────────
DEFAULT_PRINT_LAYOUTS = {
    "Default": {"paper":"SRA4","orientation":"Portrait","item_w":210.0,"item_h":297.0,"cols":1,"rows":1,"mode":"In order","duplex":False,"bleed":False,"bleed_mm":2.0,"spacing":0.0,"grayscale":False,"brightness":100,"dpi":300,"cut_marks":False,"cut_labels":False},
    "Business Card": {"paper":"A4","orientation":"Portrait","item_w":90.0,"item_h":50.0,"cols":2,"rows":5,"mode":"Repeat","duplex":True,"bleed":True,"bleed_mm":3.0,"spacing":6.0,"grayscale":False,"brightness":100,"dpi":300,"cut_marks":False,"cut_labels":False},
    "Document Photo": {"paper":"10×15 cm (102×152)","orientation":"Landscape","item_w":35.0,"item_h":45.0,"cols":4,"rows":2,"mode":"Repeat","duplex":False,"bleed":False,"bleed_mm":0.0,"spacing":2.0,"grayscale":False,"brightness":100,"dpi":300,"cut_marks":False,"cut_labels":False},
}

def load_print_layouts():
    import json
    layouts = {k: dict(v) for k,v in DEFAULT_PRINT_LAYOUTS.items()}
    try:
        if os.path.exists(PRINT_LAYOUTS_FILE):
            with open(PRINT_LAYOUTS_FILE,"r",encoding="utf-8") as f:
                saved=json.load(f)
            if isinstance(saved,dict): layouts.update(saved)
    except Exception: pass
    return layouts

def save_print_layouts(layouts):
    import json
    try:
        with open(PRINT_LAYOUTS_FILE,"w",encoding="utf-8") as f: json.dump(layouts,f,indent=2)
    except Exception: pass

    def destroy(self):
        try:
            if getattr(self, "pil_img", None) is not None:
                self.pil_img.close()
                self.pil_img = None
            self.tk_img = None
        except Exception:
            pass
        super().destroy()

class PrintCropCanvas(CropCanvas):
    def __init__(self,*args,**kwargs):
        self.preview_grayscale=False; self.preview_brightness=100
        super().__init__(*args,**kwargs)
    def _fit_to_canvas(self):
        iw,ih=self.pil_img.size; scale=min(self.canvas_w/iw,self.canvas_h/ih)
        dw,dh=max(1,int(iw*scale)),max(1,int(ih*scale))
        self.offset_x=(self.canvas_w-dw)//2; self.offset_y=(self.canvas_h-dh)//2
        self.disp_w,self.disp_h,self.scale=dw,dh,scale
        shown=self.pil_img; temporary=[]
        if self.preview_grayscale:
            shown=ImageOps.grayscale(shown).convert("RGB"); temporary.append(shown)
        if self.preview_brightness!=100:
            shown=ImageEnhance.Brightness(shown).enhance(self.preview_brightness/100); temporary.append(shown)
        resized=shown.resize((dw,dh),Image.Resampling.BILINEAR)
        try:self.tk_img=ImageTk.PhotoImage(resized)
        finally:
            resized.close()
            for image in temporary:
                if image is not self.pil_img:
                    try:image.close()
                    except Exception:pass
    def set_effects(self,gray,brightness):
        self.preview_grayscale=bool(gray); self.preview_brightness=int(brightness)
        self._fit_to_canvas(); self._draw()

class PrintLayoutTab(tk.Frame,DropMixin):
    def __init__(self,parent):
        super().__init__(parent,bg=BG)

        # Initialise preview state before building widgets or loading a preset.
        self.files=[]
        self.canvases=[]
        self.layouts=load_print_layouts()
        self.custom_sizes=load_custom_sizes()
        self._preview_job=None
        self._preview_item_cache=ImageMemoryCache(128 * 1024 * 1024)
        self.preview_images=[]
        self._proportion_ratio=None
        self._proportion_syncing=False
        self.pdf_sources={}
        self.display_names={}
        self._temp_dir=tempfile.mkdtemp(prefix="copypro_print_")

        self._build()
        self._load_layout("Default")
    def on_show(self):
        self.custom_sizes=load_custom_sizes(); self.paper_cb.config(values=self._paper_names())
    def _paper_names(self): return list(PAPER_SIZES_MM)+list(self.custom_sizes)
    def _build(self):
        # Three-column workspace:
        # fixed settings on the left, large preview in the middle,
        # and a compact source-file panel on the right.
        self.grid_rowconfigure(0,weight=1)
        self.grid_columnconfigure(0,weight=0)
        self.grid_columnconfigure(1,weight=1)
        self.grid_columnconfigure(2,weight=0)

        # LEFT: fixed compact settings. No settings scrollbar or mouse-wheel capture.
        left=tk.Frame(self,bg=SURFACE,width=286)
        left.grid(row=0,column=0,sticky="nsew",padx=(0,1))
        left.grid_propagate(False)

        lbl(left,"PRINT LAYOUT",8,MUTED,True).pack(pady=(8,3),padx=12,anchor="w")
        pr=tk.Frame(left,bg=SURFACE); pr.pack(fill="x",padx=12)
        self.layout_var=tk.StringVar(value="Default")
        self.layout_cb=ttk.Combobox(pr,textvariable=self.layout_var,values=list(self.layouts),state="readonly",width=18)
        self.layout_cb.pack(side="left",fill="x",expand=True)
        self.layout_cb.bind("<<ComboboxSelected>>",lambda e:self._load_layout(self.layout_var.get()))
        tk.Button(pr,text="Save",command=self._save_layout_dialog,bg=SURFACE2,fg=TEXT,relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="left",padx=(5,0))

        lbl(left,"Paper size",8).pack(padx=12,anchor="w",pady=(5,0))
        self.paper_var=tk.StringVar(value="A4")
        self.paper_cb=ttk.Combobox(left,textvariable=self.paper_var,values=self._paper_names(),state="readonly")
        self.paper_cb.pack(fill="x",padx=12,pady=(1,2)); self.paper_cb.bind("<<ComboboxSelected>>",self._changed)
        rr=tk.Frame(left,bg=SURFACE); rr.pack(fill="x",padx=12)
        self.orientation_var=tk.StringVar(value="Portrait")
        for x in ("Portrait","Landscape"):
            tk.Radiobutton(rr,text=x,variable=self.orientation_var,value=x,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._changed).pack(side="left")
        self.dpi_var=tk.IntVar(value=300)
        ttk.Combobox(rr,textvariable=self.dpi_var,values=[150,300,600],state="readonly",width=5).pack(side="right")
        lbl(rr,"DPI",7,MUTED).pack(side="right",padx=3)

        sep(left)
        grid_hdr=tk.Frame(left,bg=SURFACE); grid_hdr.pack(fill="x",padx=12)
        lbl(grid_hdr,"ITEM & GRID",8,MUTED,True).pack(side="left")
        tk.Button(grid_hdr,text="Auto…",command=self._open_layout_calculator,bg=SURFACE2,fg=TEXT,relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2",padx=6,pady=1).pack(side="right")
        size_row=tk.Frame(left,bg=SURFACE); size_row.pack(fill="x",padx=12,pady=(2,0))
        self.item_w=tk.DoubleVar(value=90); self.item_h=tk.DoubleVar(value=50)
        self.item_w_input=tk.StringVar(value="90"); self.item_h_input=tk.StringVar(value="50")
        self.cols_var=tk.IntVar(value=2); self.rows_var=tk.IntVar(value=5)
        dimensions=tk.Frame(size_row,bg=SURFACE); dimensions.pack(fill="x")

        lbl(dimensions,"W:",8,MUTED,True).pack(side="left",padx=(0,4),pady=(1,0))
        width_entry=entry(dimensions,self.item_w_input,8); width_entry.pack(side="left",fill="x",expand=True)
        width_entry.bind("<KeyRelease>",lambda event:self._proportional_input_changed("width"))
        width_entry.bind("<Return>",self._apply_item_size); width_entry.bind("<KP_Enter>",self._apply_item_size)

        self.proportional_var=tk.BooleanVar(value=False)
        self.chain_btn=tk.Button(
            dimensions,
            text="\U0001F517",
            command=self._toggle_proportional,
            width=3,
            padx=0,
            pady=2,
            bg=SURFACE2,
            fg=MUTED,
            activebackground=SURFACE2,
            activeforeground=TEXT,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER,
            font=("Segoe UI Emoji",10),
            cursor="hand2",
            takefocus=True,
        )
        self.chain_btn.pack(side="left",padx=5)

        lbl(dimensions,"H:",8,MUTED,True).pack(side="left",padx=(0,4),pady=(1,0))
        height_entry=entry(dimensions,self.item_h_input,8); height_entry.pack(side="left",fill="x",expand=True)
        height_entry.bind("<KeyRelease>",lambda event:self._proportional_input_changed("height"))
        height_entry.bind("<Return>",self._apply_item_size); height_entry.bind("<KP_Enter>",self._apply_item_size)

        size_actions=tk.Frame(left,bg=SURFACE); size_actions.pack(fill="x",padx=12,pady=(4,0))
        tk.Button(size_actions,text="Update",command=self._apply_item_size,bg=ACCENT,fg="white",
            relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2",padx=8,pady=4).pack(side="left",fill="x",expand=True)
        tk.Button(size_actions,text="Presets",command=self._open_item_size_presets,bg=SURFACE2,fg=TEXT,
            relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2",padx=8,pady=4).pack(side="left",fill="x",expand=True,padx=(5,0))
        self._update_chain_button()

        def grid_spin(parent,label,var):
            f=tk.Frame(parent,bg=SURFACE)
            lbl(f,label,7,MUTED).pack(anchor="w")
            sp=tk.Spinbox(f,from_=1,to=99,textvariable=var,width=7,bg=SURFACE2,fg=TEXT,
                          insertbackground=TEXT,buttonbackground=SURFACE2,relief="flat",
                          font=("Segoe UI",9),command=self._changed)
            sp.pack(fill="x")
            sp.bind("<KeyRelease>",self._changed)
            def wheel(ev):
                step=1 if ev.delta>0 else -1
                try: var.set(max(1,min(99,int(var.get())+step)))
                except Exception: var.set(1)
                self._changed(); return "break"
            sp.bind("<MouseWheel>",wheel)
            return f

        grid=tk.Frame(left,bg=SURFACE); grid.pack(fill="x",padx=12,pady=(3,0))
        for i,(t,v) in enumerate((("Columns",self.cols_var),("Rows",self.rows_var))):
            f=grid_spin(grid,t,v); f.grid(row=0,column=i,padx=(0 if i==0 else 4,0),sticky="ew"); grid.grid_columnconfigure(i,weight=1)

        mr=tk.Frame(left,bg=SURFACE); mr.pack(fill="x",padx=12,pady=(4,0)); self.mode_var=tk.StringVar(value="Repeat")
        for x in ("Repeat","Cut & stack","In order"):
            tk.Radiobutton(mr,text=x,variable=self.mode_var,value=x,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._mode_changed).pack(anchor="w")
        self.duplex_var=tk.BooleanVar(value=True)
        tk.Checkbutton(left,text="Double-sided (left bind)",variable=self.duplex_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._changed).pack(padx=12,anchor="w")

        sep(left); lbl(left,"SPACING & BLEED",8,MUTED,True).pack(padx=12,anchor="w")
        sr=tk.Frame(left,bg=SURFACE); sr.pack(fill="x",padx=12,pady=(2,0)); lbl(sr,"Trim spacing (mm)",8).pack(side="left")
        self.spacing_var=tk.DoubleVar(value=6); e=entry(sr,self.spacing_var,6); e.pack(side="right"); e.bind("<KeyRelease>",self._changed)
        br=tk.Frame(left,bg=SURFACE); br.pack(fill="x",padx=12,pady=(2,0)); self.bleed_var=tk.BooleanVar(value=True)
        tk.Checkbutton(br,text="Add bleed",variable=self.bleed_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._toggle_bleed).pack(side="left")
        self.bleed_mm=tk.DoubleVar(value=3); self.bleed_entry=entry(br,self.bleed_mm,5); self.bleed_entry.pack(side="right"); self.bleed_entry.bind("<KeyRelease>",self._changed); lbl(br,"mm",7,MUTED).pack(side="right",padx=2)
        cr=tk.Frame(left,bg=SURFACE); cr.pack(fill="x",padx=12,pady=(2,0))
        self.cut_marks_var=tk.BooleanVar(value=False); self.cut_labels_var=tk.BooleanVar(value=False)
        tk.Checkbutton(cr,text="Cut marks",variable=self.cut_marks_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._toggle_cut_marks).pack(side="left")
        self.cut_labels_chk=tk.Checkbutton(cr,text="Labels",variable=self.cut_labels_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._changed)
        self.cut_labels_chk.pack(side="right")

        sep(left); lbl(left,"IMAGE ADJUSTMENTS",8,MUTED,True).pack(padx=12,anchor="w")
        self.gray_var=tk.BooleanVar(value=False)
        tk.Checkbutton(left,text="Grayscale",variable=self.gray_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._effects_changed).pack(padx=12,anchor="w")
        self.brightness_var=tk.IntVar(value=100); self.brightness_lbl=lbl(left,"Brightness: 100%",8); self.brightness_lbl.pack(padx=12,anchor="w")
        ttk.Scale(left,from_=25,to=175,variable=self.brightness_var,orient="horizontal",command=self._brightness_changed).pack(fill="x",padx=12)
        self.fit_lbl=lbl(left,"",8,MUTED); self.fit_lbl.pack(padx=12,anchor="w",pady=(2,0))
        self.export_btn=styled_btn(left,"Export PDF",self._export_pdf,style="success"); self.export_btn.pack(fill="x",padx=12,pady=(5,2))
        styled_btn(left,"Print current layout  (Ctrl+P)",self._print_current_layout,style="secondary").pack(fill="x",padx=12,pady=2)
        self.status=lbl(left,"",8,MUTED); self.status.pack(padx=12,anchor="w",pady=(0,4))

        # MIDDLE: the imposed-sheet preview receives most of the window.
        preview_panel=tk.Frame(self,bg=BG)
        preview_panel.grid(row=0,column=1,sticky="nsew",padx=(8,4),pady=8)
        preview_panel.grid_rowconfigure(1,weight=1)
        preview_panel.grid_columnconfigure(0,weight=1)

        ph=tk.Frame(preview_panel,bg=BG)
        ph.grid(row=0,column=0,sticky="ew",pady=(0,4))
        lbl(ph,"FINAL IMPOSED SHEETS",9,MUTED,True).pack(side="left")
        self.preview_count_lbl=lbl(ph,"0 sheets",8,MUTED)
        self.preview_count_lbl.pack(side="right")

        preview_body=tk.Frame(
            preview_panel,
            bg=SURFACE,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        preview_body.grid(row=1,column=0,sticky="nsew")
        preview_body.grid_rowconfigure(0,weight=1)
        preview_body.grid_columnconfigure(0,weight=1)

        self.preview_canvas=tk.Canvas(
            preview_body,
            bg=BG,
            highlightthickness=0,
        )
        preview_v=ttk.Scrollbar(
            preview_body,
            orient="vertical",
            command=self.preview_canvas.yview,
        )
        preview_h=ttk.Scrollbar(
            preview_body,
            orient="horizontal",
            command=self.preview_canvas.xview,
        )
        self.preview_canvas.configure(
            yscrollcommand=preview_v.set,
            xscrollcommand=preview_h.set,
        )
        self.preview_canvas.grid(row=0,column=0,sticky="nsew")
        preview_v.grid(row=0,column=1,sticky="ns")
        preview_h.grid(row=1,column=0,sticky="ew")

        self.preview_frame=tk.Frame(self.preview_canvas,bg=BG)
        self.preview_window=self.preview_canvas.create_window(
            (0,0),
            window=self.preview_frame,
            anchor="nw",
        )
        self.preview_frame.bind(
            "<Configure>",
            lambda e:self.preview_canvas.configure(
                scrollregion=self.preview_canvas.bbox("all")
            ),
        )
        self.preview_canvas.bind(
            "<Configure>",
            lambda e:(
                self.preview_canvas.itemconfig(
                    self.preview_window,
                    width=max(1,e.width),
                ),
                self._schedule_preview(),
            ),
        )

        def preview_wheel(event):
            self.preview_canvas.yview_scroll(int(-event.delta/120),"units")
            return "break"

        self.preview_canvas.bind(
            "<Enter>",
            lambda e:self.preview_canvas.bind_all("<MouseWheel>",preview_wheel),
        )
        self.preview_canvas.bind(
            "<Leave>",
            lambda e:self.preview_canvas.unbind_all("<MouseWheel>"),
        )
        self.setup_drop(
            self.preview_canvas,
            self._drop_files,
            IMAGE_EXTS | {".pdf"},
        )
        self.preview_images=[]

        # RIGHT: compact vertical source-file panel. The file list scrolls,
        # while the drop area remains fixed at the bottom.
        imports=tk.Frame(
            self,
            bg=SURFACE,
            width=270,
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        imports.grid(row=0,column=2,sticky="nsew",padx=(4,8),pady=8)
        imports.grid_propagate(False)
        imports.grid_rowconfigure(1,weight=1)
        imports.grid_columnconfigure(0,weight=1)

        ih=tk.Frame(imports,bg=SURFACE)
        ih.grid(row=0,column=0,sticky="ew",padx=8,pady=(7,4))
        lbl(ih,"SOURCE FILES",8,MUTED,True).pack(side="left")
        styled_btn(ih,"Add",self._add_files).pack(side="right",padx=(4,0))
        styled_btn(ih,"Clear",self._clear,style="secondary").pack(side="right")

        cards_wrap=tk.Frame(imports,bg=SURFACE)
        cards_wrap.grid(row=1,column=0,sticky="nsew",padx=6,pady=(0,5))
        cards_wrap.grid_rowconfigure(0,weight=1)
        cards_wrap.grid_columnconfigure(0,weight=1)

        self.cards_canvas=tk.Canvas(
            cards_wrap,
            bg=SURFACE,
            highlightthickness=0,
        )
        cards_v=ttk.Scrollbar(
            cards_wrap,
            orient="vertical",
            command=self.cards_canvas.yview,
        )
        self.cards_canvas.configure(yscrollcommand=cards_v.set)
        self.cards_canvas.grid(row=0,column=0,sticky="nsew")
        cards_v.grid(row=0,column=1,sticky="ns")

        self.cards_frame=tk.Frame(self.cards_canvas,bg=SURFACE)
        self.cards_win=self.cards_canvas.create_window(
            (0,0),
            window=self.cards_frame,
            anchor="nw",
        )
        self.cards_frame.bind(
            "<Configure>",
            lambda e:self.cards_canvas.configure(
                scrollregion=self.cards_canvas.bbox("all")
            ),
        )
        self.cards_canvas.bind(
            "<Configure>",
            lambda e:self.cards_canvas.itemconfig(
                self.cards_win,
                width=max(1,e.width),
            ),
        )

        def cards_wheel(event):
            self.cards_canvas.yview_scroll(int(-event.delta/120),"units")
            return "break"

        self.cards_canvas.bind(
            "<Enter>",
            lambda e:self.cards_canvas.bind_all("<MouseWheel>",cards_wheel),
        )
        self.cards_canvas.bind(
            "<Leave>",
            lambda e:self.cards_canvas.unbind_all("<MouseWheel>"),
        )
        self.setup_drop(
            self.cards_canvas,
            self._drop_files,
            IMAGE_EXTS | {".pdf"},
        )

        self.drop_zone=tk.Label(
            imports,
            text="⬇  DROP FILES HERE",
            bg=SURFACE2,
            fg=MUTED,
            font=("Segoe UI",8,"bold"),
            cursor="hand2",
            pady=11,
        )
        self.drop_zone.grid(
            row=2,
            column=0,
            sticky="ew",
            padx=7,
            pady=(0,7),
        )
        self.drop_zone.bind("<Button-1>",lambda e:self._add_files())
        self.setup_drop(
            self.drop_zone,
            self._drop_files,
            IMAGE_EXTS | {".pdf"},
        )

        self.bind_all("<Control-p>",self._ctrl_p)
        self.bind_all("<Control-P>",self._ctrl_p)
        self._show_empty()
    def _num(self,var,default):
        try:return float(var.get())
        except:return default
    def _paper_mm(self):
        w,h=self.custom_sizes.get(self.paper_var.get(),PAPER_SIZES_MM.get(self.paper_var.get(),(210,297)))
        return (max(w,h),min(w,h)) if self.orientation_var.get()=="Landscape" else (min(w,h),max(w,h))
    def _metrics(self):
        pw,ph=self._paper_mm(); iw=max(.1,self._num(self.item_w,90)); ih=max(.1,self._num(self.item_h,50)); cols=max(1,int(self._num(self.cols_var,1))); rows=max(1,int(self._num(self.rows_var,1))); gap=max(0,self._num(self.spacing_var,0)); bleed=max(0,self._num(self.bleed_mm,0)) if self.bleed_var.get() else 0
        gw=cols*iw+(cols-1)*gap+2*bleed; gh=rows*ih+(rows-1)*gap+2*bleed; x0=(pw-gw)/2+bleed; y0=(ph-gh)/2+bleed
        return pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,gw,gh
    def _first_source_ratio(self):
        if not self.files:
            return None
        first=self.files[0]
        canvas=next((c for path,c in self.canvases if path==first),None)
        if canvas is not None:
            ratio=float(getattr(canvas,"source_aspect_ratio",0) or 0)
            if ratio>0:return ratio
        try:
            if first in self.pdf_sources:
                pdf_path,page_no=self.pdf_sources[first]
                doc=fitz.open(pdf_path)
                try:
                    rect=doc[page_no].rect
                    return rect.width/max(1e-6,rect.height)
                finally:doc.close()
            with Image.open(first) as image:
                corrected=ImageOps.exif_transpose(image)
                return corrected.width/max(1,corrected.height)
        except Exception:return None

    def _update_chain_button(self):
        if not hasattr(self,"chain_btn"):return
        linked=bool(self.proportional_var.get())
        self.chain_btn.config(
            text="\U0001F517",
            bg=ACCENT if linked else SURFACE2,
            fg="white" if linked else MUTED,
            activebackground=ACCENT if linked else SURFACE2,
            activeforeground="white" if linked else TEXT,
            highlightbackground=ACCENT if linked else BORDER,
            highlightcolor=ACCENT if linked else BORDER,
            relief="sunken" if linked else "flat",
        )

    def _toggle_proportional(self):
        self.proportional_var.set(not self.proportional_var.get())
        if self.proportional_var.get():
            ratio=self._first_source_ratio()
            if not ratio:
                self.proportional_var.set(False)
                self._update_chain_button()
                messagebox.showinfo("Proportional sizing","Add at least one source image or PDF page first.",parent=self)
                return
            self._proportion_ratio=ratio
            self._proportional_input_changed("width")
            self.status.config(text="Linked sizing uses the first source image ratio.",fg=MUTED)
        else:self._proportion_ratio=None
        self._update_chain_button()

    def _proportional_input_changed(self,axis):
        if not self.proportional_var.get() or self._proportion_syncing:return
        ratio=self._proportion_ratio or self._first_source_ratio()
        if not ratio or ratio<=0:return
        try:
            self._proportion_syncing=True
            if axis=="width":
                value=float(self.item_w_input.get().replace(",","."))
                if value>0:self.item_h_input.set(f"{value/ratio:.2f}".rstrip("0").rstrip("."))
            else:
                value=float(self.item_h_input.get().replace(",","."))
                if value>0:self.item_w_input.set(f"{value*ratio:.2f}".rstrip("0").rstrip("."))
        except Exception:pass
        finally:self._proportion_syncing=False

    def _set_item_size(self,width,height,refresh=True):
        width=max(.1,float(width)); height=max(.1,float(height))
        self.item_w.set(width); self.item_h.set(height)
        self.item_w_input.set(f"{width:g}"); self.item_h_input.set(f"{height:g}")
        if not hasattr(self,"_preview_item_cache"):
            self._preview_item_cache=ImageMemoryCache(128 * 1024 * 1024)
        self._preview_item_cache.clear()
        if refresh:self._rebuild()

    def _apply_item_size(self,event=None):
        try:self._set_item_size(float(self.item_w_input.get().replace(",",".")),
                                float(self.item_h_input.get().replace(",",".")),True)
        except Exception:messagebox.showwarning("Invalid item size",
            "Enter valid positive width and height values.",parent=self)
        return "break" if event is not None else None

    def _changed(self,event=None): self._schedule_preview()
    def _mode_changed(self):
        mode=self.mode_var.get()
        if mode=="Cut & stack": self.duplex_var.set(True)
        elif mode=="In order": self.duplex_var.set(False)
        self._changed()
    def _toggle_bleed(self): self.bleed_entry.config(state="normal" if self.bleed_var.get() else "disabled"); self._changed()
    def _toggle_cut_marks(self):
        if not self.cut_marks_var.get(): self.cut_labels_var.set(False)
        self.cut_labels_chk.config(state="normal" if self.cut_marks_var.get() else "disabled")
        self._changed()
    def _open_item_size_presets(self):
        popup=tk.Toplevel(self); popup.title("Item size presets"); popup.configure(bg=BG); popup.resizable(False,False); popup.grab_set()
        popup.geometry(f"430x480+{self.winfo_rootx()+300}+{self.winfo_rooty()+90}")
        tk.Label(popup,text="Choose item size",bg=BG,fg=TEXT,font=("Segoe UI",13,"bold")).pack(pady=(16,4))
        tk.Label(popup,text="Width and height remain directly editable in the main tab.",bg=BG,fg=MUTED,font=("Segoe UI",9)).pack()
        holder=tk.Frame(popup,bg=BG); holder.pack(fill="both",expand=True,padx=16,pady=12)
        canvas=tk.Canvas(holder,bg=SURFACE,highlightthickness=0); canvas.pack(side="left",fill="both",expand=True)
        sb=ttk.Scrollbar(holder,orient="vertical",command=canvas.yview); sb.pack(side="right",fill="y"); canvas.configure(yscrollcommand=sb.set)
        inner=tk.Frame(canvas,bg=SURFACE); win=canvas.create_window((0,0),window=inner,anchor="nw")
        inner.bind("<Configure>",lambda e:canvas.configure(scrollregion=canvas.bbox("all"))); canvas.bind("<Configure>",lambda e:canvas.itemconfig(win,width=e.width))
        def apply_size(name,w,h): self._set_item_size(w,h,True); popup.destroy()
        def add_row(parent,name,w,h,delete=False):
            row=tk.Frame(parent,bg=SURFACE2); row.pack(fill="x",padx=8,pady=2)
            tk.Button(row,text=f"{name}   {w:g} × {h:g} mm",command=lambda:apply_size(name,w,h),anchor="w",bg=SURFACE2,fg=TEXT,relief="flat",font=("Segoe UI",9),cursor="hand2").pack(side="left",fill="x",expand=True)
            if delete: tk.Button(row,text="Delete",command=lambda n=name:delete_custom(n),bg=SURFACE2,fg=DANGER,relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="right")
        lbl(inner,"DEFAULT SIZES",8,MUTED,True).pack(anchor="w",padx=8,pady=(8,3))
        for name,(w,h) in PAPER_SIZES_MM.items(): add_row(inner,name,w,h,False)
        custom_header=lbl(inner,"CUSTOM SIZES",8,MUTED,True); custom_header.pack(anchor="w",padx=8,pady=(12,3))
        custom_box=tk.Frame(inner,bg=SURFACE); custom_box.pack(fill="x")
        def render_custom():
            for child in custom_box.winfo_children(): child.destroy()
            if not self.custom_sizes: lbl(custom_box,"No custom sizes",8,MUTED).pack(anchor="w",padx=8,pady=6)
            for name,(w,h) in self.custom_sizes.items(): add_row(custom_box,name,w,h,True)
        def delete_custom(name):
            if name in self.custom_sizes: del self.custom_sizes[name]; save_custom_sizes(self.custom_sizes); self.paper_cb.config(values=self._paper_names()); render_custom()
        def save_current():
            name=simpledialog.askstring("Save item size","Preset name:",parent=popup)
            if not name:return
            w=max(.1,self._num(self.item_w,90)); h=max(.1,self._num(self.item_h,50)); self.custom_sizes[name.strip()]=(w,h); save_custom_sizes(self.custom_sizes); self.paper_cb.config(values=self._paper_names()); render_custom()
        render_custom()
        buttons=tk.Frame(popup,bg=BG); buttons.pack(pady=(0,14)); styled_btn(buttons,"Save current dimensions",save_current).pack(side="left",padx=4); styled_btn(buttons,"Close",popup.destroy,style="secondary").pack(side="left",padx=4)

    def _open_layout_calculator(self):
        popup=tk.Toplevel(self); popup.title("Automatic layout calculator"); popup.configure(bg=BG); popup.resizable(False,False); popup.grab_set()
        popup.geometry(f"470x500+{self.winfo_rootx()+330}+{self.winfo_rooty()+80}")
        tk.Label(popup,text="Automatic layout calculator",bg=BG,fg=TEXT,font=("Segoe UI",13,"bold")).pack(pady=(18,4))
        tk.Label(popup,text="Calculates the highest-capacity centred layout for the current paper.",bg=BG,fg=MUTED,font=("Segoe UI",9)).pack()
        box=tk.Frame(popup,bg=SURFACE); box.pack(fill="x",padx=18,pady=14)
        w_var=tk.DoubleVar(value=self._num(self.item_w,90)); h_var=tk.DoubleVar(value=self._num(self.item_h,50))
        bleed_var=tk.DoubleVar(value=2.0); gap_var=tk.DoubleVar(value=self._num(self.spacing_var,0)); qty_var=tk.IntVar(value=max(1,len(self.files) or 1))
        rotate_var=tk.BooleanVar(value=True)
        for label,var in (("Item width (mm)",w_var),("Item height (mm)",h_var),("Bleed (mm)",bleed_var),("Trim spacing (mm)",gap_var),("Required quantity",qty_var)):
            row=tk.Frame(box,bg=SURFACE); row.pack(fill="x",padx=12,pady=4); lbl(row,label,9).pack(side="left"); entry(row,var,9).pack(side="right")
        tk.Checkbutton(box,text="Allow rotating the item 90°",variable=rotate_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",9)).pack(anchor="w",padx=12,pady=(6,4))
        result=tk.Label(box,text="",bg=SURFACE,fg=TEXT,font=("Segoe UI",10),justify="left",anchor="w"); result.pack(fill="x",padx=12,pady=(8,12))
        state={}
        def calculate():
            pw,ph=self._paper_mm()
            try: iw=max(.1,float(w_var.get())); ih=max(.1,float(h_var.get())); bleed=max(0,float(bleed_var.get())); gap=max(0,float(gap_var.get())); qty=max(1,int(qty_var.get()))
            except Exception:
                result.config(text="Please enter valid positive dimensions and quantity.",fg=DANGER); return
            candidates=[]
            for rotated in ([False,True] if rotate_var.get() and abs(iw-ih)>.001 else [False]):
                w,h=(ih,iw) if rotated else (iw,ih)
                cols=max(0,int((pw-2*bleed+gap)//(w+gap))); rows=max(0,int((ph-2*bleed+gap)//(h+gap)))
                copies=cols*rows; used_w=cols*w+max(0,cols-1)*gap+2*bleed; used_h=rows*h+max(0,rows-1)*gap+2*bleed
                waste=100*(1-(used_w*used_h)/(pw*ph)) if pw*ph else 100
                candidates.append((copies,-waste,cols,rows,rotated,w,h,used_w,used_h))
            best=max(candidates) if candidates else (0,0,0,0,False,iw,ih,0,0)
            copies,negw,cols,rows,rotated,w,h,uw,uh=best; sheets=math.ceil(qty/copies) if copies else 0
            state.update(cols=cols,rows=rows,rotated=rotated,w=w,h=h,bleed=bleed,gap=gap,qty=qty)
            result.config(text=(f"Best grid: {cols} columns × {rows} rows\nCopies per sheet: {copies}\nSheets for {qty}: {sheets}\nCentered area: {uw:.1f} × {uh:.1f} mm\nItem rotation: {'90°' if rotated else 'No'}"),fg=SUCCESS if copies else DANGER)
        def apply_result():
            calculate()
            if state.get("cols",0)<1:return
            self._set_item_size(state["w"],state["h"],False); self.cols_var.set(state["cols"]); self.rows_var.set(state["rows"])
            self.bleed_mm.set(state["bleed"]); self.bleed_var.set(state["bleed"]>0); self.spacing_var.set(state["gap"]); self._toggle_bleed(); self._rebuild(); popup.destroy()
        actions=tk.Frame(popup,bg=BG); actions.pack(pady=4)
        styled_btn(actions,"Calculate",calculate).pack(side="left",padx=4); styled_btn(actions,"Apply layout",apply_result,style="success").pack(side="left",padx=4); styled_btn(actions,"Cancel",popup.destroy,style="secondary").pack(side="left",padx=4)
        calculate()
    def _brightness_changed(self,v): self.brightness_lbl.config(text=f"Brightness: {int(float(v))}%"); self._effects_changed()
    def _effects_changed(self):
        self._preview_item_cache.clear()
        for _,c in self.canvases:c.set_effects(self.gray_var.get(),self.brightness_var.get())
        self._schedule_preview()
    def _schedule_preview(self):
        if self._preview_job:
            try:self.after_cancel(self._preview_job)
            except:pass
        self._preview_job=self.after(180,self._draw_preview)
    def _draw_preview(self):
        self._preview_job=None
        for photo in self.preview_images:
            try:photo.__del__()
            except Exception:pass
        self.preview_images.clear()
        for child in self.preview_frame.winfo_children(): child.destroy()
        try:
            pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,gw,gh=self._metrics()
        except Exception:
            return
        fits=gw<=pw+1e-6 and gh<=ph+1e-6
        self.fit_lbl.config(text=(f"Centered: {gw:.1f} × {gh:.1f} mm" if fits else f"Does not fit: {gw:.1f} × {gh:.1f} mm"),fg=MUTED if fits else DANGER)
        if not self.files or not self.canvases:
            self.preview_count_lbl.config(text="0 sheets")
            tk.Label(self.preview_frame,text="Add source files to see the final imposed sheets",bg=BG,fg=MUTED,font=("Segoe UI",12)).grid(row=0,column=0,pady=70,padx=30)
            return
        if not fits:
            self.preview_count_lbl.config(text="Layout does not fit")
            tk.Label(self.preview_frame,text="The current grid is larger than the selected paper.",bg=BG,fg=DANGER,font=("Segoe UI",11,"bold")).grid(row=0,column=0,pady=70,padx=30)
            return
        specs=self._page_specs()
        self.preview_count_lbl.config(text=f"{len(specs)} sheet{'s' if len(specs)!=1 else ''}")
        available=max(320,self.preview_canvas.winfo_width()-30)
        # Never exceed four columns. Smaller windows naturally reduce the count.
        ncols=max(1,min(4,available//245))
        card_w=max(190,min(330,(available-(ncols-1)*12)//ncols-14))
        for idx,(assignments,back,label) in enumerate(specs):
            r,col=divmod(idx,ncols)
            card=tk.Frame(self.preview_frame,bg=SURFACE,padx=7,pady=7,highlightbackground=BORDER,highlightthickness=1)
            card.grid(row=r,column=col,padx=6,pady=6,sticky="n")
            try:
                # Rendering at 55 DPI is enough for a clear on-screen preview and
                # stays responsive even when many imposed pages are shown.
                sheet=self._sheet(assignments,back=back,dpi_override=42,preview_guides=True)
                max_h=420
                scale=min(card_w/sheet.width,max_h/sheet.height,1.0)
                shown=sheet.resize((max(1,int(sheet.width*scale)),max(1,int(sheet.height*scale))),Image.Resampling.BILINEAR)
                photo=ImageTk.PhotoImage(shown)
                shown.close(); sheet.close()
                self.preview_images.append(photo)
                tk.Label(card,image=photo,bg=SURFACE2,bd=0,highlightbackground=BORDER,highlightthickness=1).pack()
                tk.Label(card,text=label,bg=SURFACE,fg=TEXT,font=("Segoe UI",8,"bold"),wraplength=card_w).pack(fill="x",pady=(5,0))
            except Exception as ex:
                tk.Label(card,text=f"Preview failed\n{ex}",bg=BG,fg=DANGER,font=("Segoe UI",8),wraplength=card_w).pack(padx=8,pady=20)
        for col in range(ncols): self.preview_frame.grid_columnconfigure(col,weight=1)
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
    def _show_empty(self):
        for w in self.cards_frame.winfo_children(): w.destroy()
        tk.Label(self.cards_frame,text="No files loaded",bg=SURFACE,fg=MUTED,font=("Segoe UI",9)).pack(fill="x",padx=12,pady=35)
        self._schedule_preview()
    def _add_files(self):
        exts=" ".join(f"*{e}" for e in (IMAGE_EXTS | {".pdf"}))
        self._drop_files(list(filedialog.askopenfilenames(filetypes=[("Images and PDFs",exts),("All","*.*")])))
    def _drop_files(self,paths):
        changed=False
        for p in paths:
            ext=os.path.splitext(p)[1].lower()
            if ext==".pdf":
                try:
                    doc=fitz.open(p)
                    for page_no in range(len(doc)):
                        key=f"{p}::page::{page_no}"
                        if any(self.pdf_sources.get(existing)==(p,page_no) for existing in self.files): continue
                        page=doc[page_no]; pix=page.get_pixmap(matrix=fitz.Matrix(60/72,60/72),alpha=False)
                        img=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                        preview=os.path.join(self._temp_dir,f"pdf_{abs(hash(key))}_{page_no+1}.png"); img.save(preview,"PNG")
                        self.files.append(preview); self.pdf_sources[preview]=(p,page_no); self.display_names[preview]=f"{os.path.basename(p)} — page {page_no+1}"; changed=True
                    doc.close()
                except Exception as ex: messagebox.showerror("PDF import failed",f"{os.path.basename(p)}\n{ex}")
            elif ext in IMAGE_EXTS and p not in self.files:
                self.files.append(p); self.display_names[p]=os.path.basename(p); changed=True
        if changed:self._rebuild()
    def _clear(self):
        self._preview_item_cache.clear()
        for photo in self.preview_images:
            try:photo.__del__()
            except Exception:pass
        self.preview_images.clear()
        for path in list(self.pdf_sources):
            try:os.remove(path)
            except Exception:pass
        self.files.clear(); self.canvases.clear(); self.pdf_sources.clear(); self.display_names.clear(); self._show_empty(); gc.collect()
    def _remove(self,path):
        if path in self.files:self.files.remove(path)
        if path in self.pdf_sources:
            try:os.remove(path)
            except Exception:pass
        self.pdf_sources.pop(path,None); self.display_names.pop(path,None); self._rebuild()
    def _move_source(self,path,direction):
        if path not in self.files:return
        i=self.files.index(path); j=i+direction
        if 0<=j<len(self.files):
            self.files[i],self.files[j]=self.files[j],self.files[i]
            self._rebuild()
    def _rebuild(self):
        self._preview_item_cache.clear()
        for w in self.cards_frame.winfo_children():w.destroy()
        self.canvases.clear()
        if not self.files:self._show_empty();return
        preview_dpi=72; tw=mm_to_px(max(.1,self._num(self.item_w,90)),preview_dpi); th=mm_to_px(max(.1,self._num(self.item_h,50)),preview_dpi)
        for i,path in enumerate(self.files):
            cell=tk.Frame(
                self.cards_frame,
                bg=SURFACE2,
                padx=7,
                pady=7,
                highlightbackground=BORDER,
                highlightthickness=1,
            )
            cell.pack(fill="x",padx=4,pady=4)
            try:
                # The source preview is deliberately sized below the right-panel
                # width and allowed to keep its natural card height, preventing
                # crop controls, filenames, or buttons from being clipped.
                cc=PrintCropCanvas(
                    cell,
                    path,
                    tw,
                    th,
                    size=(220,105),
                    on_change=self._schedule_preview,
                )
                cc.pack(anchor="center")
                cc.set_effects(
                    self.gray_var.get(),
                    self.brightness_var.get(),
                )
                self.canvases.append((path,cc))

                tk.Label(
                    cell,
                    text=self.display_names.get(path,os.path.basename(path)),
                    bg=SURFACE2,
                    fg=MUTED,
                    font=("Segoe UI",8),
                    wraplength=220,
                    justify="left",
                    anchor="w",
                ).pack(fill="x",pady=(5,0))

                acts=tk.Frame(cell,bg=SURFACE2)
                acts.pack(fill="x",pady=(5,0))
                buttons=(
                    ("↑",lambda p=path:self._move_source(p,-1)),
                    ("↓",lambda p=path:self._move_source(p,1)),
                    ("⟲",lambda c=cc:c.rotate_manual(90)),
                    ("⟳",lambda c=cc:c.rotate_manual(-90)),
                    ("✕",lambda p=path:self._remove(p)),
                )
                for t,cmd in buttons:
                    tk.Button(
                        acts,
                        text=t,
                        command=cmd,
                        bg=SURFACE,
                        fg=TEXT,
                        relief="flat",
                        font=("Segoe UI",8,"bold"),
                        cursor="hand2",
                    ).pack(
                        side="left",
                        fill="x",
                        expand=True,
                        padx=1,
                    )
            except Exception as ex:
                tk.Label(
                    cell,
                    text=str(ex)[:100],
                    bg=SURFACE2,
                    fg=DANGER,
                    wraplength=220,
                ).pack(fill="x",padx=4,pady=12)
        self._schedule_preview()
        gc.collect(0)
    def _current_layout(self):
        return {"paper":self.paper_var.get(),"orientation":self.orientation_var.get(),"item_w":self._num(self.item_w,90),"item_h":self._num(self.item_h,50),"cols":int(self._num(self.cols_var,2)),"rows":int(self._num(self.rows_var,5)),"mode":self.mode_var.get(),"duplex":self.duplex_var.get(),"bleed":self.bleed_var.get(),"bleed_mm":self._num(self.bleed_mm,3),"spacing":self._num(self.spacing_var,6),"grayscale":self.gray_var.get(),"brightness":int(self.brightness_var.get()),"dpi":int(self.dpi_var.get()),"cut_marks":self.cut_marks_var.get(),"cut_labels":self.cut_labels_var.get(),"proportional":self.proportional_var.get(),"proportion_ratio":self._proportion_ratio}
    def _load_layout(self,name):
        d=self.layouts.get(name)
        if not d:return
        self.layout_var.set(name); self.paper_var.set(d.get("paper","A4")); self.orientation_var.set(d.get("orientation","Portrait")); self._set_item_size(d.get("item_w",90),d.get("item_h",50),False); self.cols_var.set(d.get("cols",2)); self.rows_var.set(d.get("rows",5)); self.mode_var.set(d.get("mode","Repeat")); self.duplex_var.set(d.get("duplex",False)); self.bleed_var.set(d.get("bleed",False)); self.bleed_mm.set(d.get("bleed_mm",3)); self.spacing_var.set(d.get("spacing",0)); self.gray_var.set(d.get("grayscale",False)); self.brightness_var.set(d.get("brightness",100)); self.dpi_var.set(d.get("dpi",300)); self.cut_marks_var.set(d.get("cut_marks",False)); self.cut_labels_var.set(d.get("cut_labels",False)); self.proportional_var.set(bool(d.get("proportional",False))); self._proportion_ratio=d.get("proportion_ratio"); self._update_chain_button(); self._toggle_cut_marks(); self.brightness_lbl.config(text=f"Brightness: {self.brightness_var.get()}%"); self._toggle_bleed(); self._mode_changed()
    def _save_layout_dialog(self):
        name=simpledialog.askstring("Save layout","Layout name:",initialvalue=self.layout_var.get(),parent=self)
        if not name:return
        name=name.strip(); self.layouts[name]=self._current_layout(); save_print_layouts(self.layouts); self.layout_cb.config(values=list(self.layouts)); self.layout_var.set(name); self.status.config(text=f"Saved layout: {name}",fg=SUCCESS)
    def _processed(self,path,cc,target_px=None,preview=False):
        if path in self.pdf_sources:
            pdf_path,page_no=self.pdf_sources[path]
            doc=fitz.open(pdf_path)
            try:
                page=doc[page_no]; l,t,r,b=cc.get_crop_fractions()
                frac_w=max(.001,r-l); frac_h=max(.001,b-t)
                page_w,page_h=page.rect.width,page.rect.height
                if cc.total_rotation in (90,270):page_w,page_h=page_h,page_w
                if target_px:
                    target_w,target_h=target_px; multiplier=1.15 if preview else 2
                    scale=max((target_w*multiplier)/(page_w*frac_w),(target_h*multiplier)/(page_h*frac_h))
                    render_dpi=(max(45,min(100,int(math.ceil(scale*72)))) if preview else max(int(self.dpi_var.get()),min(1200,int(math.ceil(scale*72)))))
                else:render_dpi=90 if preview else max(600,int(self.dpi_var.get())*2)
                pix=page.get_pixmap(matrix=fitz.Matrix(render_dpi/72,render_dpi/72),alpha=False)
                raw=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
            finally:doc.close()
        else:
            with Image.open(path) as opened:
                corrected=ImageOps.exif_transpose(opened)
                if preview:
                    limit=max(600,min(1400,max(target_px or (900,900))*2))
                    corrected.thumbnail((limit,limit),Image.Resampling.BILINEAR)
                raw=corrected.convert("RGB")
                if corrected is not opened:
                    try:corrected.close()
                    except Exception:pass
        if cc.total_rotation:
            rotated=raw.rotate(cc.total_rotation,expand=True); raw.close(); raw=rotated
        iw,ih=raw.size; l,t,r,b=cc.get_crop_fractions()
        cropped=raw.crop((int(l*iw),int(t*ih),int(r*iw),int(b*ih))); raw.close(); raw=cropped
        if self.gray_var.get():
            converted=ImageOps.grayscale(raw).convert("RGB"); raw.close(); raw=converted
        if self.brightness_var.get()!=100:
            adjusted=ImageEnhance.Brightness(raw).enhance(self.brightness_var.get()/100); raw.close(); raw=adjusted
        return raw
    def _item(self,path,cc,dpi,iw,ih,bleed,preview=False):
        target=(mm_to_px(iw,dpi),mm_to_px(ih,dpi)); key=None
        if preview:
            key=(path,tuple(round(v,4) for v in cc.get_crop_fractions()),int(cc.total_rotation),target,round(float(bleed),2),bool(self.gray_var.get()),int(self.brightness_var.get()))
            cached=self._preview_item_cache.get(key)
            if cached is not None:return cached
        img=self._processed(path,cc,target,preview)
        if img.width and img.height and (img.width>img.height)!=(target[0]>target[1]):
            rotated=img.rotate(90,expand=True); img.close(); img=rotated
        fitted=ImageOps.fit(img,target,method=(Image.Resampling.BILINEAR if preview else Image.Resampling.LANCZOS),centering=(.5,.5)); img.close(); img=fitted
        if bleed>0:
            finished=add_bleed_and_marks(img,bleed,dpi,False); img.close(); img=finished
        if preview:self._preview_item_cache[key]=img
        return img

    def _sheet(self,assignments,back=False,dpi_override=None,preview_guides=False):
        pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,gw,gh=self._metrics(); dpi=int(dpi_override or self.dpi_var.get()); sheet=Image.new("RGB",(mm_to_px(pw,dpi),mm_to_px(ph,dpi)),"white"); by={p:c for p,c in self.canvases}
        for slot,path in enumerate(assignments):
            if not path or path not in by:continue
            row=slot//cols; col=slot%cols
            if back:col=cols-1-col
            is_preview=dpi_override is not None
            item=self._item(path,by[path],dpi,iw,ih,bleed,preview=is_preview)
            tx=x0+col*(iw+gap)-bleed; ty=y0+row*(ih+gap)-bleed
            sheet.paste(item,(mm_to_px(tx,dpi),mm_to_px(ty,dpi)))
            if not is_preview:
                try:item.close()
                except Exception:pass
        if self.cut_marks_var.get(): self._draw_cut_guides(sheet,pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,dpi)
        if preview_guides:
            draw=ImageDraw.Draw(sheet)
            page_lw=max(1,mm_to_px(.35,dpi))
            draw.rectangle((0,0,sheet.width-1,sheet.height-1),outline=(65,65,65),width=page_lw)
            trim_lw=max(1,mm_to_px(.25,dpi))
            bleed_lw=max(1,mm_to_px(.2,dpi))
            for row in range(rows):
                for col in range(cols):
                    left=x0+col*(iw+gap); top=y0+row*(ih+gap)
                    x1,y1=mm_to_px(left,dpi),mm_to_px(top,dpi)
                    x2,y2=mm_to_px(left+iw,dpi),mm_to_px(top+ih,dpi)
                    draw.rectangle((x1,y1,x2,y2),outline=(0,90,210),width=trim_lw)
                    if bleed>0:
                        bx1,by1=mm_to_px(left-bleed,dpi),mm_to_px(top-bleed,dpi)
                        bx2,by2=mm_to_px(left+iw+bleed,dpi),mm_to_px(top+ih+bleed,dpi)
                        draw.rectangle((bx1,by1,bx2,by2),outline=(210,45,45),width=bleed_lw)
        return sheet
    def _draw_cut_guides(self,sheet,pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,dpi):
        draw=ImageDraw.Draw(sheet)
        safe_mm=4.0
        desired_mm=5.0
        artwork_gap_mm=1.0
        safe=mm_to_px(safe_mm,dpi)
        lw=max(1,mm_to_px(.25,dpi))
        def label_font(mark_len_px):
            # Make label character height visually comparable to the cut-mark length,
            # so labels remain readable when larger sheets allow longer marks.
            font_size=max(14,int(max(mark_len_px,mm_to_px(3.2,dpi))*0.9))
            try: return ImageFont.truetype("DejaVuSans-Bold.ttf",font_size)
            except Exception:
                try: return ImageFont.truetype("DejaVuSans.ttf",font_size)
                except Exception: return ImageFont.load_default()

        xs=sorted(set([x0+c*(iw+gap) for c in range(cols)]+[x0+c*(iw+gap)+iw for c in range(cols)]))
        ys=sorted(set([y0+r*(ih+gap) for r in range(rows)]+[y0+r*(ih+gap)+ih for r in range(rows)]))
        sw,sh=sheet.size
        layout_right=x0+cols*iw+max(0,cols-1)*gap
        layout_bottom=y0+rows*ih+max(0,rows-1)*gap
        top_free=max(0,(y0-bleed-artwork_gap_mm)-safe_mm)
        bottom_free=max(0,(ph-safe_mm)-(layout_bottom+bleed+artwork_gap_mm))
        left_free=max(0,(x0-bleed-artwork_gap_mm)-safe_mm)
        right_free=max(0,(pw-safe_mm)-(layout_right+bleed+artwork_gap_mm))
        top_len=mm_to_px(min(desired_mm,top_free),dpi)
        bottom_len=mm_to_px(min(desired_mm,bottom_free),dpi)
        left_len=mm_to_px(min(desired_mm,left_free),dpi)
        right_len=mm_to_px(min(desired_mm,right_free),dpi)
        label_gap=max(2,mm_to_px(.8,dpi))

        def draw_horizontal_label(cx,line_y,txt,mark_len):
            font=label_font(mark_len)
            box=draw.textbbox((0,0),txt,font=font)
            tw=box[2]-box[0]; th=box[3]-box[1]
            draw.text((int(cx-tw/2),int(line_y-label_gap-th-box[1])),txt,fill="black",font=font)

        def draw_vertical_label(line_x,cy,txt,mark_len):
            font=label_font(mark_len)
            box=draw.textbbox((0,0),txt,font=font)
            tw=box[2]-box[0]; th=box[3]-box[1]
            pad=2
            tile=Image.new("RGBA",(tw+pad*2,th+pad*2),(255,255,255,0))
            td=ImageDraw.Draw(tile)
            td.text((pad-box[0],pad-box[1]),txt,fill="black",font=font)
            tile=tile.rotate(90,expand=True,resample=Image.Resampling.BICUBIC)
            sheet.paste(tile,(int(line_x+label_gap),int(cy-tile.height/2)),tile)

        for x in xs:
            px=mm_to_px(x,dpi)
            txt=f"{min(x,pw-x):.1f}"
            if top_len>0:
                y1=safe; y2=safe+top_len
                draw.line((px,y1,px,y2),fill="black",width=lw)
                if self.cut_labels_var.get(): draw_vertical_label(px,(y1+y2)/2,txt,top_len)
            if bottom_len>0:
                y2=sh-safe; y1=y2-bottom_len
                draw.line((px,y1,px,y2),fill="black",width=lw)
                if self.cut_labels_var.get(): draw_vertical_label(px,(y1+y2)/2,txt,bottom_len)

        for y in ys:
            py=mm_to_px(y,dpi)
            txt=f"{min(y,ph-y):.1f}"
            if left_len>0:
                x1=safe; x2=safe+left_len
                draw.line((x1,py,x2,py),fill="black",width=lw)
                if self.cut_labels_var.get(): draw_horizontal_label((x1+x2)/2,py,txt,left_len)
            if right_len>0:
                x2=sw-safe; x1=x2-right_len
                draw.line((x1,py,x2,py),fill="black",width=lw)
                if self.cut_labels_var.get(): draw_horizontal_label((x1+x2)/2,py,txt,right_len)
    def _page_specs(self):
        slots=max(1,int(self.cols_var.get())*int(self.rows_var.get()))
        mode=self.mode_var.get(); duplex=self.duplex_var.get(); specs=[]
        sheet_no=1
        if mode=="Repeat":
            step=2 if duplex else 1
            for i in range(0,len(self.files),step):
                specs.append(([self.files[i]]*slots,False,f"Sheet {sheet_no} — front" if duplex else f"Sheet {sheet_no}"))
                if duplex:
                    back_path=self.files[i+1] if i+1<len(self.files) else None
                    specs.append(([back_path]*slots,True,f"Sheet {sheet_no} — back"))
                sheet_no+=1
        elif mode=="Cut & stack":
            pairs=[self.files[i:i+2] for i in range(0,len(self.files),2)]
            ns=math.ceil(len(pairs)/slots) if pairs else 0
            for p in range(ns):
                front=[]; back=[]
                for s in range(slots):
                    idx=s*ns+p; pair=pairs[idx] if idx<len(pairs) else []
                    front.append(pair[0] if pair else None)
                    back.append(pair[1] if len(pair)>1 else None)
                specs.append((front,False,f"Sheet {sheet_no} — front"))
                specs.append((back,True,f"Sheet {sheet_no} — back"))
                sheet_no+=1
        else:
            for i in range(0,len(self.files),slots):
                chunk=self.files[i:i+slots]+[None]*max(0,slots-len(self.files[i:i+slots]))
                specs.append((chunk,False,f"Sheet {sheet_no}")); sheet_no+=1
        return specs
    def _pages(self):
        return [self._sheet(assignments,back=back) for assignments,back,_label in self._page_specs()]
    def _write_layout_pdf(self,out):
        pdf=fitz.open(); pw,ph=self._paper_mm()
        try:
            for index,(assignments,back,_label) in enumerate(self._page_specs()):
                sheet=self._sheet(assignments,back=back); buffer=io.BytesIO()
                try:
                    sheet.save(buffer,"PNG",compress_level=3)
                    page=pdf.new_page(width=pw*72/25.4,height=ph*72/25.4)
                    page.insert_image(page.rect,stream=buffer.getvalue())
                finally:
                    buffer.close(); sheet.close()
                if index and index%4==0:gc.collect(0)
            if not len(pdf):raise ValueError("No output pages were generated.")
            pdf.save(out,garbage=3,deflate=True)
        finally:
            pdf.close(); gc.collect()

    def _ctrl_p(self,event=None):
        focused=self.focus_get()
        if isinstance(focused,(tk.Entry,ttk.Entry,tk.Text)):
            return None
        self._print_current_layout()
        return "break"
    def _print_current_layout(self):
        if not self.canvases:
            messagebox.showwarning("Nothing to print","Add images or PDF pages first.")
            return
        pw,ph,*rest=self._metrics(); gw,gh=rest[-2:]
        if gw>pw+1e-6 or gh>ph+1e-6:
            messagebox.showerror("Layout does not fit","The centered grid is larger than the selected paper.")
            return
        self.status.config(text="Preparing print PDF…",fg=MUTED)
        threading.Thread(target=self._prepare_print_pdf,daemon=True).start()
    def _prepare_print_pdf(self):
        try:
            out=os.path.join(tempfile.gettempdir(),f"CopyPro_Print_{int(time.time())}.pdf")
            self._write_layout_pdf(out)
            self.after(0,lambda p=out:self._open_acrobat_print(p))
        except Exception as ex:
            self.after(0,lambda e=ex:messagebox.showerror("Print failed",str(e)))
            self.after(0,lambda:self.status.config(text="Print preparation failed",fg=DANGER))
    def _open_acrobat_print(self,pdf_path):
        self.status.config(text="Opening Acrobat print dialog…",fg=SUCCESS)
        if not sys.platform.startswith("win"):
            try: subprocess.Popen([pdf_path])
            except Exception: pass
            return
        candidates=[
            os.path.expandvars(r"%ProgramFiles%\Adobe\Acrobat DC\Acrobat\Acrobat.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
            os.path.expandvars(r"%ProgramFiles%\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Adobe\Reader 11.0\Reader\AcroRd32.exe"),
        ]
        exe=next((p for p in candidates if p and os.path.exists(p)),None)
        try:
            if exe:
                subprocess.Popen([exe,pdf_path])
            else:
                os.startfile(pdf_path)
            # Acrobat has no reliable documented command that only opens its print
            # dialog. Activate it after opening and send Ctrl+P as a best-effort step.
            ps=("Start-Sleep -Milliseconds 1400; "
                "$w=New-Object -ComObject WScript.Shell; "
                "if($w.AppActivate('Adobe Acrobat') -or $w.AppActivate('Adobe Acrobat Reader'))"
                "{Start-Sleep -Milliseconds 250; $w.SendKeys('^p')}")
            subprocess.Popen(["powershell","-NoProfile","-WindowStyle","Hidden","-Command",ps],creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0))
        except Exception as ex:
            messagebox.showerror("Print failed",f"The PDF was created, but Acrobat could not be opened.\n\n{ex}\n\nFile: {pdf_path}")
    def _export_pdf(self):
        if not self.canvases:messagebox.showwarning("Nothing to export","Add images first.");return
        pw,ph,*rest=self._metrics(); gw,gh=rest[-2:]
        if gw>pw+1e-6 or gh>ph+1e-6:messagebox.showerror("Layout does not fit","The centered grid is larger than the selected paper.");return
        out=filedialog.asksaveasfilename(title="Export print layout",defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile=(str(self.paper_var.get()).replace(" ","_")+f"_{self.item_w.get():g}x{self.item_h.get():g}.pdf"))
        if not out:return
        self.export_btn.config(state="disabled"); self.status.config(text="Exporting…",fg=MUTED); threading.Thread(target=self._do_export,args=(out,),daemon=True).start()
    def _do_export(self,out):
        try:
            page_count=len(self._page_specs()); self._write_layout_pdf(out)
            self.after(0,lambda:self.status.config(text=f"Exported {page_count} PDF page(s)",fg=SUCCESS)); self.after(0,lambda:messagebox.showinfo("Done",f"Print layout exported to:\n{out}"))
        except Exception as ex:
            self.after(0,lambda e=ex:messagebox.showerror("Export failed",str(e))); self.after(0,lambda:self.status.config(text="Export failed",fg=DANGER))
        finally:self.after(0,lambda:self.export_btn.config(state="normal"))

    def clear_memory_cache(self):
        self._preview_item_cache.clear()
        for photo in self.preview_images:
            try:photo.__del__()
            except Exception:pass
        self.preview_images.clear(); gc.collect(); self._schedule_preview()

    def destroy(self):
        try:
            self.unbind_all("<Control-p>"); self.unbind_all("<Control-P>")
            if self._preview_job:self.after_cancel(self._preview_job)
            refresh=getattr(self,"_copypro_drop_refresh_job",None)
            if refresh:self.after_cancel(refresh)
        except Exception: pass
        self._preview_item_cache.clear()
        for photo in self.preview_images:
            try:photo.__del__()
            except Exception:pass
        self.preview_images.clear()
        try: shutil.rmtree(self._temp_dir,ignore_errors=True)
        except Exception: pass
        gc.collect(); super().destroy()




# ── Polaroid Maker Tab ───────────────────────────────────────────────────────
class PolaroidMakerTab(tk.Frame,DropMixin):
    """Arrange movable photo crops inside white instant-photo frames on SRA sheets."""
    PRESETS=OrderedDict([
        ("Classic 88 × 107 mm",(88.0,107.0,4.5,5.0,25.0,"Square")),
        ("Mini 54 × 86 mm",(54.0,86.0,4.0,4.0,20.0,"3:4")),
        ("Square 72 × 86 mm",(72.0,86.0,5.0,5.0,19.0,"Square")),
        ("Wide 108 × 86 mm",(108.0,86.0,4.5,5.0,19.0,"3:2")),
        ("Custom",(88.0,107.0,5.0,5.0,25.0,"Square")),
    ])
    RATIOS={"Square":1.0,"3:4":3/4,"2:3":2/3,"3:2":3/2,"4:5":4/5}

    def __init__(self,parent):
        super().__init__(parent,bg=BG)
        self.files=[]; self.canvases=[]; self.display_names={}
        self._temp_dir=tempfile.mkdtemp(prefix="copypro_polaroid_")
        self._generated=[]; self._preview_job=None; self.preview_photo=None
        self._build(); self._preset_changed()

    def _build(self):
        self.grid_rowconfigure(0,weight=4)
        self.grid_rowconfigure(1,weight=1,minsize=190)
        self.grid_columnconfigure(0,weight=0)
        self.grid_columnconfigure(1,weight=1)

        left=tk.Frame(self,bg=SURFACE,width=278)
        left.grid(row=0,column=0,rowspan=2,sticky="nsew")
        left.grid_propagate(False)
        lbl(left,"POLAROID MAKER",9,MUTED,True).pack(anchor="w",padx=14,pady=(12,5))
        tk.Label(left,text="Create white instant-photo frames and arrange them automatically on SRA4 or SRA3.",
                 bg=SURFACE,fg=MUTED,font=("Segoe UI",8),wraplength=245,justify="left").pack(anchor="w",padx=14,pady=(0,8))

        lbl(left,"Frame preset",8).pack(anchor="w",padx=14)
        self.preset_var=tk.StringVar(value="Classic 88 × 107 mm")
        self.preset_cb=ttk.Combobox(left,textvariable=self.preset_var,values=list(self.PRESETS),state="readonly")
        self.preset_cb.pack(fill="x",padx=14,pady=(2,5)); self.preset_cb.bind("<<ComboboxSelected>>",lambda e:self._preset_changed())

        dimensions=tk.Frame(left,bg=SURFACE); dimensions.pack(fill="x",padx=14)
        self.outer_w_var=tk.StringVar(value="88"); self.outer_h_var=tk.StringVar(value="107")
        for i,(caption,var) in enumerate((("Width mm",self.outer_w_var),("Height mm",self.outer_h_var))):
            f=tk.Frame(dimensions,bg=SURFACE); f.grid(row=0,column=i,sticky="ew",padx=(0 if i==0 else 5,0)); dimensions.grid_columnconfigure(i,weight=1)
            lbl(f,caption,7,MUTED).pack(anchor="w"); entry(f,var,7).pack(fill="x")

        lbl(left,"Image crop ratio",8).pack(anchor="w",padx=14,pady=(7,0))
        self.ratio_var=tk.StringVar(value="Square")
        self.ratio_cb=ttk.Combobox(left,textvariable=self.ratio_var,
            values=["Square","3:4","2:3","3:2","4:5","Original","Custom"],state="readonly")
        self.ratio_cb.pack(fill="x",padx=14,pady=(2,2)); self.ratio_cb.bind("<<ComboboxSelected>>",lambda e:self._ratio_changed())
        self.custom_ratio_frame=tk.Frame(left,bg=SURFACE)
        lbl(self.custom_ratio_frame,"Custom W:H",7,MUTED).pack(side="left")
        self.custom_ratio_w=tk.StringVar(value="4"); self.custom_ratio_h=tk.StringVar(value="5")
        entry(self.custom_ratio_frame,self.custom_ratio_w,5).pack(side="left",padx=(6,3)); lbl(self.custom_ratio_frame,":",8,MUTED).pack(side="left"); entry(self.custom_ratio_frame,self.custom_ratio_h,5).pack(side="left",padx=(3,0))

        sep(left); lbl(left,"WHITE FRAME MARGINS",8,MUTED,True).pack(anchor="w",padx=14)
        self.side_margin_var=tk.StringVar(value="4.5"); self.top_margin_var=tk.StringVar(value="5"); self.bottom_margin_var=tk.StringVar(value="25")
        margins=tk.Frame(left,bg=SURFACE); margins.pack(fill="x",padx=14,pady=(3,2))
        for i,(caption,var) in enumerate((("Sides",self.side_margin_var),("Top",self.top_margin_var),("Bottom",self.bottom_margin_var))):
            f=tk.Frame(margins,bg=SURFACE); f.grid(row=0,column=i,sticky="ew",padx=(0 if i==0 else 4,0)); margins.grid_columnconfigure(i,weight=1)
            lbl(f,caption+" mm",7,MUTED).pack(anchor="w"); entry(f,var,5).pack(fill="x")
        styled_btn(left,"Apply frame settings",self._apply_frame_settings,style="secondary").pack(fill="x",padx=14,pady=(4,5))

        sep(left); lbl(left,"SHEET LAYOUT",8,MUTED,True).pack(anchor="w",padx=14)
        sheet_row=tk.Frame(left,bg=SURFACE); sheet_row.pack(fill="x",padx=14,pady=(3,2))
        self.sheet_var=tk.StringVar(value="SRA4")
        ttk.Combobox(sheet_row,textvariable=self.sheet_var,values=["SRA4","SRA3"],state="readonly",width=8).pack(side="left",fill="x",expand=True)
        self.spacing_var=tk.StringVar(value="3")
        lbl(sheet_row,"Spacing",7,MUTED).pack(side="left",padx=(7,3)); entry(sheet_row,self.spacing_var,5).pack(side="left"); lbl(sheet_row,"mm",7,MUTED).pack(side="left",padx=(2,0))

        self.auto_layout_var=tk.BooleanVar(value=True)
        tk.Checkbutton(left,text="Automatic rows, columns and rotation",variable=self.auto_layout_var,
            bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._layout_mode_changed).pack(anchor="w",padx=14,pady=(3,1))
        manual=tk.Frame(left,bg=SURFACE); manual.pack(fill="x",padx=14)
        self.cols_var=tk.IntVar(value=2); self.rows_var=tk.IntVar(value=2)
        self.manual_spins=[]
        for i,(caption,var) in enumerate((("Columns",self.cols_var),("Rows",self.rows_var))):
            f=tk.Frame(manual,bg=SURFACE); f.grid(row=0,column=i,sticky="ew",padx=(0 if i==0 else 5,0)); manual.grid_columnconfigure(i,weight=1)
            lbl(f,caption,7,MUTED).pack(anchor="w")
            sp=tk.Spinbox(f,from_=1,to=20,textvariable=var,bg=SURFACE2,fg=TEXT,insertbackground=TEXT,
                          buttonbackground=SURFACE2,relief="flat",font=("Segoe UI",9),command=self._schedule_preview)
            sp.pack(fill="x"); sp.bind("<KeyRelease>",lambda e:self._schedule_preview()); self.manual_spins.append(sp)
        lbl(left,"Cut guide",7,MUTED).pack(anchor="w",padx=14,pady=(5,0))
        self.mark_mode_var=tk.StringVar(value="Crop marks")
        self.mark_mode_cb=ttk.Combobox(
            left,textvariable=self.mark_mode_var,
            values=["Crop marks","Outline 0.15 mm"],state="readonly"
        )
        self.mark_mode_cb.pack(fill="x",padx=14,pady=(2,3))
        self.mark_mode_cb.bind("<<ComboboxSelected>>",lambda _event:self._schedule_preview())

        actions=tk.Frame(left,bg=SURFACE); actions.pack(fill="x",padx=14,pady=(4,2))
        styled_btn(actions,"Add images",self._add_files).pack(side="left",fill="x",expand=True)
        styled_btn(actions,"Clear",self._clear,style="secondary").pack(side="left",fill="x",expand=True,padx=(5,0))
        self.export_btn=styled_btn(left,"Export print-ready PDF",self._export_pdf,style="success")
        self.export_btn.pack(fill="x",padx=14,pady=(4,3))
        self.status=lbl(left,"Add images or a PDF.",8,MUTED); self.status.pack(fill="x",padx=14,pady=(0,10))

        # The sheet preview is the primary workspace and gets most of the tab.
        preview_panel=tk.Frame(self,bg=SURFACE,highlightbackground=BORDER,highlightthickness=1)
        preview_panel.grid(row=0,column=1,sticky="nsew",padx=8,pady=(8,4))
        preview_panel.grid_rowconfigure(1,weight=1)
        preview_panel.grid_columnconfigure(0,weight=1)
        hdr=tk.Frame(preview_panel,bg=SURFACE)
        hdr.grid(row=0,column=0,sticky="ew",padx=12,pady=(9,5))
        lbl(hdr,"SHEET PREVIEW",9,MUTED,True).pack(side="left")
        self.preview_info=lbl(hdr,"",8,MUTED); self.preview_info.pack(side="right")
        self.preview_canvas=tk.Canvas(preview_panel,bg=BG,highlightthickness=0)
        self.preview_canvas.grid(row=1,column=0,sticky="nsew",padx=8,pady=(0,8))
        self.preview_canvas.bind("<Configure>",lambda _event:self._schedule_preview())

        # Smaller crop cards sit in a horizontal strip under the sheet preview.
        crops_panel=tk.Frame(self,bg=BG)
        crops_panel.grid(row=1,column=1,sticky="nsew",padx=8,pady=(4,8))
        crops_panel.grid_rowconfigure(1,weight=1)
        crops_panel.grid_columnconfigure(0,weight=1)
        head=tk.Frame(crops_panel,bg=BG)
        head.grid(row=0,column=0,sticky="ew",pady=(0,4))
        lbl(head,"IMAGE CROPS",9,MUTED,True).pack(side="left")
        self.count_lbl=lbl(head,"0 images",8,MUTED); self.count_lbl.pack(side="right")
        self.cards_canvas=tk.Canvas(crops_panel,bg=BG,highlightthickness=0,height=158)
        self.cards_canvas.grid(row=1,column=0,sticky="nsew")
        crop_scroll=ttk.Scrollbar(crops_panel,orient="horizontal",command=self.cards_canvas.xview)
        crop_scroll.grid(row=2,column=0,sticky="ew",pady=(3,0))
        self.cards_canvas.config(xscrollcommand=crop_scroll.set)
        self.cards_frame=tk.Frame(self.cards_canvas,bg=BG)
        self.cards_win=self.cards_canvas.create_window((0,0),window=self.cards_frame,anchor="nw")
        self.cards_frame.bind("<Configure>",lambda _event:self.cards_canvas.config(scrollregion=self.cards_canvas.bbox("all")))
        self.cards_canvas.bind("<Configure>",lambda event:self.cards_canvas.itemconfig(self.cards_win,height=max(1,event.height)))
        def crop_wheel(event):
            self.cards_canvas.xview_scroll(int(-event.delta/120),"units")
            return "break"
        self.cards_canvas.bind("<Enter>",lambda _event:self.cards_canvas.bind_all("<MouseWheel>",crop_wheel))
        self.cards_canvas.bind("<Leave>",lambda _event:self.cards_canvas.unbind_all("<MouseWheel>"))

        self._layout_mode_changed(); self._show_empty()

    def _num(self,var,default):
        try:return float(str(var.get()).replace(",","."))
        except Exception:return default

    def _preset_changed(self):
        w,h,side,top,bottom,ratio=self.PRESETS.get(self.preset_var.get(),self.PRESETS["Classic 88 × 107 mm"])
        self.outer_w_var.set(f"{w:g}"); self.outer_h_var.set(f"{h:g}")
        self.side_margin_var.set(f"{side:g}"); self.top_margin_var.set(f"{top:g}"); self.bottom_margin_var.set(f"{bottom:g}")
        self.ratio_var.set(ratio); self._ratio_changed(); self._apply_frame_settings()

    def _ratio_changed(self):
        if self.ratio_var.get()=="Custom":self.custom_ratio_frame.pack(fill="x",padx=14,pady=(1,3),after=self.ratio_cb)
        else:self.custom_ratio_frame.pack_forget()

    def _ratio_for(self,path=None):
        name=self.ratio_var.get()
        if name in self.RATIOS:return self.RATIOS[name]
        if name=="Custom":
            w=self._num(self.custom_ratio_w,4); h=self._num(self.custom_ratio_h,5)
            return max(.05,w/max(.05,h))
        if name=="Original" and path:
            canvas=next((c for p,c in self.canvases if p==path),None)
            if canvas is not None:return canvas.pil_img.width/max(1,canvas.pil_img.height)
            try:
                with Image.open(path) as im:
                    corrected=ImageOps.exif_transpose(im); return corrected.width/max(1,corrected.height)
            except Exception:pass
        return 1.0

    def _frame_values(self,path=None):
        ow=max(10,self._num(self.outer_w_var,88)); oh=max(10,self._num(self.outer_h_var,107))
        side=max(0,self._num(self.side_margin_var,4.5)); top=max(0,self._num(self.top_margin_var,5)); bottom=max(0,self._num(self.bottom_margin_var,25))
        aw=max(1,ow-side*2); ah=max(1,oh-top-bottom); ratio=max(.05,self._ratio_for(path))
        iw=aw; ih=iw/ratio
        if ih>ah:ih=ah; iw=ih*ratio
        x=(ow-iw)/2; y=top
        return ow,oh,x,y,iw,ih

    def _layout_mode_changed(self):
        state="disabled" if self.auto_layout_var.get() else "normal"
        for sp in self.manual_spins:sp.config(state=state)
        self._schedule_preview()

    def _apply_frame_settings(self):
        self._rebuild_cards(); self._schedule_preview()

    def _expand_pdf(self,path):
        result=[]; doc=fitz.open(path)
        try:
            for page_no,page in enumerate(doc):
                saved=None
                images=page.get_images(full=True)
                if len(images)==1:
                    try:
                        data=doc.extract_image(images[0][0]); ext=data.get("ext","png")
                        saved=os.path.join(self._temp_dir,f"pdf_{len(self._generated)}_{page_no}.{ext}")
                        Path(saved).write_bytes(data["image"])
                    except Exception:saved=None
                if not saved:
                    pix=page.get_pixmap(matrix=fitz.Matrix(300/72,300/72),alpha=False)
                    image=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                    saved=os.path.join(self._temp_dir,f"pdf_{len(self._generated)}_{page_no}.png")
                    try:image.save(saved,"PNG",compress_level=2)
                    finally:image.close()
                self._generated.append(saved); self.display_names[saved]=f"{os.path.basename(path)} — page {page_no+1}"; result.append(saved)
        finally:doc.close()
        return result

    def _add_files(self,paths=None):
        if paths is None:
            paths=filedialog.askopenfilenames(parent=self,title="Add photos",filetypes=[("Images and PDF","*.jpg *.jpeg *.png *.tif *.tiff *.bmp *.webp *.heic *.heif *.pdf"),("All files","*.*")])
        added=[]
        for path in paths or []:
            ext=os.path.splitext(path)[1].lower()
            if ext==".pdf":
                try:added.extend(self._expand_pdf(path))
                except Exception as ex:messagebox.showerror("PDF import failed",str(ex),parent=self)
            elif ext in IMAGE_EXTS:added.append(path); self.display_names.setdefault(path,os.path.basename(path))
        for path in added:
            if path not in self.files:self.files.append(path)
        self._rebuild_cards(); self._schedule_preview()

    def _drop_files(self,paths):self._add_files(paths)

    def _dispose_canvases(self):
        for _path,canvas in self.canvases:
            try:canvas.pil_img.close()
            except Exception:pass
            try:canvas.tk_img=None; canvas.destroy()
            except Exception:pass
        self.canvases.clear()

    def _show_empty(self):
        for widget in self.cards_frame.winfo_children():widget.destroy()
        tk.Label(
            self.cards_frame,
            text="Drop images here • drag each photo inside its crop frame",
            bg=BG,fg=MUTED,font=("Segoe UI",10,"bold"),justify="center"
        ).grid(row=0,column=0,padx=30,pady=55)

    def _rebuild_cards(self):
        old_state={path:(canvas.get_crop_fractions(),canvas.total_rotation) for path,canvas in self.canvases}
        self._dispose_canvases()
        for widget in self.cards_frame.winfo_children():widget.destroy()
        if not self.files:
            self._show_empty(); self.count_lbl.config(text="0 images"); return
        for index,path in enumerate(self.files):
            card=tk.Frame(
                self.cards_frame,bg=SURFACE,width=205,height=150,
                highlightbackground=BORDER,highlightthickness=1
            )
            card.grid(row=0,column=index,sticky="ns",padx=(0,7),pady=1)
            card.grid_propagate(False)
            top=tk.Frame(card,bg=SURFACE); top.pack(fill="x",padx=6,pady=(5,2))
            name=self.display_names.get(path,os.path.basename(path))
            lbl(top,f"{index+1}. {name}",7,TEXT,True).pack(side="left",fill="x",expand=True)
            tk.Button(
                top,text="×",command=lambda p=path:self._remove(p),bg=SURFACE2,fg=MUTED,
                relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2",width=2
            ).pack(side="right")
            _ow,_oh,_x,_y,iw,ih=self._frame_values(path)
            canvas=CropCanvas(card,path,iw,ih,size=(172,105),on_change=self._schedule_preview)
            canvas.pack(padx=6,pady=(0,3)); self.canvases.append((path,canvas))
            if path in old_state:
                fractions,rotation=old_state[path]
                try:
                    canvas.set_crop_fractions(fractions)
                    if rotation:canvas.rotate_manual(rotation)
                except Exception:pass
            controls=tk.Frame(card,bg=SURFACE); controls.pack(fill="x",padx=6,pady=(0,4))
            tk.Button(
                controls,text="↶ Rotate",command=lambda c=canvas:c.rotate_manual(90),
                bg=SURFACE2,fg=TEXT,relief="flat",font=("Segoe UI",7),cursor="hand2"
            ).pack(side="left")
            lbl(controls,"Drag to position",7,MUTED).pack(side="right")
        self.cards_frame.update_idletasks()
        self.cards_canvas.config(scrollregion=self.cards_canvas.bbox("all"))
        self.count_lbl.config(text=f"{len(self.files)} image(s)")

    def _remove(self,path):
        if path in self.files:self.files.remove(path)
        self._rebuild_cards(); self._schedule_preview()

    def _paper_size(self):
        size=PAPER_SIZES_MM.get(self.sheet_var.get())
        if not size:return None
        return float(size[0]),float(size[1])

    def _layout(self):
        paper=self._paper_size()
        if not paper:return None
        ow,oh,*_=self._frame_values(self.files[0] if self.files else None)
        spacing=max(0,self._num(self.spacing_var,3)); margin=5.0
        candidates=[]
        for sw,sh in (paper,(paper[1],paper[0])):
            for rotated in (False,True):
                iw,ih=(oh,ow) if rotated else (ow,oh)
                cols=max(0,int((sw-margin*2+spacing)//(iw+spacing)))
                rows=max(0,int((sh-margin*2+spacing)//(ih+spacing)))
                candidates.append((cols*rows,not rotated,sw*sh-(cols*iw+(cols-1)*spacing)*(rows*ih+(rows-1)*spacing),sw,sh,rotated,cols,rows,iw,ih))
        if self.auto_layout_var.get():
            best=max(candidates,key=lambda x:(x[0],x[1],x[2]))
            _cap,_prefer,_waste,sw,sh,rotated,cols,rows,iw,ih=best
            self.cols_var.set(max(1,cols)); self.rows_var.set(max(1,rows))
        else:
            wanted_cols=max(1,int(self.cols_var.get())); wanted_rows=max(1,int(self.rows_var.get()))
            fitting=[c for c in candidates if c[6]>=wanted_cols and c[7]>=wanted_rows]
            best=max(fitting,key=lambda x:(x[1],x[2])) if fitting else max(candidates,key=lambda x:x[0])
            _cap,_prefer,_waste,sw,sh,rotated,_auto_c,_auto_r,iw,ih=best
            cols,rows=wanted_cols,wanted_rows
        capacity=max(1,cols*rows)
        total_w=cols*iw+max(0,cols-1)*spacing; total_h=rows*ih+max(0,rows-1)*spacing
        ox=(sw-total_w)/2; oy=(sh-total_h)/2
        fits=total_w<=sw-2 and total_h<=sh-2
        return dict(sw=sw,sh=sh,rotated=rotated,cols=cols,rows=rows,iw=iw,ih=ih,spacing=spacing,ox=ox,oy=oy,capacity=capacity,fits=fits)

    def _source_image(self,path,canvas,preview):
        if preview:return canvas.pil_img.copy()
        with Image.open(path) as opened:
            corrected=ImageOps.exif_transpose(opened); image=corrected.convert("RGB")
            if corrected is not opened:
                try:corrected.close()
                except Exception:pass
        if canvas.total_rotation:image=image.rotate(canvas.total_rotation,expand=True)
        return image

    def _render_item(self,path,canvas,dpi,preview=False):
        ow,oh,x,y,iw,ih=self._frame_values(path)
        outer=Image.new("RGB",(max(1,mm_to_px(ow,dpi)),max(1,mm_to_px(oh,dpi))),"white")
        source=self._source_image(path,canvas,preview)
        try:
            l,t,r,b=canvas.get_crop_fractions(); box=(int(l*source.width),int(t*source.height),int(r*source.width),int(b*source.height))
            cropped=source.crop(box)
            try:photo=ImageOps.fit(cropped,(max(1,mm_to_px(iw,dpi)),max(1,mm_to_px(ih,dpi))),method=Image.Resampling.BILINEAR if preview else Image.Resampling.LANCZOS,centering=(.5,.5))
            finally:cropped.close()
            try:outer.paste(photo,(mm_to_px(x,dpi),mm_to_px(y,dpi)))
            finally:photo.close()
        finally:source.close()
        return outer

    def _draw_cut_marks(self,draw,x,y,w,h,dpi):
        gap=mm_to_px(1.5,dpi); length=mm_to_px(4,dpi); lw=max(1,mm_to_px(.2,dpi))
        for cx,dx in ((x,-1),(x+w,1)):
            for cy,dy in ((y,-1),(y+h,1)):
                draw.line((cx+dx*gap,cy,cx+dx*(gap+length),cy),fill="black",width=lw)
                draw.line((cx,cy+dy*gap,cx,cy+dy*(gap+length)),fill="black",width=lw)

    def _draw_outline(self,draw,x,y,w,h,dpi,color="black",stroke_mm=.15):
        line_width=max(1,mm_to_px(stroke_mm,dpi))
        draw.rectangle((x,y,x+w-1,y+h-1),outline=color,width=line_width)

    def _build_sheet(self,start,dpi,preview=False):
        layout=self._layout()
        if not layout:return None,0,None
        sheet=Image.new("RGB",(mm_to_px(layout["sw"],dpi),mm_to_px(layout["sh"],dpi)),"white")
        draw=ImageDraw.Draw(sheet); used=0
        by=dict(self.canvases)
        for slot in range(layout["capacity"]):
            index=start+slot
            if index>=len(self.files):break
            path=self.files[index]; canvas=by.get(path)
            if canvas is None:continue
            item=self._render_item(path,canvas,dpi,preview)
            try:
                if layout["rotated"]:
                    rotated_item=item.rotate(90,expand=True)
                    item.close(); item=rotated_item
                col=slot%layout["cols"]; row=slot//layout["cols"]
                x=mm_to_px(layout["ox"]+col*(layout["iw"]+layout["spacing"]),dpi)
                y=mm_to_px(layout["oy"]+row*(layout["ih"]+layout["spacing"]),dpi)
                sheet.paste(item,(x,y))
                # Preview-only grey outlines make each white polaroid visible on
                # the white sheet. They are not exported unless Outline is chosen.
                if preview:
                    self._draw_outline(draw,x,y,item.width,item.height,dpi,color="#707080",stroke_mm=.35)
                    if self.mark_mode_var.get()=="Crop marks":
                        self._draw_cut_marks(draw,x,y,item.width,item.height,dpi)
                elif self.mark_mode_var.get()=="Crop marks":
                    self._draw_cut_marks(draw,x,y,item.width,item.height,dpi)
                else:
                    self._draw_outline(draw,x,y,item.width,item.height,dpi,color="black",stroke_mm=.15)
                used+=1
            finally:item.close()
        return sheet,used,layout

    def _schedule_preview(self,*_args):
        if self._preview_job:
            try:self.after_cancel(self._preview_job)
            except Exception:pass
        self._preview_job=self.after(180,self._draw_preview)

    def _draw_preview(self):
        self._preview_job=None; self.preview_canvas.delete("all")
        if not self.files:
            self.preview_canvas.create_text(max(1,self.preview_canvas.winfo_width())/2,max(1,self.preview_canvas.winfo_height())/2,text="Add photos to preview the first sheet",fill=MUTED,font=("Segoe UI",11,"bold")); self.preview_info.config(text=""); return
        sheet,used,layout=self._build_sheet(0,45,True)
        if sheet is None:
            self.preview_canvas.create_text(190,180,text=f"{self.sheet_var.get()} was not found in the paper-size Excel file.",fill=WARNING,width=330,font=("Segoe UI",9,"bold")); return
        try:
            cw=max(80,self.preview_canvas.winfo_width()-18); ch=max(80,self.preview_canvas.winfo_height()-18)
            scale=min(cw/sheet.width,ch/sheet.height); shown=sheet.resize((max(1,int(sheet.width*scale)),max(1,int(sheet.height*scale))),Image.Resampling.BILINEAR)
            try:
                self.preview_photo=ImageTk.PhotoImage(shown); self.preview_canvas.create_image(self.preview_canvas.winfo_width()/2,self.preview_canvas.winfo_height()/2,image=self.preview_photo,anchor="center")
            finally:shown.close()
        finally:sheet.close()
        pages=math.ceil(len(self.files)/layout["capacity"]); self.preview_info.config(text=f"{layout['cols']} × {layout['rows']} • {pages} sheet(s)")
        self.status.config(text=(f"{used} per sheet • {'rotated automatically' if layout['rotated'] else 'upright'}" if layout["fits"] else "Manual layout does not fit the selected sheet."),fg=SUCCESS if layout["fits"] else WARNING)

    def _export_pdf(self):
        if not self.files:messagebox.showwarning("Nothing to export","Add at least one photo first.",parent=self); return
        layout=self._layout()
        if not layout:messagebox.showerror("Paper size missing",f"{self.sheet_var.get()} is missing from the paper-size Excel file.",parent=self); return
        if not layout["fits"]:messagebox.showwarning("Layout does not fit","Reduce rows, columns, spacing, or frame size.",parent=self); return
        out=filedialog.asksaveasfilename(parent=self,title="Export polaroid sheets",defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile=f"polaroids_{self.sheet_var.get()}.pdf")
        if not out:return
        self.export_btn.config(state="disabled"); self.status.config(text="Exporting print-ready PDF…",fg=MUTED); self.update_idletasks()
        pdf=fitz.open(); start=0
        try:
            while start<len(self.files):
                sheet,used,current=self._build_sheet(start,300,False)
                if sheet is None or used<=0:break
                buffer=io.BytesIO()
                try:
                    sheet.save(buffer,"PNG",compress_level=2)
                    page=pdf.new_page(width=current["sw"]*72/25.4,height=current["sh"]*72/25.4)
                    page.insert_image(page.rect,stream=buffer.getvalue())
                finally:buffer.close(); sheet.close()
                start+=used; gc.collect(0)
            pdf.save(out,garbage=3,deflate=True)
            self.status.config(text=f"Exported {math.ceil(len(self.files)/layout['capacity'])} sheet(s).",fg=SUCCESS)
            messagebox.showinfo("Done",f"Polaroid PDF exported to:\n{out}",parent=self)
        except Exception as ex:messagebox.showerror("Export failed",str(ex),parent=self)
        finally:pdf.close(); self.export_btn.config(state="normal")

    def _clear(self):
        self._dispose_canvases(); self.files.clear(); self.display_names.clear(); self.preview_photo=None
        for path in self._generated:
            try:os.remove(path)
            except Exception:pass
        self._generated.clear(); self._show_empty(); self.count_lbl.config(text="0 images"); self.preview_canvas.delete("all"); self.status.config(text="Add images or a PDF.",fg=MUTED); gc.collect()

# ── Scan Cropping Tab ────────────────────────────────────────────────────────
class ScanCroppingTab(tk.Frame,DropMixin):
    """Detect separate photographs on white scanned PDF pages."""
    def __init__(self,parent):
        super().__init__(parent,bg=BG)
        self.pages=[]; self.current_page=-1; self.selected_crop=-1
        self.canvas_photo=None; self._drag_corner=None; self._busy=False
        self._worker_results=queue.Queue(); self._scan_generation=0
        self._poll_job=None
        self._build(); self._poll_job=self.after(80,self._poll_worker_results)

    def _build(self):
        self.grid_rowconfigure(0,weight=1); self.grid_columnconfigure(1,weight=1)
        left=tk.Frame(self,bg=SURFACE,width=270); left.grid(row=0,column=0,sticky="nsew"); left.grid_propagate(False)
        lbl(left,"SCAN CROPPING",9,MUTED,True).pack(anchor="w",padx=14,pady=(12,4))
        tk.Label(left,text="Detect photos separated by gaps on a white scanner background.",bg=SURFACE,fg=MUTED,font=("Segoe UI",8),wraplength=235,justify="left").pack(anchor="w",padx=14,pady=(0,8))
        row=tk.Frame(left,bg=SURFACE); row.pack(fill="x",padx=14)
        styled_btn(row,"Add scanned PDFs",self._add_files).pack(side="left",fill="x",expand=True)
        styled_btn(row,"Clear",self._clear,style="secondary").pack(side="left",padx=(6,0))
        sep(left)
        self.sensitivity_var=tk.IntVar(value=22)
        self.sensitivity_lbl=lbl(left,"Detection sensitivity: 22",8); self.sensitivity_lbl.pack(anchor="w",padx=14)
        ttk.Scale(left,from_=5,to=70,variable=self.sensitivity_var,orient="horizontal",command=lambda v:self.sensitivity_lbl.config(text=f"Detection sensitivity: {int(float(v))}")).pack(fill="x",padx=14)
        minrow=tk.Frame(left,bg=SURFACE); minrow.pack(fill="x",padx=14,pady=(7,2))
        lbl(minrow,"Minimum photo area (%)",8).pack(side="left")
        self.min_area_var=tk.DoubleVar(value=3.0); entry(minrow,self.min_area_var,6).pack(side="right")
        self.straighten_var=tk.BooleanVar(value=True)
        tk.Checkbutton(left,text="Straighten detected photos",variable=self.straighten_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",9)).pack(anchor="w",padx=14,pady=4)
        styled_btn(left,"Detect current page",self._detect_current,style="success").pack(fill="x",padx=14,pady=(5,2))
        controls=tk.Frame(left,bg=SURFACE); controls.pack(fill="x",padx=14,pady=2)
        styled_btn(controls,"Add manual crop",self._add_manual_crop,style="secondary").pack(side="left",fill="x",expand=True)
        styled_btn(controls,"Delete crop",self._delete_crop,style="secondary").pack(side="left",fill="x",expand=True,padx=(5,0))
        sep(left); lbl(left,"PAGES",8,MUTED,True).pack(anchor="w",padx=14)
        self.page_list=tk.Listbox(left,bg=SURFACE2,fg=TEXT,selectbackground=ACCENT,selectforeground="white",relief="flat",font=("Segoe UI",8),highlightthickness=0)
        self.page_list.pack(fill="both",expand=True,padx=14,pady=(4,6)); self.page_list.bind("<<ListboxSelect>>",self._page_selected)
        self.status=lbl(left,"Add one or more scanned PDF files.",8,MUTED); self.status.pack(fill="x",padx=14,pady=(0,5))
        self.export_btn=styled_btn(left,"Export cropped photos to one PDF",self._export_pdf,style="success"); self.export_btn.pack(fill="x",padx=14,pady=(0,12))

        center=tk.Frame(self,bg=BG); center.grid(row=0,column=1,sticky="nsew",padx=8,pady=8); center.grid_rowconfigure(1,weight=1); center.grid_columnconfigure(0,weight=1)
        head=tk.Frame(center,bg=BG); head.grid(row=0,column=0,sticky="ew",pady=(0,5))
        lbl(head,"SCAN PREVIEW",9,MUTED,True).pack(side="left")
        self.crop_count_lbl=lbl(head,"0 detected",8,MUTED); self.crop_count_lbl.pack(side="right")
        self.canvas=tk.Canvas(center,bg=SURFACE,highlightbackground=BORDER,highlightthickness=1,cursor="crosshair")
        self.canvas.grid(row=1,column=0,sticky="nsew")
        self.canvas.bind("<Configure>",lambda e:self._draw_page())
        self.canvas.bind("<ButtonPress-1>",self._canvas_press)
        self.canvas.bind("<B1-Motion>",self._canvas_drag)
        self.canvas.bind("<ButtonRelease-1>",lambda e:setattr(self,"_drag_corner",None))
        self.setup_drop_everywhere(self,self._drop_files,{".pdf"})

    def _poll_worker_results(self):
        try:
            while True:
                kind,payload=self._worker_results.get_nowait()
                if kind=="detect_done":
                    generation,results=payload
                    if generation==self._scan_generation:self._apply_detection(results)
                elif kind=="detect_error":
                    generation,error=payload
                    if generation==self._scan_generation:self._detection_error(error)
                elif kind=="export_done":
                    path,count=payload; self._export_done(path,count)
                elif kind=="export_error":
                    self._detection_error(payload)
        except queue.Empty:
            pass
        try:self._poll_job=self.after(80,self._poll_worker_results)
        except Exception:self._poll_job=None

    def _add_files(self):
        self._drop_files(list(filedialog.askopenfilenames(parent=self,title="Add scanned PDFs",filetypes=[("PDF files","*.pdf")])))

    def _drop_files(self,paths):
        self._scan_generation+=1
        added=0
        for path in paths:
            if os.path.splitext(path)[1].lower()!=".pdf":continue
            try:
                doc=fitz.open(path)
                try:
                    for page_no in range(len(doc)):
                        if any(p["source"]==path and p["page_no"]==page_no for p in self.pages):continue
                        pix=doc[page_no].get_pixmap(matrix=fitz.Matrix(100/72,100/72),alpha=False)
                        preview=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                        preview.thumbnail((1500,1500),Image.Resampling.BILINEAR)
                        self.pages.append({"source":path,"page_no":page_no,"preview":preview,"crops":[]}); added+=1
                finally:doc.close()
            except Exception as ex:messagebox.showerror("PDF import failed",f"{os.path.basename(path)}\n{ex}",parent=self)
        if added:
            self._refresh_page_list(); self._select_page(0 if self.current_page<0 else self.current_page)
            self.status.config(text=f"Loaded {len(self.pages)} page(s). Detecting photos…",fg=MUTED); self._detect_all_async()

    def _refresh_page_list(self):
        self.page_list.delete(0,"end")
        for i,page in enumerate(self.pages):
            self.page_list.insert("end",f"{i+1}. {os.path.basename(page['source'])} — page {page['page_no']+1} ({len(page['crops'])})")

    def _select_page(self,index):
        if not self.pages:return
        self.current_page=max(0,min(index,len(self.pages)-1)); self.selected_crop=0 if self.pages[self.current_page]["crops"] else -1
        self.page_list.selection_clear(0,"end"); self.page_list.selection_set(self.current_page); self.page_list.see(self.current_page); self._draw_page()

    def _page_selected(self,_event=None):
        selected=self.page_list.curselection()
        if selected:self._select_page(selected[0])

    @staticmethod
    def _order_points(points):
        pts=np.asarray(points,dtype=np.float32); sums=pts.sum(axis=1); diffs=np.diff(pts,axis=1).ravel()
        return np.array([pts[np.argmin(sums)],pts[np.argmin(diffs)],pts[np.argmax(sums)],pts[np.argmax(diffs)]],dtype=np.float32)

    def _detect_image(self,image,sensitivity=None,min_area_percent=None):
        rgb=np.array(image.convert("RGB")); gray=cv2.cvtColor(rgb,cv2.COLOR_RGB2GRAY)
        threshold=max(5,min(70,int(self.sensitivity_var.get() if sensitivity is None else sensitivity)))
        mask=np.where(gray<255-threshold,255,0).astype(np.uint8)
        size=max(3,int(min(mask.shape[:2])*.006)); size+=1-size%2
        kernel=cv2.getStructuringElement(cv2.MORPH_RECT,(size,size))
        mask=cv2.morphologyEx(mask,cv2.MORPH_CLOSE,kernel,iterations=2)
        contours,_=cv2.findContours(mask,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
        area_percent=float(self.min_area_var.get() if min_area_percent is None else min_area_percent)
        min_area=max(.001,area_percent/100)*image.width*image.height
        crops=[]
        for contour in contours:
            if cv2.contourArea(contour)<min_area:continue
            rect=cv2.minAreaRect(contour); (_cx,_cy),(w,h),_angle=rect
            if min(w,h)<35:continue
            box=self._order_points(cv2.boxPoints(rect)); centre=box.mean(axis=0); box=centre+(box-centre)*0.992
            box[:,0]=np.clip(box[:,0],0,image.width-1); box[:,1]=np.clip(box[:,1],0,image.height-1)
            crops.append([(float(x/image.width),float(y/image.height)) for x,y in box])
        crops.sort(key=lambda points:(sum(y for _x,y in points)/4,sum(x for x,_y in points)/4))
        return crops

    def _detect_all_async(self):
        if self._busy:return
        self._busy=True; self.export_btn.config(state="disabled")
        previews=[p["preview"].copy() for p in self.pages]
        sensitivity=int(self.sensitivity_var.get()); min_area=float(self.min_area_var.get())
        generation=self._scan_generation
        def work():
            results=[]
            try:
                for image in previews:
                    try:results.append(self._detect_image(image,sensitivity,min_area))
                    finally:image.close()
                self._worker_results.put(("detect_done",(generation,results)))
            except Exception as ex:
                for image in previews:
                    try:image.close()
                    except Exception:pass
                self._worker_results.put(("detect_error",(generation,str(ex))))
        threading.Thread(target=work,daemon=True).start()

    def _apply_detection(self,results):
        for page,crops in zip(self.pages,results):page["crops"]=crops
        self._busy=False; self.export_btn.config(state="normal"); self._refresh_page_list(); self._select_page(max(0,self.current_page))
        total=sum(len(p["crops"]) for p in self.pages); self.status.config(text=f"Detected {total} photo(s). Drag blue corner handles to adjust.",fg=SUCCESS)

    def _detection_error(self,error):
        self._busy=False; self.export_btn.config(state="normal"); messagebox.showerror("Detection failed",str(error),parent=self)

    def _detect_current(self):
        if self.current_page<0:return
        page=self.pages[self.current_page]
        try:page["crops"]=self._detect_image(page["preview"]); self.selected_crop=0 if page["crops"] else -1; self._refresh_page_list(); self._draw_page()
        except Exception as ex:messagebox.showerror("Detection failed",str(ex),parent=self)

    def _geometry(self):
        if self.current_page<0 or not self.pages:return None
        image=self.pages[self.current_page]["preview"]; cw=max(1,self.canvas.winfo_width()); ch=max(1,self.canvas.winfo_height())
        scale=min((cw-24)/image.width,(ch-24)/image.height); dw=max(1,int(image.width*scale)); dh=max(1,int(image.height*scale)); ox=(cw-dw)//2; oy=(ch-dh)//2
        return image,ox,oy,dw,dh

    def _draw_page(self):
        self.canvas.delete("all"); geometry=self._geometry()
        if not geometry:
            self.canvas.create_text(max(1,self.canvas.winfo_width())/2,max(1,self.canvas.winfo_height())/2,text="Drop scanned PDF files here",fill=MUTED,font=("Segoe UI",13,"bold")); return
        image,ox,oy,dw,dh=geometry; resized=image.resize((dw,dh),Image.Resampling.BILINEAR)
        old_photo=self.canvas_photo
        try:self.canvas_photo=ImageTk.PhotoImage(resized)
        finally:resized.close()
        if old_photo is not None:
            try:old_photo.__del__()
            except Exception:pass
        self.canvas.create_image(ox,oy,anchor="nw",image=self.canvas_photo)
        crops=self.pages[self.current_page]["crops"]
        for index,points in enumerate(crops):
            flat=[]
            for x,y in points:flat.extend((ox+x*dw,oy+y*dh))
            colour=ACCENT if index==self.selected_crop else SUCCESS
            self.canvas.create_polygon(*flat,outline=colour,fill="",width=3 if index==self.selected_crop else 2)
            self.canvas.create_text(sum(flat[0::2])/4,sum(flat[1::2])/4,text=str(index+1),fill="white",font=("Segoe UI",10,"bold"))
            if index==self.selected_crop:
                for corner in range(4):
                    x,y=flat[corner*2],flat[corner*2+1]; self.canvas.create_oval(x-7,y-7,x+7,y+7,fill=ACCENT,outline="white")
        self.crop_count_lbl.config(text=f"{len(crops)} detected")

    def _canvas_press(self,event):
        geometry=self._geometry()
        if not geometry:return
        _image,ox,oy,dw,dh=geometry; page=self.pages[self.current_page]; nearest=None
        for index,points in enumerate(page["crops"]):
            for corner,(nx,ny) in enumerate(points):
                dist=(event.x-(ox+nx*dw))**2+(event.y-(oy+ny*dh))**2
                if dist<=18**2 and (nearest is None or dist<nearest[0]):nearest=(dist,index,corner)
        if nearest:
            self.selected_crop=nearest[1]; self._drag_corner=nearest[2]; self._draw_page(); return
        for index,points in enumerate(page["crops"]):
            polygon=np.array([(ox+x*dw,oy+y*dh) for x,y in points],np.float32)
            if cv2.pointPolygonTest(polygon,(event.x,event.y),False)>=0:self.selected_crop=index; self._draw_page(); return

    def _canvas_drag(self,event):
        if self._drag_corner is None or self.selected_crop<0:return
        geometry=self._geometry()
        if not geometry:return
        _image,ox,oy,dw,dh=geometry
        nx=max(0,min(1,(event.x-ox)/max(1,dw))); ny=max(0,min(1,(event.y-oy)/max(1,dh)))
        self.pages[self.current_page]["crops"][self.selected_crop][self._drag_corner]=(nx,ny); self._draw_page()

    def _add_manual_crop(self):
        if self.current_page<0:return
        self.pages[self.current_page]["crops"].append([(0.15,0.15),(0.85,0.15),(0.85,0.85),(0.15,0.85)])
        self.selected_crop=len(self.pages[self.current_page]["crops"])-1; self._refresh_page_list(); self._draw_page()

    def _delete_crop(self):
        if self.current_page<0 or self.selected_crop<0:return
        crops=self.pages[self.current_page]["crops"]
        if self.selected_crop<len(crops):crops.pop(self.selected_crop)
        self.selected_crop=min(self.selected_crop,len(crops)-1); self._refresh_page_list(); self._draw_page()

    @staticmethod
    def _native_dpi(page):
        dpi=300
        try:
            for info in page.get_image_info(xrefs=True):
                bbox=fitz.Rect(info.get("bbox")); coverage=(bbox.width*bbox.height)/(page.rect.width*page.rect.height)
                if coverage>.85:
                    dpi=max(dpi,int(max(info.get("width",0)/(page.rect.width/72),info.get("height",0)/(page.rect.height/72))))
        except Exception:pass
        return max(150,min(1200,dpi))

    def _crop_photo(self,image,points,straighten):
        array=image if isinstance(image,np.ndarray) else np.array(image.convert("RGB"))
        h,w=array.shape[:2]; pts=self._order_points([(x*w,y*h) for x,y in points])
        if not straighten:
            x1=max(0,int(np.floor(pts[:,0].min()))); y1=max(0,int(np.floor(pts[:,1].min()))); x2=min(w,int(np.ceil(pts[:,0].max()))); y2=min(h,int(np.ceil(pts[:,1].max())))
            return Image.fromarray(array[y1:y2,x1:x2].copy())
        tl,tr,br,bl=pts; out_w=max(1,int(round(max(np.linalg.norm(br-bl),np.linalg.norm(tr-tl))))); out_h=max(1,int(round(max(np.linalg.norm(tr-br),np.linalg.norm(tl-bl)))))
        target=np.array([[0,0],[out_w-1,0],[out_w-1,out_h-1],[0,out_h-1]],np.float32)
        matrix=cv2.getPerspectiveTransform(pts,target)
        return Image.fromarray(cv2.warpPerspective(array,matrix,(out_w,out_h),flags=cv2.INTER_LANCZOS4,borderMode=cv2.BORDER_REPLICATE))

    def _export_pdf(self):
        total=sum(len(p["crops"]) for p in self.pages)
        if not total:messagebox.showwarning("Nothing to export","Detect or add at least one crop first.",parent=self); return
        out=filedialog.asksaveasfilename(parent=self,title="Export cropped photos",defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile="scanned_photos.pdf")
        if not out:return
        snapshot=[(p["source"],p["page_no"],[list(c) for c in p["crops"]]) for p in self.pages]; straighten=self.straighten_var.get()
        self.export_btn.config(state="disabled"); self.status.config(text="Exporting full-resolution photos…",fg=MUTED)
        def work():
            pdf=fitz.open(); count=0
            try:
                for source,page_no,crops in snapshot:
                    doc=fitz.open(source)
                    try:
                        page=doc[page_no]; dpi=self._native_dpi(page); pix=page.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),alpha=False); full=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                    finally:doc.close()
                    full_array=np.array(full); full.close()
                    try:
                        for points in crops:
                            photo=self._crop_photo(full_array,points,straighten); buffer=io.BytesIO()
                            try:
                                photo.save(buffer,"PNG",compress_level=3); output_page=pdf.new_page(width=photo.width*72/dpi,height=photo.height*72/dpi); output_page.insert_image(output_page.rect,stream=buffer.getvalue()); count+=1
                            finally:buffer.close(); photo.close()
                    finally:
                        del full_array
                    gc.collect(0)
                pdf.save(out,garbage=3,deflate=True); self._worker_results.put(("export_done",(out,count)))
            except Exception as ex:self._worker_results.put(("export_error",str(ex)))
            finally:pdf.close()
        threading.Thread(target=work,daemon=True).start()

    def _export_done(self,path,count):
        self.export_btn.config(state="normal"); self.status.config(text=f"Exported {count} photo(s).",fg=SUCCESS); messagebox.showinfo("Done",f"Exported {count} cropped photo(s) to:\n{path}",parent=self)

    def _clear(self):
        self._scan_generation+=1; self._busy=False
        try:self.export_btn.config(state="normal")
        except Exception:pass
        for page in self.pages:
            try:page["preview"].close()
            except Exception:pass
        self.pages.clear(); self.current_page=-1; self.selected_crop=-1; self.canvas_photo=None; self.page_list.delete(0,"end"); self.crop_count_lbl.config(text="0 detected"); self.status.config(text="Add one or more scanned PDF files.",fg=MUTED); self._draw_page(); gc.collect()

    def clear_memory_cache(self):
        self.canvas_photo=None; self._draw_page(); gc.collect()

    def destroy(self):
        try:
            refresh=getattr(self,"_copypro_drop_refresh_job",None)
            if refresh:self.after_cancel(refresh)
            if self._poll_job:self.after_cancel(self._poll_job)
        except Exception:pass
        self._poll_job=None; self._clear(); super().destroy()


# ── Print Counter Tab ────────────────────────────────────────────────────────
#
# PLU code lookup table.
# Structure: input_plu → { (sides, paper_type): [(max_qty, output_code), ...] }
# Ranges are inclusive upper bounds; last entry has None = no upper limit.
#
# Paper types:  "photo"  = photo paper  (7xx codes)
#               "plain"  = plain A4     (5xx codes)
#
# Sides:        "one" / "double"
# Size:         "A4" / "A3"
#
# The table maps known PLU groups to their tiered output codes.
# Input PLUs that are NOT in the table are passed through unchanged.

def _build_plu_table():
    """
    Returns dict:  input_plu_str → list of (max_qty_or_None, output_code_str)
    sorted by max_qty ascending (None = infinity, always last).
    """
    table = {}

    def add(base_in, base_out, tiers):
        # tiers = [(max_qty, suffix), ...]  e.g. [(10,'0'),(50,'1'),...]
        for max_q, suffix in tiers:
            code_in  = str(base_in)
            code_out = str(base_out) + suffix
            if code_in not in table:
                table[code_in] = []
            table[code_in].append((max_q, code_out))

    TIERS = [(10, "0"), (50, "1"), (100, "2"), (None, "3")]

    # Photo paper  (7xx)
    add(710, 71, TIERS)   # one-sided colour A4  photo
    add(720, 72, TIERS)   # double-sided colour A4 photo
    add(730, 73, TIERS)   # one-sided colour A3  photo
    add(740, 74, TIERS)   # double-sided colour A3 photo

    # Plain paper  (5xx) — same structure
    add(510, 51, TIERS)
    add(520, 52, TIERS)
    add(530, 53, TIERS)
    add(540, 54, TIERS)

    return table

PLU_TABLE = _build_plu_table()

# Product codes, names, prices, search aliases, wide-format rows and paper sizes
# are loaded exclusively from the selected Excel workbooks.
#
# No catalogue or pricing fallback is embedded in the Python application.
PLU_PRICES = {}
PLU_DESCRIPTIONS = {}
PLU_SEARCH_ALIASES = {}
WIDE_FORMAT_ITEMS = {}
WIDE_FORMAT_ORDER = []

def _editable_paths():
    settings=load_app_settings()
    return {
        "codes_file": settings.get("codes_file", DEFAULT_CODES_FILE),
        "paper_sizes_file": settings.get("paper_sizes_file", DEFAULT_PAPER_SIZES_FILE),
    }

def _copy_default_if_missing(destination, bundled_name):
    if os.path.exists(destination): return
    os.makedirs(os.path.dirname(os.path.abspath(destination)),exist_ok=True)
    bundled=resource_path(bundled_name)
    if not os.path.exists(bundled): bundled=resource_path(os.path.join("data",bundled_name))
    if os.path.exists(bundled): shutil.copy2(bundled,destination)

def _normalise_excel_value(value):
    return "" if value is None else value

def _xlsx_rows(path, sheet_name):
    wb=load_workbook(path,read_only=True,data_only=True)
    try:
        ws=wb[sheet_name] if sheet_name in wb.sheetnames else wb[wb.sheetnames[0]]
        iterator=ws.iter_rows(values_only=True)
        headers=[str(x).strip() if x is not None else "" for x in next(iterator)]
        result=[]
        for values in iterator:
            row={headers[i]:_normalise_excel_value(values[i]) if i<len(values) else "" for i in range(len(headers)) if headers[i]}
            if any(str(v).strip() for v in row.values()): result.append(row)
        return result
    finally: wb.close()

def _csv_rows(path):
    with open(path,"r",encoding="utf-8-sig",newline="") as f:
        sample=f.read(4096); f.seek(0)
        try: dialect=csv.Sniffer().sniff(sample,delimiters=",;\t")
        except Exception: dialect=csv.excel
        return list(csv.DictReader(f,dialect=dialect))

def _data_rows(path, sheet_name):
    return _xlsx_rows(path,sheet_name) if os.path.splitext(path)[1].lower() in (".xlsx",".xlsm") else _csv_rows(path)

def _write_sizes_workbook(path, rows):
    # Preserve the Instructions sheet and workbook formatting when possible.
    if os.path.exists(path) and os.path.splitext(path)[1].lower() in (".xlsx",".xlsm"):
        wb=load_workbook(path)
        ws=wb["Sizes"] if "Sizes" in wb.sheetnames else wb[wb.sheetnames[0]]
        if ws.max_row>1: ws.delete_rows(2,ws.max_row-1)
    else:
        wb=Workbook(); ws=wb.active; ws.title="Sizes"
        ws.append(["group","name","width_mm","height_mm","active","is_custom","notes_lt"])
    fields=["group","name","width_mm","height_mm","active","is_custom","notes_lt"]
    for row in rows: ws.append([row.get(field,"") for field in fields])
    wb.save(path)

def _bundled_codes_file():
    direct = resource_path("copypro_kodai.xlsx")
    if os.path.isfile(direct):
        return direct
    nested = resource_path(os.path.join("data", "copypro_kodai.xlsx"))
    return nested if os.path.isfile(nested) else None


def _safe_version_for_filename(value):
    return "".join(
        ch if ch.isalnum() or ch in ("-", "_", ".") else "_"
        for ch in str(value or "unknown")
    )


def sync_code_workbook_after_app_update():
    """
    Replace only the active item-code workbook once per application version.

    The previous workbook is backed up to:
        %APPDATA%\\CopyPro\\backups\\

    The paper-size workbook is intentionally never replaced here.
    """
    bundled = _bundled_codes_file()
    if not bundled:
        return {
            "updated": False,
            "reason": "No bundled code workbook was found.",
        }

    settings = load_app_settings()
    installed_version = str(current_version()).strip() or "0.0.0"
    previous_data_version = str(
        settings.get("codes_workbook_app_version", "")
    ).strip()

    # Run exactly once for each installed application version.
    if previous_data_version == installed_version:
        return {
            "updated": False,
            "reason": "Code workbook already synced for this version.",
        }

    active_path = os.path.abspath(
        settings.get("codes_file", DEFAULT_CODES_FILE)
    )
    os.makedirs(os.path.dirname(active_path), exist_ok=True)

    backup_path = None
    if os.path.isfile(active_path):
        timestamp = __import__("datetime").datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S"
        )
        old_version = _safe_version_for_filename(
            previous_data_version or "before_first_sync"
        )
        extension = os.path.splitext(active_path)[1] or ".xlsx"
        backup_name = (
            f"copypro_kodai_{old_version}_{timestamp}{extension}"
        )
        backup_path = os.path.join(BACKUPS_DIR, backup_name)
        shutil.copy2(active_path, backup_path)

    temporary_path = active_path + ".updating"
    try:
        shutil.copy2(bundled, temporary_path)
        os.replace(temporary_path, active_path)
    except Exception:
        try:
            if os.path.exists(temporary_path):
                os.remove(temporary_path)
        except Exception:
            pass
        raise

    settings["codes_file"] = active_path
    settings["codes_workbook_app_version"] = installed_version
    save_app_settings(settings)

    return {
        "updated": True,
        "version": installed_version,
        "codes_file": active_path,
        "backup_file": backup_path,
    }


def _ensure_editable_files():
    paths=_editable_paths()
    _copy_default_if_missing(paths["codes_file"],"copypro_kodai.xlsx")
    _copy_default_if_missing(paths["paper_sizes_file"],"copypro_popieriaus_dydziai.xlsx")

def reload_editable_data():
    """Reload all product and paper data exclusively from the active Excel files."""
    global PAPER_SIZES_MM, WIDE_FORMAT_ITEMS, WIDE_FORMAT_ORDER, PLU_SEARCH_ALIASES

    _ensure_editable_files()
    paths = _editable_paths()

    codes_path = os.path.abspath(paths["codes_file"])
    sizes_path = os.path.abspath(paths["paper_sizes_file"])

    if not os.path.isfile(codes_path):
        raise FileNotFoundError(f"Code list Excel file was not found:\n{codes_path}")
    if not os.path.isfile(sizes_path):
        raise FileNotFoundError(f"Paper-size Excel file was not found:\n{sizes_path}")

    code_rows = _data_rows(codes_path, "Items")
    size_rows = _data_rows(sizes_path, "Sizes")

    required_code_columns = {
        "code", "name_lt", "price_eur", "active",
        "search_aliases_lt", "wide_format_coverage", "wide_format_label_lt",
    }
    required_size_columns = {
        "name", "width_mm", "height_mm", "active", "is_custom",
    }

    if code_rows:
        missing = required_code_columns.difference(code_rows[0].keys())
        if missing:
            raise ValueError(
                "The Items sheet is missing required columns: "
                + ", ".join(sorted(missing))
            )
    else:
        raise ValueError(f"The Items sheet contains no data:\n{codes_path}")

    if size_rows:
        missing = required_size_columns.difference(size_rows[0].keys())
        if missing:
            raise ValueError(
                "The Sizes sheet is missing required columns: "
                + ", ".join(sorted(missing))
            )
    else:
        raise ValueError(f"The Sizes sheet contains no data:\n{sizes_path}")

    new_descriptions = {}
    new_prices = {}
    new_aliases = {}
    new_wide_items = {}
    new_wide_order = []

    coverage_to_internal = {
        "dalinis": "partial",
        "pilnas": "full",
        "brėžinys": "drawing",
        "brezinys": "drawing",
        "fiksuotas": "fixed",
        "partial": "partial",
        "full": "full",
        "drawing": "drawing",
        "fixed": "fixed",
    }

    for item in code_rows:
        if str(item.get("active", "taip")).strip().lower() not in (
            "taip", "yes", "1", "true"
        ):
            continue

        code = str(item.get("code", "")).strip()
        name = str(item.get("name_lt", "")).strip()
        if not code or not name:
            continue

        new_descriptions[code] = name
        new_aliases[code] = str(item.get("search_aliases_lt", "")).strip()

        raw_price = str(item.get("price_eur", "")).strip().replace(",", ".")
        if raw_price:
            try:
                new_prices[code] = float(raw_price)
            except ValueError:
                raise ValueError(
                    f"Invalid price for code {code}: {item.get('price_eur')}"
                )

        raw_coverage = str(
            item.get("wide_format_coverage", "")
        ).strip().lower()
        coverage = coverage_to_internal.get(raw_coverage, raw_coverage)
        label = str(item.get("wide_format_label_lt", "")).strip()

        # Wide-format order is exactly the top-to-bottom row order in Excel.
        if coverage:
            new_wide_items[code] = (
                name,
                new_prices.get(code, 0.0),
                coverage,
                label or name,
            )
            new_wide_order.append(code)

    new_paper_sizes = {}
    for item in size_rows:
        if str(item.get("active", "taip")).strip().lower() not in (
            "taip", "yes", "1", "true"
        ):
            continue
        if str(item.get("is_custom", "ne")).strip().lower() in (
            "taip", "yes", "1", "true"
        ):
            continue

        name = str(item.get("name", "")).strip()
        if not name:
            continue

        try:
            width = float(str(item.get("width_mm", "")).replace(",", "."))
            height = float(str(item.get("height_mm", "")).replace(",", "."))
        except ValueError:
            raise ValueError(f"Invalid paper size dimensions for {name}")

        new_paper_sizes[name] = (width, height)

    # Replace live data only after both workbooks have loaded successfully.
    PLU_DESCRIPTIONS.clear()
    PLU_DESCRIPTIONS.update(new_descriptions)

    PLU_PRICES.clear()
    PLU_PRICES.update(new_prices)

    PLU_SEARCH_ALIASES = new_aliases
    WIDE_FORMAT_ITEMS = new_wide_items
    WIDE_FORMAT_ORDER = new_wide_order
    PAPER_SIZES_MM = new_paper_sizes

    return {
        "codes_file": codes_path,
        "paper_sizes_file": sizes_path,
        "code_count": len(PLU_DESCRIPTIONS),
        "wide_format_count": len(WIDE_FORMAT_ORDER),
        "paper_size_count": len(PAPER_SIZES_MM),
    }

def load_custom_sizes():
    try:
        result={}
        for item in _data_rows(_editable_paths()["paper_sizes_file"],"Sizes"):
            if str(item.get("is_custom","ne")).strip().lower() not in ("taip","yes","1","true"): continue
            result[str(item["name"]).strip()]=(float(str(item["width_mm"]).replace(",",".")),float(str(item["height_mm"]).replace(",",".")))
        return result
    except Exception:return {}

def save_custom_sizes(sizes_dict):
    path=_editable_paths()["paper_sizes_file"]
    try: rows=_data_rows(path,"Sizes")
    except Exception: rows=[]
    rows=[x for x in rows if str(x.get("is_custom","ne")).strip().lower() not in ("taip","yes","1","true")]
    for name,(w,h) in sizes_dict.items(): rows.append({"group":"Pasirinktiniai","name":name,"width_mm":w,"height_mm":h,"active":"taip","is_custom":"taip","notes_lt":""})
    if os.path.splitext(path)[1].lower() in (".xlsx",".xlsm"): _write_sizes_workbook(path,rows)
    else:
        fields=["group","name","width_mm","height_mm","active","is_custom","notes_lt"]
        with open(path,"w",encoding="utf-8-sig",newline="") as f:
            writer=csv.DictWriter(f,fieldnames=fields);writer.writeheader();writer.writerows(rows)

try:
    CODE_SYNC_RESULT = sync_code_workbook_after_app_update()
    reload_editable_data()
except Exception as _data_load_error:
    # Store the error so the UI can show it after Tk has started.
    CODE_SYNC_RESULT = {
        "updated": False,
        "reason": str(_data_load_error),
    }
    DATA_LOAD_ERROR = str(_data_load_error)
else:
    DATA_LOAD_ERROR = None

def get_plu_description(code):
    return PLU_DESCRIPTIONS.get(str(code).strip(), "")

def get_plu_price(code):
    return PLU_PRICES.get(str(code).strip())

def resolve_output_code(plu_str, total):
    """
    Given a raw PLU string and a total quantity, return the correct output code.
    If the PLU is not in the table, return it unchanged.
    """
    key = plu_str.strip()
    if key not in PLU_TABLE:
        return key
    for max_q, out_code in PLU_TABLE[key]:
        if max_q is None or total <= max_q:
            return out_code
    return PLU_TABLE[key][-1][1]   # fallback to last tier

ROW_BG = "#1C1C26"   # single consistent row background (dark gray)

class PrintCounterTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.rows = []
        self.active_row = None
        self.search_codes = []
        self._build()

    def on_show(self):
        try:
            reload_editable_data()
            self._update_search_results()
        except Exception as ex:
            messagebox.showerror(
                "Excel data error",
                f"Could not reload the active Excel files.\n\n{ex}",
                parent=self,
            )

    def _build(self):
        toolbar = tk.Frame(self, bg=SURFACE, pady=8)
        toolbar.pack(fill="x")
        styled_btn(toolbar, "➕  Add Row", self._add_row).pack(side="left", padx=12)
        styled_btn(toolbar, "🗑  Clear All", self._clear_all, style="secondary").pack(side="left", padx=4)
        styled_btn(toolbar, "📋  Copy Summary", self._copy_summary, style="secondary").pack(side="left", padx=4)
        styled_btn(toolbar, "↔  WIDE FORMAT PRICING", self._open_wide_format_quote, style="success").pack(side="right", padx=12)
        tk.Label(toolbar,
                 text="Enter amount*code, for example 8*710 — Enter adds the next row",
                 bg=SURFACE, fg=MUTED, font=("Segoe UI", 8)).pack(side="right", padx=8)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=BG, width=480)
        left.pack(side="left", fill="both", expand=False)
        left.pack_propagate(False)

        search_box = tk.Frame(left, bg=SURFACE, padx=12, pady=12, highlightbackground=ACCENT, highlightthickness=2)
        search_box.pack(fill="x")
        tk.Label(search_box, text="🔎  ITEM SEARCH — type a name or PLU code", bg=SURFACE, fg=MUTED,
                 font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", pady=(0, 4))
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_box, textvariable=self.search_var,
            bg=SURFACE2, fg=TEXT, insertbackground=TEXT, relief="flat",
            font=("Segoe UI", 12, "bold"))
        self.search_entry.pack(fill="x", ipady=7)
        self.search_entry.insert(0, "")
        self.search_var.trace_add("write", lambda *_: self._update_search_results())
        self.search_entry.bind("<Down>", lambda e: self._focus_search_results())
        self.search_entry.bind("<Return>", lambda e: self._use_first_search_result())

        self.search_list = tk.Listbox(search_box, height=6, bg=SURFACE2, fg=TEXT,
            selectbackground=ACCENT, selectforeground=TEXT, relief="flat", bd=0,
            highlightthickness=0, activestyle="none", font=("Segoe UI", 9))
        self.search_list.bind("<ButtonRelease-1>", lambda e: self._use_selected_search_result())
        self.search_list.bind("<Return>", lambda e: self._use_selected_search_result())
        self.search_list.bind("<Escape>", lambda e: self.search_entry.focus_set())

        hdr = tk.Frame(left, bg=SURFACE2, pady=7)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Amount × code", bg=SURFACE2, fg=MUTED,
                 font=("Segoe UI", 8, "bold"), anchor="w").pack(side="left", padx=(14, 4))

        outer = tk.Frame(left, bg=BG)
        outer.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(outer, bg=BG, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(fill="both", expand=True)
        self._rows_frame = tk.Frame(self._canvas, bg=BG)
        self._win = self._canvas.create_window((0, 0), window=self._rows_frame, anchor="nw")
        self._rows_frame.bind("<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
            lambda e: self._canvas.itemconfig(self._win, width=e.width))
        self._canvas.bind("<MouseWheel>",
            lambda e: self._canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._sum_frame = tk.Frame(body, bg=SURFACE)
        self._sum_frame.pack(side="left", fill="both", expand=True, padx=(1, 0))
        tk.Label(self._sum_frame, text="SUMMARY — codes auto-resolved by total quantity",
                 bg=SURFACE, fg=MUTED, font=("Segoe UI", 8, "bold"), anchor="w"
                 ).pack(fill="x", padx=16, pady=(14, 6))
        tk.Frame(self._sum_frame, bg=BORDER, height=1).pack(fill="x", padx=16)
        self._sum_inner = tk.Frame(self._sum_frame, bg=SURFACE)
        self._sum_inner.pack(fill="both", expand=True, padx=2, pady=(8, 12))

        self._add_row()

    def _add_row(self, value=""):
        frame = tk.Frame(self._rows_frame, bg=ROW_BG, pady=4)
        frame.pack(fill="x", pady=1)

        var = tk.StringVar(value=value)
        field = tk.Entry(frame, textvariable=var, bg=SURFACE2, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         font=("Segoe UI", 11, "bold"))
        field.pack(side="left", fill="x", expand=True, padx=(14, 6), pady=2)

        del_btn = tk.Button(frame, text="✕", bg=ROW_BG, fg=MUTED,
                            relief="flat", cursor="hand2", font=("Segoe UI", 9),
                            activebackground=DANGER, activeforeground=TEXT,
                            takefocus=False,
                            command=lambda f=frame: self._delete_row(f))
        del_btn.pack(side="left", padx=(2, 10))

        row = {"frame": frame, "value": var, "entry": field}
        self.rows.append(row)
        var.trace_add("write", lambda *_: self._refresh_summary())
        field.bind("<Return>", lambda e: self._add_row_and_focus())
        field.bind("<Tab>", lambda e, w=field: self._tab_from_entry(w))
        field.bind("<FocusIn>", lambda e, r=row: setattr(self, "active_row", r))

        self.after(50, lambda: self._canvas.yview_moveto(1.0))
        self._refresh_summary()
        return row

    def _normalise_search_text(self, value):
        value = unicodedata.normalize("NFKD", str(value).lower())
        value = "".join(ch for ch in value if not unicodedata.combining(ch))
        for ch in ("×", "–", "—", "-", "/", ",", ".", "(", ")"):
            value = value.replace(ch, " ")
        return " ".join(value.split())

    def _search_catalog(self, query, limit=12):
        q = self._normalise_search_text(query)
        if not q:
            return []
        # Search accepts approximate Lithuanian or English wording.
        synonym_map = {
            "remelis":"frame", "remeliai":"frame", "pasportas":"backing board",
            "klijai":"glue adhesive", "klijus":"glue adhesive", "lipalas":"pva glue",
            "vokas":"envelope", "vokai":"envelope", "imaute":"sleeve", "imautes":"sleeve",
            "segtuvas":"binder file", "segtuvai":"binder file", "spirale":"spiral binding",
            "spirales":"spiral binding", "irisimas":"binding", "virselis":"cover",
            "virseliai":"cover", "nugarele":"back cover", "laminavimas":"lamination",
            "skenavimas":"scan scanning", "skenuoti":"scan scanning", "spausdinimas":"print printing",
            "kopijavimas":"copy copying", "spalvotas":"colour color", "spalvota":"colour color",
            "nespalvotas":"black white bw", "nespalvota":"black white bw",
            "vienpusis":"one sided", "dvipusis":"double sided", "popierius":"paper",
            "lipnus":"adhesive sticky", "lipni":"adhesive sticky", "teksturinis":"textured",
            "blizgus":"glossy", "matinis":"matte", "skaidrus":"clear transparent",
            "zirkles":"scissors", "piestukas":"pencil", "rasiklis":"pen", "tusinukas":"pen",
            "flomasteriai":"marker set", "markeriai":"markers", "korektorius":"correction fluid",
            "maiseliai":"bags", "juostele":"ribbon tape", "kaspinas":"ribbon",
            "kalendoriai":"calendars", "skylamusis":"hole punch", "susegimas":"stapling",
            "pjaustymas":"cutting", "pristatymas":"delivery", "siuntimas":"delivery",
            "puodelis":"mug", "maikute":"t shirt", "marskineliai":"t shirt",
            "drobe":"canvas", "antspaudas":"stamp", "medalis":"medal", "taure":"trophy",
            "deze":"box", "dezute":"box", "tuta":"tube", "atvirukai":"postcards",
            "usb":"usb flash drive", "atmintine":"usb flash drive", "magnetas":"magnet",
            "delione":"jigsaw puzzle", "zenkliukas":"badge", "raktu":"keychain",
            "smeigtukai":"push pins", "savarzeles":"paper clips", "sasageles":"staples",
            "spaustukai":"binder clips", "fotopopierius":"photo paper",
        }
        expanded = [q]
        for token in q.split():
            if token in synonym_map:
                expanded.append(synonym_map[token])
        q_expanded = self._normalise_search_text(" ".join(expanded))
        q_tokens = q_expanded.split()
        ranked = []
        for code, description in PLU_DESCRIPTIONS.items():
            price = PLU_PRICES.get(code)
            hay = self._normalise_search_text(f"{code} {description}")
            words = hay.split()
            score = 0.0
            if code == q:
                score += 1000
            elif code.startswith(q):
                score += 300
            if q in hay:
                score += 180
            if q_expanded in hay:
                score += 80
            for token in q_tokens:
                if token in words:
                    score += 75
                elif any(word.startswith(token) for word in words):
                    score += 50
                elif any(token in word for word in words):
                    score += 30
                else:
                    best = max((difflib.SequenceMatcher(None, token, word).ratio() for word in words), default=0)
                    if best >= 0.72:
                        score += best * 22
            coverage = sum(1 for token in q_tokens if token in hay)
            score += coverage * 35
            score += max(
                difflib.SequenceMatcher(None, q, hay).ratio(),
                difflib.SequenceMatcher(None, q_expanded, hay).ratio(),
            ) * 18
            if score > 22:
                ranked.append((score, code, description, price))
        ranked.sort(key=lambda item: (-item[0], int(item[1]) if item[1].isdigit() else 10**9, item[1]))
        return ranked[:limit]

    def _update_search_results(self):
        results = self._search_catalog(self.search_var.get())
        self.search_codes = [item[1] for item in results]
        self.search_list.delete(0, "end")
        for _, code, description, price in results:
            price_text = f"€{price:.2f}" if price is not None else "no price"
            self.search_list.insert("end", f"{code}   {description}   ·   {price_text}")
        if results:
            if not self.search_list.winfo_ismapped():
                self.search_list.pack(fill="x", pady=(6, 0))
            self.search_list.selection_set(0)
        else:
            self.search_list.pack_forget()

    def _focus_search_results(self):
        if self.search_codes:
            self.search_list.focus_set()
            self.search_list.selection_clear(0, "end")
            self.search_list.selection_set(0)
            self.search_list.activate(0)
        return "break"

    def _use_first_search_result(self):
        if self.search_codes:
            self._insert_search_code(self.search_codes[0])
        return "break"

    def _use_selected_search_result(self):
        selected = self.search_list.curselection()
        if selected and selected[0] < len(self.search_codes):
            self._insert_search_code(self.search_codes[selected[0]])
        return "break"

    def _insert_search_code(self, code):
        target = self.active_row if self.active_row in self.rows else None
        if target is None or target["value"].get().strip():
            target = next((row for row in self.rows if not row["value"].get().strip()), None)
        if target is None:
            target = self._add_row()
        target["value"].set(f"1*{code}")
        self.active_row = target
        self.search_var.set("")
        target["entry"].focus_set()
        target["entry"].selection_range(0, 1)
        target["entry"].icursor(1)

    def _tab_from_entry(self, widget):
        for i, row in enumerate(self.rows):
            if row["entry"] is widget:
                if i + 1 < len(self.rows):
                    nxt = self.rows[i + 1]["entry"]
                    nxt.focus_set()
                    nxt.select_range(0, "end")
                else:
                    self._add_row_and_focus()
                return "break"
        return "break"

    def _add_row_and_focus(self):
        row = self._add_row()
        self.after(60, lambda: row["entry"].focus_set())
        return "break"

    def _delete_row(self, frame):
        self.rows = [r for r in self.rows if r["frame"] is not frame]
        frame.destroy()
        if not self.rows:
            self._add_row()
        else:
            self._refresh_summary()

    def _clear_all(self):
        for row in self.rows:
            row["frame"].destroy()
        self.rows.clear()
        self._add_row()

    def _parse_entry(self, text):
        """Parse amount*code. A code by itself is treated as 1*code."""
        raw = text.strip().replace("×", "*")
        if not raw:
            return None
        if "*" in raw:
            amount_text, code = raw.split("*", 1)
            code = code.strip()
            try:
                amount = float(amount_text.strip().replace(",", "."))
            except ValueError:
                return None
        else:
            code = raw
            amount = 1.0
        if not code or amount <= 0:
            return None
        return amount, code

    def _collect_totals(self):
        totals = {}
        for row in self.rows:
            parsed = self._parse_entry(row["value"].get())
            if not parsed:
                continue
            amount, code = parsed
            totals[code] = totals.get(code, 0.0) + amount
        return totals

    def _fmt_qty(self, value):
        if value == int(value):
            return str(int(value))
        return f"{value:.2f}".rstrip("0").rstrip(".")

    def _refresh_summary(self):
        for widget in self._sum_inner.winfo_children():
            widget.destroy()
        totals = self._collect_totals()

        if not totals:
            tk.Label(self._sum_inner,
                     text="No valid entries yet\n\nExample: 8*710",
                     bg=SURFACE, fg=MUTED, justify="left",
                     font=("Segoe UI", 10)).pack(padx=14, pady=6, anchor="nw")
            return

        header = tk.Frame(self._sum_inner, bg=SURFACE)
        header.pack(fill="x", padx=14, pady=(2, 3))
        columns = (("Input", 8), ("Qty", 6), ("PLU", 7),
                   ("Description", 31), ("Unit €", 8), ("Total €", 10))
        for text, width in columns:
            tk.Label(header, text=text, bg=SURFACE, fg=MUTED,
                     font=("Segoe UI", 8, "bold"), width=width,
                     anchor="w").pack(side="left")
        tk.Frame(self._sum_inner, bg=BORDER, height=1).pack(fill="x", padx=14, pady=(0, 4))

        grand_total = 0.0
        missing_prices = 0
        for code, total in sorted(totals.items()):
            output = resolve_output_code(code, total)
            price = get_plu_price(output)
            line_total = total * price if price is not None else None
            if line_total is not None:
                grand_total += line_total
            else:
                missing_prices += 1

            line = tk.Frame(self._sum_inner, bg=SURFACE)
            line.pack(fill="x", padx=14, pady=3)
            description = get_plu_description(output) or "Unknown code"
            values = (
                (code, 8, MUTED if output != code else TEXT, 9),
                (self._fmt_qty(total), 6, TEXT, 9),
                (output, 7, ACCENT if output != code else TEXT, 10),
                (description, 31, TEXT, 9),
                (f"{price:.2f}" if price is not None else "—", 8, TEXT if price is not None else WARNING, 9),
                (f"{line_total:.2f}" if line_total is not None else "—", 10, SUCCESS if line_total is not None else WARNING, 10),
            )
            for text, width, color, size in values:
                tk.Label(line, text=text, bg=SURFACE, fg=color,
                         font=("Segoe UI", size, "bold" if size >= 11 else "normal"),
                         width=width, anchor="w").pack(side="left")

        tk.Frame(self._sum_inner, bg=BORDER, height=1).pack(fill="x", padx=14, pady=(8, 6))
        total_row = tk.Frame(self._sum_inner, bg=SURFACE)
        total_row.pack(fill="x", padx=14)
        tk.Label(total_row, text="GRAND TOTAL", bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 11, "bold")).pack(side="left")
        tk.Label(total_row, text=f"€{grand_total:.2f}", bg=SURFACE, fg=SUCCESS,
                 font=("Segoe UI", 18, "bold")).pack(side="right")
        if missing_prices:
            tk.Label(self._sum_inner,
                     text=f"{missing_prices} line(s) have no stored price",
                     bg=SURFACE, fg=WARNING, font=("Segoe UI", 8)
                     ).pack(anchor="e", padx=14, pady=(4, 0))

    def _copy_summary(self):
        totals = self._collect_totals()
        if not totals:
            messagebox.showinfo("Nothing to copy", "No valid amount*code entries yet.")
            return
        lines = ["Input code\tQty\tOutput PLU\tDescription\tUnit EUR\tLine total EUR", "-" * 96]
        grand_total = 0.0
        for code, total in sorted(totals.items()):
            output = resolve_output_code(code, total)
            price = get_plu_price(output)
            if price is None:
                unit_text = line_text = "—"
            else:
                unit_text = f"{price:.2f}"
                line_total = total * price
                line_text = f"{line_total:.2f}"
                grand_total += line_total
            description = get_plu_description(output) or "Unknown code"
            lines.append(f"{code}\t{self._fmt_qty(total)}\t{output}\t{description}\t{unit_text}\t{line_text}")
        lines.extend(["-" * 96, f"GRAND TOTAL\t\t\t\t\t{grand_total:.2f} EUR"])
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        messagebox.showinfo("Copied", "Priced summary copied to clipboard.")


    def _open_wide_format_quote(self):
        try:
            reload_editable_data()
        except Exception as ex:
            messagebox.showerror(
                "Excel data error",
                f"Could not reload the active code workbook.\n\n{ex}",
                parent=self,
            )
            return

        popup = tk.Toplevel(self)
        popup.title("Wide format quote")
        popup.configure(bg=BG)
        popup.geometry("760x620")
        popup.minsize(650, 500)
        popup.transient(self.winfo_toplevel())
        popup.grab_set()

        title = tk.Frame(popup, bg=SURFACE, pady=12)
        title.pack(fill="x")
        tk.Label(title, text="WIDE FORMAT PRICE LIST", bg=SURFACE, fg=TEXT,
                 font=("Segoe UI", 12, "bold")).pack(side="left", padx=16)
        tk.Label(title, text="Amount × unit price, exact total", bg=SURFACE, fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="right", padx=16)

        controls = tk.Frame(popup, bg=BG)
        controls.pack(fill="x", padx=16, pady=14)

        tk.Label(controls, text="Amount", bg=BG, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        amount_var = tk.StringVar(value="4.3")
        amount_entry = tk.Entry(controls, textvariable=amount_var, bg=SURFACE2, fg=TEXT,
                                insertbackground=TEXT, relief="flat",
                                font=("Segoe UI", 11, "bold"), width=12)
        amount_entry.grid(row=1, column=0, sticky="w", pady=(3, 0), ipady=4)

        tk.Label(controls, text="Print coverage", bg=BG, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=1, sticky="w", padx=(24, 0))
        coverage_var = tk.StringVar(value="partial")
        coverage_row = tk.Frame(controls, bg=BG)
        coverage_row.grid(row=1, column=1, sticky="w", padx=(20, 0))
        for text, value in (("Not full — tekstas+pav.", "partial"),
                            ("Full — paveikslas", "full")):
            tk.Radiobutton(coverage_row, text=text, variable=coverage_var, value=value,
                           bg=BG, fg=TEXT, selectcolor=SURFACE2,
                           activebackground=BG, activeforeground=TEXT,
                           font=("Segoe UI", 9), command=lambda: refresh()).pack(side="left", padx=4)

        note = tk.Label(popup,
            text="Use the same amount you would enter before the PLU, for example 4.3 for an A2 job.",
            bg=BG, fg=MUTED, font=("Segoe UI", 8), anchor="w")
        note.pack(fill="x", padx=16, pady=(0, 8))

        result_frame = tk.Frame(popup, bg=SURFACE)
        result_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        result = tk.Text(result_frame, bg=SURFACE, fg=TEXT, insertbackground=TEXT,
                         relief="flat", bd=0, highlightthickness=0,
                         font=("Consolas", 10), wrap="none", padx=12, pady=12)
        ybar = ttk.Scrollbar(result_frame, orient="vertical", command=result.yview)
        xbar = ttk.Scrollbar(result_frame, orient="horizontal", command=result.xview)
        result.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        ybar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")
        result.pack(fill="both", expand=True)

        def parse_amount():
            try:
                amount = float(amount_var.get().strip().replace(",", "."))
                if amount <= 0:
                    raise ValueError
                return amount
            except ValueError:
                return None

        def build_lines():
            amount = parse_amount()
            if amount is None:
                return ["Enter a valid positive amount."]

            coverage = coverage_var.get()
            # Short customer-facing labels. Totals are kept at exact cent precision.
            labels = {
                "4800": "80gsm", "4810": "80gsm", "4820": "80gsm",
                "4830": "120gsm", "4840": "120gsm",
                "4913": "140gsm", "4914": "140gsm",
                "4850": "180gsm", "4860": "180gsm",
                "4894": "Satin", "4893": "Satin",
                "4915": "Sintetinė drobė", "4900": "Natūrali drobė",
                "4902": "Film plėvelė", "4903": "Film plėvelė",
                "4908": "PVC", "4909": "PVC",
                "4910": "Kalkė", "4911": "Kalkė", "4912": "Kalkė",
                "3009": "Vatmanas", "4916": "Karštas laminatas",
                "4917": "Magnetinė plėvelė", "3008": "Vatmanas 90×64cm",
            }

            lines = []
            # Exact top-to-bottom order from copypro_kodai.xlsx.
            for code in WIDE_FORMAT_ORDER:
                item=WIDE_FORMAT_ITEMS.get(code)
                if not item: continue
                _description,price,item_coverage,item_label=item
                searchable = f"{_description} {item_label}".lower()
                if "vatman" in searchable or code in {"3008", "3009"}:
                    continue
                include = item_coverage in (coverage, "fixed")
                if item_coverage == "drawing":
                    include = coverage == "partial"
                if not include:
                    continue
                total = amount * price
                lines.append(f"{item_label or labels.get(code, _description)} - {total:.2f} Eur")
            return lines

        def refresh(*_):
            result.configure(state="normal")
            result.delete("1.0", "end")
            result.insert("1.0", "\n".join(build_lines()))
            result.configure(state="disabled")

        def copy_text():
            lines = build_lines()
            if len(lines) == 1 and lines[0].startswith("Enter"):
                messagebox.showwarning("Invalid amount", lines[0], parent=popup)
                amount_entry.focus_set()
                return
            self.clipboard_clear()
            self.clipboard_append("\n".join(lines))
            messagebox.showinfo("Copied", "Wide-format price list copied to clipboard.", parent=popup)

        buttons = tk.Frame(popup, bg=BG)
        buttons.pack(fill="x", padx=16, pady=(0, 14))
        styled_btn(buttons, "Copy price list", copy_text, style="success").pack(side="left")
        styled_btn(buttons, "Close", popup.destroy, style="secondary").pack(side="right")

        amount_var.trace_add("write", refresh)
        amount_entry.bind("<Return>", lambda e: copy_text())
        popup.bind("<Escape>", lambda e: popup.destroy())
        refresh()
        amount_entry.focus_set()
        amount_entry.select_range(0, "end")


# ── PDF Page Counter Tab ─────────────────────────────────────────────────────
class PageCounterTab(tk.Frame, DropMixin):
    """Analyse PDF pages by colour usage and physical page size."""
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.files = []
        self.results = []
        self._build()

    def _build(self):
        left = tk.Frame(self, bg=SURFACE, width=275)
        left.pack(side="left", fill="y", padx=(0,1))
        left.pack_propagate(False)

        lbl(left,"PAGE COUNTER",8,MUTED,True).pack(pady=(16,4),padx=16,anchor="w")
        tk.Label(left, text="Analyse colour usage and detect pages\nthat differ from the dominant paper size.",
                 bg=SURFACE, fg=MUTED, justify="left", font=("Segoe UI",8),
                 wraplength=235).pack(padx=16, anchor="w", pady=(0,10))

        lbl(left,"Colour detection sensitivity").pack(padx=16,anchor="w")
        self.tolerance_var = tk.IntVar(value=12)
        self.tolerance_lbl = lbl(left,"Tolerance: 12",8,MUTED)
        self.tolerance_lbl.pack(padx=16,anchor="w")
        ttk.Scale(left,from_=2,to=35,variable=self.tolerance_var,orient="horizontal",
                  command=lambda v:self.tolerance_lbl.config(text=f"Tolerance: {int(float(v))}")) \
            .pack(fill="x",padx=16,pady=(0,8))

        lbl(left,"Odd-size tolerance").pack(padx=16,anchor="w")
        size_row=tk.Frame(left,bg=SURFACE); size_row.pack(fill="x",padx=16,pady=(2,8))
        self.size_tol_var=tk.StringVar(value="1.0")
        entry(size_row,self.size_tol_var,6).pack(side="left")
        lbl(size_row,"mm",8,MUTED).pack(side="left",padx=5)

        sep(left)
        lbl(left,"EXPORT",8,MUTED,True).pack(pady=(3,4),padx=16,anchor="w")
        tk.Label(left,text="After analysis, choose colour/B&W pages,\ndetected sizes, and whether blank pages\nshould be removed.",
                 bg=SURFACE,fg=MUTED,justify="left",font=("Segoe UI",8)) \
            .pack(padx=16,anchor="w",pady=(0,5))
        tk.Label(left,text="Pages are copied directly from the source PDF,\nso original vectors and image quality remain intact.",
                 bg=SURFACE,fg=MUTED,justify="left",font=("Segoe UI",8)) \
            .pack(padx=16,anchor="w",pady=(0,4))

        sep(left)
        addrow=tk.Frame(left,bg=SURFACE); addrow.pack(fill="x",padx=16,pady=2)
        styled_btn(addrow,"➕ Add",self._add_files).pack(side="left",fill="x",expand=True,padx=(0,3))
        styled_btn(addrow,"🗑 Clear",self._clear,style="secondary").pack(side="left",fill="x",expand=True,padx=(3,0))
        self.analyse_btn=styled_btn(left,"🔎 Analyse PDFs",self._analyse,style="primary")
        self.analyse_btn.pack(fill="x",padx=16,pady=(5,3))
        self.extract_btn=styled_btn(left,"📄 Export pages…",self._extract,style="success")
        self.extract_btn.pack(fill="x",padx=16,pady=3)
        self.extract_btn.config(state="disabled")
        self.status=lbl(left,"",8,MUTED); self.status.pack(padx=16,pady=7,anchor="w")

        right=tk.Frame(self,bg=BG); right.pack(side="left",fill="both",expand=True,padx=12,pady=12)
        self.drop_zone=tk.Label(right,text="⬇  Drop PDF files here",bg=SURFACE2,fg=MUTED,
                                font=("Segoe UI",10),pady=10,cursor="hand2")
        self.drop_zone.pack(fill="x")
        self.drop_zone.bind("<Button-1>",lambda e:self._add_files())
        self.setup_drop(self.drop_zone,self._drop_files,{".pdf"})

        summary=tk.Frame(right,bg=BG); summary.pack(fill="x",pady=(8,6))
        self.total_card=self._card(summary,"Total pages")
        self.colour_card=self._card(summary,"Colour")
        self.bw_card=self._card(summary,"B&W")
        self.odd_card=self._card(summary,"Odd size")
        self.blank_card=self._card(summary,"Blank")

        table_frame=tk.Frame(right,bg=SURFACE2,bd=0,highlightthickness=0); table_frame.pack(fill="both",expand=True)
        cols=("file","page","colour","size","status")
        self.tree=ttk.Treeview(table_frame,columns=cols,show="headings",selectmode="extended",style="CopyPro.Treeview")
        headings={"file":"File","page":"Page","colour":"Type","size":"Size","status":"Size status"}
        widths={"file":190,"page":55,"colour":95,"size":125,"status":145}
        for c in cols:
            self.tree.heading(c,text=headings[c]); self.tree.column(c,width=widths[c],anchor="w")
        vs=ttk.Scrollbar(table_frame,orient="vertical",command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right",fill="y"); self.tree.pack(fill="both",expand=True)

        self.progress=ttk.Progressbar(right,orient="horizontal",mode="determinate")
        self.progress.pack(fill="x",pady=(7,0))

        # Accept PDF drops anywhere in this tab, not only on the drop banner.
        for target in (self, left, right, table_frame, self.tree, self.drop_zone):
            self.setup_drop(target,self._drop_files,{".pdf"})

    def _card(self,parent,title):
        f=tk.Frame(parent,bg=SURFACE2); f.pack(side="left",fill="x",expand=True,padx=(0,6))
        tk.Label(f,text=title,bg=SURFACE2,fg=MUTED,font=("Segoe UI",8)).pack(anchor="w",padx=9,pady=(6,0))
        v=tk.Label(f,text="—",bg=SURFACE2,fg=TEXT,font=("Segoe UI",16,"bold"))
        v.pack(anchor="w",padx=9,pady=(0,6)); return v

    def _add_files(self):
        paths=filedialog.askopenfilenames(filetypes=[("PDF files","*.pdf"),("All","*.*")])
        self._drop_files(list(paths))

    def _drop_files(self,paths):
        for path in paths:
            if path.lower().endswith(".pdf") and path not in self.files:
                self.files.append(path)
        n=len(self.files)
        self.status.config(text=f"{n} PDF{'s' if n!=1 else ''} queued",fg=MUTED)

    def _clear(self):
        self.files.clear(); self.results.clear()
        for item in self.tree.get_children(): self.tree.delete(item)
        for card in (self.total_card,self.colour_card,self.bw_card,self.odd_card,self.blank_card): card.config(text="—")
        self.progress["value"]=0; self.extract_btn.config(state="disabled")
        self.status.config(text="",fg=MUTED)

    @staticmethod
    def _size_mm(page):
        r=page.rect
        return r.width/72*25.4, r.height/72*25.4

    @staticmethod
    def _size_key(w,h,tol):
        a,b=sorted((w,h)); tol=max(0.1,tol)
        return (round(a/tol)*tol,round(b/tol)*tol)

    @staticmethod
    def _has_colour(page,tolerance):
        mat=fitz.Matrix(110/72,110/72)
        pix=page.get_pixmap(matrix=mat,colorspace=fitz.csRGB,alpha=False)
        data=pix.samples; coloured=0; total=pix.width*pix.height
        required=max(25,int(total*0.00005))
        for i in range(0,len(data),3):
            r,g,b=data[i],data[i+1],data[i+2]
            if max(r,g,b)-min(r,g,b)>tolerance:
                coloured+=1
                if coloured>=required:return True
        return False

    @staticmethod
    def _is_blank(page):
        """Treat a page as blank when almost all rendered pixels are near-white."""
        mat=fitz.Matrix(72/72,72/72)
        pix=page.get_pixmap(matrix=mat,colorspace=fitz.csGRAY,alpha=False)
        data=pix.samples; total=max(1,pix.width*pix.height)
        required=max(30,int(total*0.00008))
        nonwhite=0
        for value in data:
            if value < 245:
                nonwhite += 1
                if nonwhite >= required:
                    return False
        return True

    def _analyse(self):
        if not self.files:
            messagebox.showwarning("Nothing","Add at least one PDF first."); return
        try: size_tol=float(self.size_tol_var.get().replace(",","."))
        except: size_tol=1.0
        self.analyse_btn.config(state="disabled"); self.extract_btn.config(state="disabled")
        self.progress["value"]=0; self.status.config(text="Analysing…",fg=MUTED)
        threading.Thread(target=self._do_analyse,args=(max(.1,size_tol),),daemon=True).start()

    def _do_analyse(self,size_tol):
        tolerance=int(self.tolerance_var.get()); results=[]; total_pages=0
        for path in self.files:
            try:
                with fitz.open(path) as d: total_pages+=d.page_count
            except: pass
        done=0
        for path in self.files:
            doc=None
            try:
                doc=fitz.open(path); pages=[]; counts={}
                for idx,page in enumerate(doc):
                    w,h=self._size_mm(page); key=self._size_key(w,h,size_tol)
                    counts[key]=counts.get(key,0)+1
                    colour=self._has_colour(page,tolerance)
                    blank=self._is_blank(page)
                    pages.append({"index":idx,"page":idx+1,"colour":colour,"blank":blank,"w":w,"h":h,"key":key})
                    done+=1
                    self.after(0,lambda d=done,t=total_pages,n=os.path.basename(path):self._progress(d,t,n))
                dominant=max(counts,key=counts.get) if counts else None
                for item in pages:item["odd"]=(item["key"]!=dominant)
                results.append({"path":path,"pages":pages,"dominant":dominant})
            except Exception as ex:
                results.append({"path":path,"pages":[],"dominant":None,"error":str(ex)})
            finally:
                if doc: doc.close()
        self.after(0,lambda:self._finish_analysis(results))

    def _progress(self,done,total,name):
        self.progress["maximum"]=max(1,total); self.progress["value"]=done
        self.status.config(text=f"Analysing {name}: {done}/{total}")

    def _finish_analysis(self,results):
        self.results=results
        for item in self.tree.get_children():self.tree.delete(item)
        total=colour=bw=odd=blank=0
        for result in results:
            name=os.path.basename(result["path"])
            for p in result["pages"]:
                total+=1; colour+=int(p["colour"]); bw+=int(not p["colour"]); odd+=int(p["odd"]); blank+=int(p.get("blank",False))
                size=f"{p['w']:.1f} × {p['h']:.1f} mm"
                type_text="Blank" if p.get("blank",False) else ("Colour" if p["colour"] else "B&W")
                self.tree.insert("", "end", values=(name,p["page"],type_text,size,
                                                     "Non-dominant" if p["odd"] else "Dominant"))
        self.total_card.config(text=str(total)); self.colour_card.config(text=str(colour))
        self.bw_card.config(text=str(bw)); self.odd_card.config(text=str(odd)); self.blank_card.config(text=str(blank))
        self.analyse_btn.config(state="normal"); self.extract_btn.config(state="normal" if total else "disabled")
        self.status.config(text=f"Finished: {colour} colour, {bw} B&W, {odd} odd-size, {blank} blank",fg=SUCCESS)

    def _detected_sizes(self):
        counts={}
        labels={}
        for result in self.results:
            for p in result.get("pages",[]):
                key=p["key"]
                counts[key]=counts.get(key,0)+1
                labels[key]=f"{key[0]:g} × {key[1]:g} mm"
        return sorted(counts.items(), key=lambda item:(-item[1], item[0])), labels

    def _extract(self):
        if not self.results:return
        popup=tk.Toplevel(self)
        popup.title("Export analysed pages")
        popup.configure(bg=BG)
        popup.resizable(False,False)
        popup.grab_set()
        self.update_idletasks()
        px=self.winfo_rootx()+self.winfo_width()//2-230
        py=self.winfo_rooty()+self.winfo_height()//2-260
        popup.geometry(f"460x520+{px}+{py}")

        tk.Label(popup,text="Export analysed pages",bg=BG,fg=TEXT,
                 font=("Segoe UI",13,"bold")).pack(anchor="w",padx=18,pady=(16,4))
        tk.Label(popup,text="Selected filters are combined. A page is exported when it matches\nthe chosen colour type and one of the chosen sizes.",
                 bg=BG,fg=MUTED,justify="left",font=("Segoe UI",8)).pack(anchor="w",padx=18,pady=(0,10))

        body=tk.Frame(popup,bg=SURFACE); body.pack(fill="both",expand=True,padx=18,pady=(0,10))

        lbl(body,"PAGE TYPE",8,MUTED,True).pack(anchor="w",padx=14,pady=(12,4))
        colour_var=tk.BooleanVar(value=True)
        bw_var=tk.BooleanVar(value=True)
        for var,text in ((colour_var,"Colour pages"),(bw_var,"Black & white / grayscale pages")):
            tk.Checkbutton(body,text=text,variable=var,bg=SURFACE,fg=TEXT,
                           selectcolor=SURFACE2,activebackground=SURFACE,activeforeground=TEXT,
                           relief="flat",bd=0,highlightthickness=0,font=("Segoe UI",9)).pack(anchor="w",padx=14,pady=1)

        remove_blank_var=tk.BooleanVar(value=True)
        tk.Checkbutton(body,text="Remove blank pages",variable=remove_blank_var,
                       bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,activeforeground=TEXT,
                       relief="flat",bd=0,highlightthickness=0,font=("Segoe UI",9,"bold")).pack(anchor="w",padx=14,pady=(7,3))

        sep(body)
        lbl(body,"DETECTED SIZES",8,MUTED,True).pack(anchor="w",padx=14,pady=(3,4))
        tk.Label(body,text="Choose one or more sizes. Leave all selected to keep every size.",
                 bg=SURFACE,fg=MUTED,font=("Segoe UI",8)).pack(anchor="w",padx=14,pady=(0,4))

        size_canvas=tk.Canvas(body,bg=SURFACE2,highlightthickness=0,height=185)
        size_canvas.pack(fill="both",expand=True,padx=14,pady=(0,8))
        size_inner=tk.Frame(size_canvas,bg=SURFACE2)
        win=size_canvas.create_window((0,0),window=size_inner,anchor="nw")
        size_inner.bind("<Configure>",lambda e:size_canvas.configure(scrollregion=size_canvas.bbox("all")))
        size_canvas.bind("<Configure>",lambda e:size_canvas.itemconfig(win,width=e.width))

        size_vars={}
        sizes,labels=self._detected_sizes()
        for key,count in sizes:
            var=tk.BooleanVar(value=True); size_vars[key]=var
            tk.Checkbutton(size_inner,text=f"{labels[key]}   ({count} page{'s' if count!=1 else ''})",
                           variable=var,bg=SURFACE2,fg=TEXT,selectcolor=SURFACE,activebackground=SURFACE2,activeforeground=TEXT,
                           relief="flat",bd=0,highlightthickness=0,font=("Segoe UI",9)).pack(anchor="w",padx=8,pady=2)

        def set_all(value):
            for var in size_vars.values():var.set(value)
        small=tk.Frame(body,bg=SURFACE); small.pack(fill="x",padx=14,pady=(0,8))
        styled_btn(small,"Select all",lambda:set_all(True),style="secondary",padx=9,pady=4).pack(side="left")
        styled_btn(small,"Clear sizes",lambda:set_all(False),style="secondary",padx=9,pady=4).pack(side="left",padx=6)

        def run_export():
            types=[]
            if colour_var.get():types.append("colour")
            if bw_var.get():types.append("bw")
            chosen_sizes={key for key,var in size_vars.items() if var.get()}
            if not types:
                messagebox.showwarning("Nothing selected","Choose at least one page type.",parent=popup); return
            if not chosen_sizes:
                messagebox.showwarning("Nothing selected","Choose at least one detected page size.",parent=popup); return
            out=filedialog.askdirectory(title="Output folder",parent=popup)
            if not out:return
            popup.destroy()
            self.extract_btn.config(state="disabled")
            self.status.config(text="Exporting original PDF pages…",fg=MUTED)
            threading.Thread(target=self._do_extract,
                             args=(out,types,chosen_sizes,remove_blank_var.get()),daemon=True).start()

        btns=tk.Frame(popup,bg=BG); btns.pack(fill="x",padx=18,pady=(0,16))
        styled_btn(btns,"Export",run_export,style="success").pack(side="right")
        styled_btn(btns,"Cancel",popup.destroy,style="secondary").pack(side="right",padx=8)

    def _do_extract(self,out,types,chosen_sizes,remove_blank):
        made=[]; errors=[]
        for result in self.results:
            path=result["path"]; base=os.path.splitext(os.path.basename(path))[0]
            try:
                src=fitz.open(path)
                indices=[]
                for p in result.get("pages",[]):
                    type_ok=("colour" in types and p["colour"]) or ("bw" in types and not p["colour"])
                    size_ok=p["key"] in chosen_sizes
                    blank_ok=not (remove_blank and p.get("blank",False))
                    if type_ok and size_ok and blank_ok:
                        indices.append(p["index"])
                if indices:
                    dest=fitz.open()
                    for idx in indices:dest.insert_pdf(src,from_page=idx,to_page=idx)
                    suffix=[]
                    if set(types)=={"colour"}:suffix.append("colour")
                    elif set(types)=={"bw"}:suffix.append("bw")
                    else:suffix.append("selected")
                    if len(chosen_sizes)==1:
                        key=next(iter(chosen_sizes)); suffix.append(f"{key[0]:g}x{key[1]:g}mm")
                    if remove_blank:suffix.append("no_blanks")
                    out_path=_unique(os.path.join(out,f"{base}_{'_'.join(suffix)}.pdf"))
                    dest.save(out_path,garbage=4,deflate=True); dest.close(); made.append(out_path)
                src.close()
            except Exception as ex:errors.append(f"{os.path.basename(path)}: {ex}")
        self.after(0,lambda:self._finish_extract(made,errors))

    def _finish_extract(self,made,errors):
        self.extract_btn.config(state="normal")
        n=len(made)
        self.status.config(text=f"Created {n} PDF{'s' if n!=1 else ''}",fg=SUCCESS if not errors else WARNING)
        detail=f"Created {n} exported PDF file(s)."
        if not made and not errors:detail+="\n\nNo pages matched the selected filters."
        if errors:detail+="\n\nErrors:\n"+"\n".join(errors[:6])
        messagebox.showinfo("Export complete",detail)


# ── Batch Organiser Tab ───────────────────────────────────────────────────────
class BatchOrganiserTab(tk.Frame, DropMixin):
    """Analyse, group and copy large batches of PDFs safely."""
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        self.files=[]
        self.records=[]
        self.prefs=load_app_settings().get("batch_organiser", {})
        self._build()

    @staticmethod
    def _paper_name(w,h):
        a,b=sorted((w,h))
        best=None
        for name,(pw,ph) in PAPER_SIZES_MM.items():
            x,y=sorted((float(pw),float(ph)))
            error=max(abs(a-x),abs(b-y))
            if best is None or error<best[0]: best=(error,name)
        if best and best[0] <= 2.0: return best[1]
        return f"{a:.0f}x{b:.0f}mm"

    def _build(self):
        left=tk.Frame(self,bg=SURFACE,width=285); left.pack(side="left",fill="y",padx=(0,1)); left.pack_propagate(False)
        lbl(left,"BATCH ORGANISER",8,MUTED,True).pack(pady=(15,4),padx=16,anchor="w")
        tk.Label(left,text="Analyse PDFs and organise copies into folders\nwithout changing the original files.",
                 bg=SURFACE,fg=MUTED,justify="left",font=("Segoe UI",8)).pack(padx=16,anchor="w",pady=(0,10))

        lbl(left,"Group files by").pack(padx=16,anchor="w")
        self.group_var=tk.StringVar(value=self.prefs.get("group","Dominant paper size") if self.prefs.get("group") in ("Dominant paper size","Colour type","No folders") else "Dominant paper size")
        ttk.Combobox(left,textvariable=self.group_var,state="readonly",width=23,
                     values=["Dominant paper size","Colour type","No folders"]).pack(fill="x",padx=16,pady=(2,8))

        self.split_var=tk.BooleanVar(value=self.prefs.get("split_mixed",True))
        tk.Checkbutton(left,text="Split mixed-size PDFs by page size",variable=self.split_var,
                       bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,activeforeground=TEXT,
                       relief="flat",bd=0,highlightthickness=0,font=("Segoe UI",9)).pack(padx=16,anchor="w",pady=2)
        self.blank_var=tk.BooleanVar(value=self.prefs.get("remove_blank",False))
        tk.Checkbutton(left,text="Remove blank pages when splitting",variable=self.blank_var,
                       bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,activeforeground=TEXT,
                       relief="flat",bd=0,highlightthickness=0,font=("Segoe UI",9)).pack(padx=16,anchor="w",pady=2)

        sep(left)
        row=tk.Frame(left,bg=SURFACE); row.pack(fill="x",padx=16,pady=2)
        styled_btn(row,"➕ Add",self._add).pack(side="left",fill="x",expand=True,padx=(0,3))
        styled_btn(row,"🗑 Clear",self._clear,style="secondary").pack(side="left",fill="x",expand=True,padx=(3,0))
        self.an_btn=styled_btn(left,"🔎 Analyse batch",self._analyse); self.an_btn.pack(fill="x",padx=16,pady=(5,3))
        self.org_btn=styled_btn(left,"📁 Organise copies",self._organise,style="success"); self.org_btn.pack(fill="x",padx=16,pady=3); self.org_btn.config(state="disabled")
        self.status=lbl(left,"",8,MUTED); self.status.pack(padx=16,pady=6,anchor="w")
        styled_btn(left,"Open settings folder",lambda:self._open_path(COPYPRO_DATA_DIR),style="secondary",padx=10,pady=4).pack(fill="x",padx=16,pady=(8,0))

        right=tk.Frame(self,bg=BG); right.pack(side="left",fill="both",expand=True,padx=12,pady=12)
        self.drop=tk.Label(right,text="⬇  Drop PDF files here",bg=SURFACE2,fg=MUTED,font=("Segoe UI",10),pady=10,cursor="hand2")
        self.drop.pack(fill="x"); self.drop.bind("<Button-1>",lambda e:self._add()); self.setup_drop(self.drop,self._drop,{".pdf"})
        frame=tk.Frame(right,bg=SURFACE2,bd=0,highlightthickness=0); frame.pack(fill="both",expand=True,pady=(7,0))
        cols=("file","size","type","notes")
        self.tree=ttk.Treeview(frame,columns=cols,show="headings",selectmode="extended",style="CopyPro.Treeview")
        heads={"file":"File","size":"Dominant size","type":"Colour","notes":"Notes"}
        widths={"file":280,"size":150,"type":110,"notes":220}
        for c in cols:self.tree.heading(c,text=heads[c]); self.tree.column(c,width=widths[c],anchor="w")
        vs=ttk.Scrollbar(frame,orient="vertical",command=self.tree.yview); self.tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right",fill="y"); self.tree.pack(fill="both",expand=True)
        self.progress=ttk.Progressbar(right,orient="horizontal",mode="determinate"); self.progress.pack(fill="x",pady=(7,0))

        # Accept PDF drops anywhere in this tab, not only on the drop banner.
        for target in (self, left, right, frame, self.tree, self.drop):
            self.setup_drop(target,self._drop,{".pdf"})

    def _open_path(self,path):
        try:
            if sys.platform.startswith("win"): os.startfile(path)
            elif sys.platform=="darwin": subprocess.Popen(["open",path])
            else: subprocess.Popen(["xdg-open",path])
        except Exception as ex: messagebox.showerror("Could not open folder",str(ex))

    def _add(self):
        self._drop(list(filedialog.askopenfilenames(filetypes=[("PDF files","*.pdf"),("All","*.*")])))
    def _drop(self,paths):
        for p in paths:
            if p.lower().endswith(".pdf") and p not in self.files:self.files.append(p)
        self.status.config(text=f"{len(self.files)} PDFs queued",fg=MUTED)
    def _clear(self):
        self.files.clear(); self.records.clear()
        for i in self.tree.get_children():self.tree.delete(i)
        self.progress["value"]=0; self.org_btn.config(state="disabled"); self.status.config(text="")

    def _analyse(self):
        if not self.files: messagebox.showwarning("Nothing","Add PDF files first."); return
        self.an_btn.config(state="disabled"); self.org_btn.config(state="disabled"); self.status.config(text="Analysing…",fg=MUTED)
        threading.Thread(target=self._do_analyse,daemon=True).start()

    def _do_analyse(self):
        records=[]
        for num,path in enumerate(self.files,1):
            try:
                doc=fitz.open(path); size_counts={}; colour=0; bw=0; blanks=0; portrait=landscape=0
                for page in doc:
                    w=page.rect.width/72*25.4; h=page.rect.height/72*25.4
                    key=(round(min(w,h)),round(max(w,h))); size_counts[key]=size_counts.get(key,0)+1
                    if w<=h:portrait+=1
                    else:landscape+=1
                    if PageCounterTab._is_blank(page): blanks+=1
                    elif PageCounterTab._has_colour(page,12): colour+=1
                    else:bw+=1
                count=doc.page_count; doc.close()
                dominant=max(size_counts,key=size_counts.get) if size_counts else (0,0)
                size_name=self._paper_name(*dominant)
                ctype="Colour" if colour and not bw else ("B&W" if bw and not colour else ("Blank" if blanks==count else "Mixed"))
                orient="Portrait" if portrait>landscape else ("Landscape" if landscape>portrait else "Mixed")
                mixed=len(size_counts)>1
                notes=[]
                if mixed:notes.append(f"{len(size_counts)} sizes")
                if blanks:notes.append(f"{blanks} blank")
                records.append({"path":path,"pages":count,"sizes":size_counts,"size":size_name,"type":ctype,"orientation":orient,"mixed":mixed,"blank":blanks})
            except Exception as ex:
                records.append({"path":path,"pages":0,"sizes":{},"size":"Error","type":"—","orientation":"—","mixed":False,"blank":0,"error":str(ex)})
            self.after(0,lambda v=num,t=len(self.files):self.progress.config(maximum=t,value=v))
        self.after(0,lambda:self._finish(records))

    def _finish(self,records):
        self.records=records
        for i in self.tree.get_children():self.tree.delete(i)
        errors=0
        for r in records:
            if r.get("error"):errors+=1
            notes=r.get("error") or ((f"{len(r['sizes'])} sizes" if r['mixed'] else "") + (f"; {r['blank']} blank" if r['blank'] else ""))
            self.tree.insert("","end",values=(os.path.basename(r["path"]),r["size"],r["type"],notes.strip("; ")))
        self.an_btn.config(state="normal"); self.org_btn.config(state="normal" if records else "disabled")
        self.status.config(text=f"Analysed {len(records)} files"+(f", {errors} errors" if errors else ""),fg=WARNING if errors else SUCCESS)

    @staticmethod
    def _safe(text):
        bad='<>:"/\\|?*'
        text=''.join('_' if c in bad else c for c in str(text)).strip().strip('.')
        return text or "unnamed"

    def _group_name(self,r):
        mode=self.group_var.get()
        if mode=="Dominant paper size":return r["size"]
        if mode=="Colour type":return r["type"]
        return ""

    def _render_name(self,r,index,size_override=None,part=None):
        """Keep the original filename; append a size only for split outputs."""
        base=os.path.splitext(os.path.basename(r["path"]))[0]
        if part:
            base += f"_{part}"
        return self._safe(base)+".pdf"

    def _save_prefs(self):
        data=load_app_settings(); data["batch_organiser"]={"group":self.group_var.get(),"split_mixed":self.split_var.get(),"remove_blank":self.blank_var.get()}; save_app_settings(data)

    def _organise(self):
        if not self.records:return
        out=filedialog.askdirectory(title="Choose organised output folder")
        if not out:return
        self._save_prefs(); self.org_btn.config(state="disabled"); self.status.config(text="Organising copies…",fg=MUTED)
        threading.Thread(target=self._do_organise,args=(out,),daemon=True).start()

    def _do_organise(self,out):
        import shutil
        made=0; errors=[]
        for index,r in enumerate(self.records,1):
            if r.get("error"):continue
            try:
                group=self._safe(self._group_name(r)) if self._group_name(r) else ""
                dest_dir=os.path.join(out,group) if group else out; os.makedirs(dest_dir,exist_ok=True)
                if self.split_var.get() and r["mixed"]:
                    src=fitz.open(r["path"])
                    grouped={}
                    for pi,page in enumerate(src):
                        if self.blank_var.get() and PageCounterTab._is_blank(page):continue
                        w=page.rect.width/72*25.4; h=page.rect.height/72*25.4; key=(round(min(w,h)),round(max(w,h)))
                        grouped.setdefault(key,[]).append(pi)
                    for key,indices in grouped.items():
                        size=self._paper_name(*key); doc=fitz.open()
                        for pi in indices:doc.insert_pdf(src,from_page=pi,to_page=pi)
                        path=_unique(os.path.join(dest_dir,self._render_name(r,index,size_override=size,part=size)))
                        doc.save(path,garbage=4,deflate=True); doc.close(); made+=1
                    src.close()
                else:
                    path=_unique(os.path.join(dest_dir,self._render_name(r,index)))
                    shutil.copy2(r["path"],path); made+=1
            except Exception as ex:errors.append(f"{os.path.basename(r['path'])}: {ex}")
        self.after(0,lambda:self._done(made,errors,out))

    def _done(self,made,errors,out):
        self.org_btn.config(state="normal"); self.status.config(text=f"Created {made} organised PDF files",fg=WARNING if errors else SUCCESS)
        msg=f"Created {made} organised PDF file(s) in:\n{out}"
        if errors:msg+="\n\nErrors:\n"+"\n".join(errors[:6])
        messagebox.showinfo("Organising complete",msg)

# ── Main App ──────────────────────────────────────────────────────────────────
class CopyProApp(dnd.Tk if HAS_DND else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CopyPro Tools")
        self.geometry("1160x740")
        self.minsize(920,620)
        self.configure(bg=BG)
        self._style(); self._build()
        if DATA_LOAD_ERROR:
            self.after(
                300,
                lambda: messagebox.showerror(
                    "Excel data error",
                    "CopyPro Tools could not load its active Excel data files.\n\n"
                    + DATA_LOAD_ERROR
                    + "\n\nOpen Settings and select the correct workbooks.",
                    parent=self,
                ),
            )
        elif CODE_SYNC_RESULT.get("updated"):
            backup = CODE_SYNC_RESULT.get("backup_file")
            message = (
                "Prekių ir paslaugų kodų Excel failas buvo atnaujintas "
                f"į programos versiją {CODE_SYNC_RESULT.get('version')}.\n\n"
                f"Naujas failas:\n{CODE_SYNC_RESULT.get('codes_file')}"
            )
            if backup:
                message += f"\n\nSeno failo atsarginė kopija:\n{backup}"
            self.after(
                350,
                lambda m=message: messagebox.showinfo(
                    "Kodų sąrašas atnaujintas",
                    m,
                    parent=self,
                ),
            )
        self.after(1800, lambda: self._check_for_updates(manual=False))

    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TCombobox", fieldbackground=SURFACE2, background=SURFACE2,
                    foreground=TEXT, arrowcolor=MUTED, borderwidth=0,
                    relief="flat", padding=6)
        s.map("TCombobox", fieldbackground=[("readonly",SURFACE2)],
              foreground=[("readonly",TEXT)])
        s.configure("TScrollbar", background=SURFACE2, troughcolor=SURFACE,
                    borderwidth=0, arrowsize=12)
        s.configure("Horizontal.TProgressbar", troughcolor=SURFACE,
                    background=ACCENT, borderwidth=0, thickness=6)
        s.configure("TScale", background=SURFACE, troughcolor=SURFACE2, sliderlength=16)
        s.configure("TCheckbutton", background=SURFACE, foreground=TEXT,
                    focuscolor=SURFACE, borderwidth=0)
        s.map("TCheckbutton", background=[("active",SURFACE)],
              foreground=[("active",TEXT)])

        # Dark table styling used by Page Counter and Batch Organiser.
        s.configure("CopyPro.Treeview", background=SURFACE2, fieldbackground=SURFACE2,
                    foreground=TEXT, borderwidth=0, relief="flat", rowheight=28,
                    font=("Segoe UI",9))
        s.map("CopyPro.Treeview", background=[("selected",ACCENT)],
              foreground=[("selected",TEXT)])
        s.configure("CopyPro.Treeview.Heading", background=SURFACE, foreground=MUTED,
                    borderwidth=0, relief="flat", padding=(8,7),
                    font=("Segoe UI",8,"bold"))
        s.map("CopyPro.Treeview.Heading", background=[("active",SURFACE2)],
              foreground=[("active",TEXT)])
        # Remove native light borders/focus rings around the two data tables.
        s.layout("CopyPro.Treeview", [("Treeview.treearea", {"sticky":"nswe"})])
        s.layout("CopyPro.Treeview.Heading", [("Treeheading.cell", {"sticky":"nswe", "children":[
            ("Treeheading.padding", {"sticky":"nswe", "children":[
                ("Treeheading.image", {"side":"right", "sticky":""}),
                ("Treeheading.text", {"sticky":"we"})
            ]})
        ]})])

        # Keep combobox pop-down lists consistent with the app palette.
        self.option_add("*TCombobox*Listbox.background", SURFACE2)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", TEXT)

    def _check_for_updates(self, manual=True):
        if manual:
            self.configure(cursor="watch")

        def finished(release, error):
            def show_result():
                self.configure(cursor="")
                if error:
                    if manual:
                        messagebox.showerror(
                            "Atnaujinimo klaida",
                            f"Nepavyko patikrinti atnaujinimų.\n\n{error}",
                            parent=self,
                        )
                    return
                if not release:
                    if manual:
                        messagebox.showinfo(
                            "Atnaujinimai",
                            f"Naudojate naujausią versiją ({current_version()}).",
                            parent=self,
                        )
                    return

                notes = release.get("notes", "").strip()
                message = (
                    f"Yra nauja CopyPro Tools versija.\n\n"
                    f"Dabartinė: {current_version()}\n"
                    f"Nauja: {release['version']}"
                )
                if notes:
                    message += "\n\n" + notes[:1000]

                if messagebox.askyesno(
                    "CopyPro Tools atnaujinimas",
                    message + "\n\nAtnaujinti dabar?",
                    parent=self,
                ):
                    try:
                        launch_updater(release)
                        self.after(250, self.destroy)
                    except Exception as ex:
                        messagebox.showerror(
                            "Atnaujinimo klaida",
                            str(ex),
                            parent=self,
                        )
            self.after(0, show_result)

        check_for_update_async(finished)

    def _diagnostic_snapshot(self):
        layout=self.tabs.get("layout") if hasattr(self,"tabs") else None
        scan=self.tabs.get("scan") if hasattr(self,"tabs") else None
        cache=getattr(layout,"_preview_item_cache",None)
        tk_images=0
        try:tk_images=len(self.tk.call("image","names"))
        except Exception:pass
        return {
            "ram":process_memory_mb(),
            "cache_mb":getattr(cache,"current_bytes",0)/(1024*1024),
            "cache_items":len(cache) if cache is not None else 0,
            "layout_sources":len(getattr(layout,"files",[])) if layout else 0,
            "preview_images":len(getattr(layout,"preview_images",[])) if layout else 0,
            "preview_job":bool(getattr(layout,"_preview_job",None)) if layout else False,
            "scan_pages":len(getattr(scan,"pages",[])) if scan else 0,
            "tk_images":tk_images,
        }

    def _clear_all_preview_caches(self):
        for tab in self.tabs.values():
            if hasattr(tab,"clear_memory_cache"):
                try:tab.clear_memory_cache()
                except Exception:pass
        gc.collect()

    def _open_diagnostics(self):
        popup=tk.Toplevel(self); popup.title("CopyPro diagnostics"); popup.configure(bg=BG); popup.geometry("470x350"); popup.resizable(False,False)
        lbl(popup,"MEMORY & PREVIEW DIAGNOSTICS",11,TEXT,True,bg=BG).pack(anchor="w",padx=18,pady=(16,8))
        output=tk.Label(popup,bg=SURFACE,fg=TEXT,font=("Consolas",10),justify="left",anchor="nw",padx=14,pady=12); output.pack(fill="both",expand=True,padx=18,pady=(0,10))
        def refresh():
            if not popup.winfo_exists():return
            data=self._diagnostic_snapshot()
            output.config(text=(
                f"Current RAM:            {data['ram']:.1f} MB\n"
                f"Print preview cache:    {data['cache_mb']:.1f} MB / 128 MB\n"
                f"Cached preview items:   {data['cache_items']}\n"
                f"Loaded layout sources:  {data['layout_sources']}\n"
                f"Displayed sheet images: {data['preview_images']}\n"
                f"Preview refresh queued: {'yes' if data['preview_job'] else 'no'}\n"
                f"Scan pages loaded:       {data['scan_pages']}\n"
                f"Tk image objects:        {data['tk_images']}\n"
            ))
            popup.after(1000,refresh)
        footer=tk.Frame(popup,bg=BG); footer.pack(fill="x",padx=18,pady=(0,14))
        styled_btn(footer,"Clear preview caches",self._clear_all_preview_caches,style="secondary").pack(side="left")
        styled_btn(footer,"Close",popup.destroy).pack(side="right"); refresh()

    def _open_data_settings(self):
        popup=tk.Toplevel(self); popup.title("Duomenu failu nustatymai"); popup.configure(bg=BG); popup.geometry("690x300"); popup.grab_set()
        paths=_editable_paths(); vars_={k:tk.StringVar(value=v) for k,v in paths.items()}
        labels=(("codes_file","Prekiu ir paslaugu kodai"),("paper_sizes_file","Popieriaus dydziai"))
        lbl(popup,"DUOMENU FAILAI",11,TEXT,True,bg=BG).pack(anchor="w",padx=18,pady=(16,10))
        body=tk.Frame(popup,bg=BG); body.pack(fill="both",expand=True,padx=18)
        def browse(key):
            p=filedialog.askopenfilename(parent=popup,title="Pasirinkti Excel failą",filetypes=[("Excel","*.xlsx"),("CSV (senas formatas)","*.csv"),("Visi failai","*.*")])
            if p: vars_[key].set(p)
        for r,(key,title) in enumerate(labels):
            lbl(body,title,9,MUTED,bg=BG).grid(row=r*2,column=0,columnspan=2,sticky="w",pady=(4,2))
            e=tk.Entry(body,textvariable=vars_[key],bg=SURFACE2,fg=TEXT,insertbackground=TEXT,relief="flat",font=("Segoe UI",9))
            e.grid(row=r*2+1,column=0,sticky="ew",ipady=6)
            styled_btn(body,"Pasirinkti…",lambda k=key:browse(k),style="secondary",padx=10,pady=5).grid(row=r*2+1,column=1,padx=(8,0))
        body.grid_columnconfigure(0,weight=1)
        def save_reload():
            data=load_app_settings(); data.update({k:v.get().strip() for k,v in vars_.items()}); save_app_settings(data)
            try:
                result=reload_editable_data()
                for tab in self.tabs.values():
                    if hasattr(tab,"on_show"): tab.on_show()
                messagebox.showinfo(
                    "Išsaugota",
                    "Duomenys įkelti tik iš šių Excel failų:\n\n"
                    f"Kodai: {result['codes_file']}\n"
                    f"Popieriaus dydžiai: {result['paper_sizes_file']}\n\n"
                    f"Kodų: {result['code_count']}\n"
                    f"Plataus formato eilučių: {result['wide_format_count']}\n"
                    f"Popieriaus dydžių: {result['paper_size_count']}",
                    parent=popup,
                )
                popup.destroy()
            except Exception as ex: messagebox.showerror("Klaida",str(ex),parent=popup)
        footer=tk.Frame(popup,bg=BG); footer.pack(fill="x",padx=18,pady=14)
        styled_btn(footer,"Atidaryti AppData aplanka",lambda:subprocess.Popen(["explorer",COPYPRO_DATA_DIR]) if sys.platform.startswith("win") else None,style="secondary").pack(side="left")
        styled_btn(footer,"Tikrinti atnaujinimus",lambda:self._check_for_updates(manual=True),style="secondary").pack(side="left",padx=8)
        styled_btn(footer,"Diagnostika",self._open_diagnostics,style="secondary").pack(side="left")
        lbl(footer,f"v{current_version()}",8,MUTED,bg=BG).pack(side="left",padx=4)
        styled_btn(footer,"Išsaugoti",save_reload,style="success").pack(side="right")

    def _build(self):
        tab_bar = tk.Frame(self, bg=SURFACE2, height=38)
        tab_bar.pack(fill="x"); tab_bar.pack_propagate(False)
        tk.Button(tab_bar,text="⚙",command=self._open_data_settings,bg=SURFACE2,fg=TEXT,relief="flat",bd=0,
            font=("Segoe UI",12,"bold"),padx=12,activebackground=BG,activeforeground=TEXT,cursor="hand2").pack(side="right",fill="y")

        content = tk.Frame(self, bg=BG)
        content.pack(fill="both", expand=True)

        tab_defs = [
            ("resizer",   "📐  Bulk Resizer",       ResizerTab),
            ("converter", "🔄  File Converter",     ConverterTab),
            ("layout",    "🗂  Print Layout",       PrintLayoutTab),
            ("polaroid",  "▧  Polaroid Maker",     PolaroidMakerTab),
            ("scan",      "▣  Scan Crop",          ScanCroppingTab),
            ("cutline",   "✂  Sticker Outline",    StickerCutlineTab),
            ("pdf",       "📄  PDF Optimiser",      PdfToolsTab),
            ("pages",     "📊  Page Counter",       PageCounterTab),
            ("organiser", "🗃  Batch Organiser",    BatchOrganiserTab),
            ("counter",   "🖨  Print Counter",      PrintCounterTab),
        ]
        self.tabs = {}; self.tab_btns = {}
        for key,_,cls in tab_defs:
            self.tabs[key] = cls(content)

        # Make the complete visible area of every file-based tab a drop target.
        drop_rules = {
            "resizer": IMAGE_EXTS | {".pdf"},
            "converter": None,
            "layout": IMAGE_EXTS | {".pdf"},
            "polaroid": IMAGE_EXTS | {".pdf"},
            "scan": {".pdf"},
            "cutline": IMAGE_EXTS | {".pdf"},
            "pdf": {".pdf"},
            "pages": {".pdf"},
            "organiser": {".pdf"},
        }
        for key, extensions in drop_rules.items():
            tab = self.tabs.get(key)
            if tab is not None and hasattr(tab, "setup_drop_everywhere") and hasattr(tab, "_drop_files"):
                tab.setup_drop_everywhere(tab, tab._drop_files, extensions)

        def switch(k):
            for kk,t in self.tabs.items():
                t.pack_forget()
                self.tab_btns[kk].config(bg=SURFACE2,fg=MUTED)
            self.tabs[k].pack(fill="both",expand=True)
            self.tab_btns[k].config(bg=BG,fg=TEXT)
            if hasattr(self.tabs[k], "on_show"):
                self.tabs[k].on_show()

        for key,label_text,_ in tab_defs:
            btn = tk.Button(tab_bar, text=label_text,
                bg=SURFACE2, fg=MUTED, relief="flat", bd=0,
                font=("Segoe UI",8,"bold"), padx=8,
                activebackground=BG, activeforeground=TEXT,
                cursor="hand2", command=lambda k=key: switch(k))
            btn.pack(side="left",fill="y")
            self.tab_btns[key] = btn

        switch("resizer")

if __name__ == "__main__":
    if ensure_installed_in_appdata():
        raise SystemExit(0)
    app = CopyProApp()
    app.mainloop()
