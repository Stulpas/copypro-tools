import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os, sys, threading, subprocess, io, math, tempfile, shutil, unicodedata, difflib, json, csv
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

PAPER_SIZES_MM = {
    "A0": (841,1189), "A1": (594,841), "A2": (420,594),
    "A3": (297,420), "A4": (210,297), "A5": (148,210),
    "A6": (105,148), "A7": (74,105),
    "SRA3": (320,450), "SRA4": (225,320),
    "10×15 cm (102×152)": (102,152),
    "21×15 cm (210×152)": (210,152),
    "Square 10×10": (100,100), "Square 15×15": (150,150),
}
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
DEFAULT_CODES_FILE = os.path.join(COPYPRO_DATA_DIR, "copypro_kodai.csv")
DEFAULT_PAPER_SIZES_FILE = os.path.join(COPYPRO_DATA_DIR, "copypro_popieriaus_dydziai.csv")
DEFAULT_WIDE_FORMAT_FILE = DEFAULT_CODES_FILE

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
ALL_EXTS   = IMAGE_EXTS | {".pdf",".svg",".eps",".ai",".ps"}

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
    """Call setup_drop(widget, callback, extensions) to enable DnD on a widget."""
    def setup_drop(self, widget, callback, extensions=None):
        if not HAS_DND:
            return
        try:
            widget.drop_target_register(dnd.DND_FILES)
            widget.dnd_bind("<<DropEnter>>",  lambda e: self._drop_enter(widget))
            widget.dnd_bind("<<DropLeave>>",  lambda e: self._drop_leave(widget))
            widget.dnd_bind("<<Drop>>",       lambda e: self._on_drop(e, widget, callback, extensions))
        except Exception:
            pass

    def _drop_enter(self, w):
        try: w.config(bg=DROP_HL)
        except: pass
    def _drop_leave(self, w):
        try: w.config(bg=SURFACE2)
        except: pass
    def _on_drop(self, event, widget, callback, extensions):
        try: widget.config(bg=SURFACE2)
        except: pass
        paths = parse_dropped_paths(event.data)
        filtered = []
        for p in paths:
            if os.path.isfile(p):
                if extensions is None or os.path.splitext(p)[1].lower() in extensions:
                    filtered.append(p)
            elif os.path.isdir(p):
                for f in os.listdir(p):
                    fp = os.path.join(p, f)
                    if os.path.isfile(fp):
                        ext = os.path.splitext(fp)[1].lower()
                        if extensions is None or ext in extensions:
                            filtered.append(fp)
        if filtered:
            callback(filtered)

# ── Crop Canvas ───────────────────────────────────────────────────────────────
PREVIEW_MAX_DIM = 900  # cap preview image's long side in px — keeps UI fast
HANDLE_SIZE = 9         # corner handle hit-box half-size, in canvas px
MIN_CROP_FRAC = 0.15    # don't allow shrinking the crop box below 15% of the fitted size

class CropCanvas(tk.Canvas):
    def __init__(self, parent, img_path, target_w, target_h, size=(300,300), **kw):
        super().__init__(parent, width=size[0], height=size[1],
                         bg=SURFACE2, highlightthickness=0, **kw)
        self.target_w, self.target_h = target_w, target_h
        self.canvas_w, self.canvas_h = size
        self.src_path = img_path
        self.total_rotation = 0  # cumulative rotation (deg, multiple of 90) vs. EXIF-corrected original
        self._load(img_path)
        self._init_crop()
        self._draw()
        self.bind("<ButtonPress-1>",   self._press)
        self.bind("<B1-Motion>",       self._drag)
        self.bind("<ButtonRelease-1>", lambda e: setattr(self,'_ds',None))
        self.bind("<Motion>",          self._on_hover)
        self._ds = None
        self._drag_mode = None  # "move" or one of "nw","ne","sw","se"

    @property
    def rotated(self):
        """Kept for backward compatibility: True if a 90/270 rotation is in effect."""
        return self.total_rotation in (90, 270)

    def _load(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            imgs = open_pdf_pages(path, dpi=72)
            full = imgs[0].convert("RGB") if imgs else Image.new("RGB",(100,100))
        else:
            raw = Image.open(path)
            # Honour EXIF rotation so phone photos aren't sideways
            raw = ImageOps.exif_transpose(raw)
            full = raw.convert("RGB")

        iw, ih = full.size
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
        self.tk_img = ImageTk.PhotoImage(self.pil_img.resize((dw,dh), Image.LANCZOS))

    def rotate_manual(self, delta):
        """Rotate the preview (and the eventual export) by delta degrees (±90).
        Lossless since it's always a multiple of 90."""
        self.pil_img = self.pil_img.rotate(delta, expand=True)
        self.total_rotation = (self.total_rotation + delta) % 360
        self._fit_to_canvas()
        self._init_crop()
        self._draw()

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
def add_bleed_and_marks(img, bleed_mm, dpi, crop_marks=True, mark_len_mm=5, mark_gap_mm=1):
    """
    Extend image with mirror-bleed on all sides, optionally add crop marks.
    Returns new PIL Image.
    """
    bp   = mm_to_px(bleed_mm, dpi)
    iw, ih = img.size
    new_w = iw + bp*2
    new_h = ih + bp*2

    # Create canvas (white if crop marks, else transparent-safe)
    extra = mm_to_px(mark_len_mm + mark_gap_mm + 2, dpi) if crop_marks else 0
    canvas_w = new_w + extra*2
    canvas_h = new_h + extra*2
    canvas = Image.new("RGB", (canvas_w, canvas_h), (255,255,255))

    # Mirror bleed strips
    left_strip  = img.crop((0, 0, bp, ih)).transpose(Image.FLIP_LEFT_RIGHT)
    right_strip = img.crop((iw-bp, 0, iw, ih)).transpose(Image.FLIP_LEFT_RIGHT)
    top_strip   = img.crop((0, 0, iw, bp)).transpose(Image.FLIP_TOP_BOTTOM)
    bot_strip   = img.crop((0, ih-bp, iw, ih)).transpose(Image.FLIP_TOP_BOTTOM)

    # Corners (mirror of corner squares)
    tl = img.crop((0,0,bp,bp)).transpose(Image.ROTATE_180)
    tr = img.crop((iw-bp,0,iw,bp)).transpose(Image.ROTATE_180)
    bl = img.crop((0,ih-bp,bp,ih)).transpose(Image.ROTATE_180)
    br = img.crop((iw-bp,ih-bp,iw,ih)).transpose(Image.ROTATE_180)

    ox = extra + bp  # origin x of the original image on canvas
    oy = extra + bp

    canvas.paste(img,          (ox, oy))
    canvas.paste(left_strip,   (ox-bp, oy))
    canvas.paste(right_strip,  (ox+iw, oy))
    canvas.paste(top_strip,    (ox, oy-bp))
    canvas.paste(bot_strip,    (ox, oy+ih))
    canvas.paste(tl,           (ox-bp, oy-bp))
    canvas.paste(tr,           (ox+iw, oy-bp))
    canvas.paste(bl,           (ox-bp, oy+ih))
    canvas.paste(br,           (ox+iw, oy+ih))

    if crop_marks:
        draw = ImageDraw.Draw(canvas)
        gap = mm_to_px(mark_gap_mm, dpi)
        mlen = mm_to_px(mark_len_mm, dpi)
        lw = max(1, mm_to_px(0.25, dpi))
        col = (0,0,0)

        # Corner positions on the full canvas: corners of the bleed box
        corners = [
            (ox-bp, oy-bp),  # TL
            (ox+iw+bp, oy-bp),  # TR
            (ox-bp, oy+ih+bp),  # BL
            (ox+iw+bp, oy+ih+bp),  # BR
        ]
        dirs = [(-1,-1), (1,-1), (-1,1), (1,1)]  # outward direction per corner

        for (cx,cy),(dx,dy) in zip(corners, dirs):
            # horizontal mark
            x0 = cx + dx*gap; x1 = cx + dx*(gap+mlen)
            draw.line([(x0,cy),(x1,cy)], fill=col, width=lw)
            # vertical mark
            y0 = cy + dy*gap; y1 = cy + dy*(gap+mlen)
            draw.line([(cx,y0),(cx,y1)], fill=col, width=lw)

    return canvas

def add_crop_marks_only(img, dpi, mark_len_mm=5, mark_gap_mm=1):
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

def apply_resizer_finishing(img, dpi, add_bleed, bleed_mm, crop_marks, mark_len_mm):
    if add_bleed and bleed_mm > 0:
        return add_bleed_and_marks(img, bleed_mm, dpi, crop_marks, mark_len_mm)
    if crop_marks:
        return add_crop_marks_only(img, dpi, mark_len_mm)
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
            font=("Segoe UI", 8, "bold"), anchor="w")
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

        # Bleed amount is shown only when bleed is enabled.
        self.bleed_opts = tk.Frame(left, bg=SURFACE)
        br = tk.Frame(self.bleed_opts, bg=SURFACE); br.pack(fill="x",padx=4,pady=1)
        lbl(br,"Bleed (mm)",8).pack(side="left",padx=(0,4))
        self.bleed_mm = entry(br, w=5); self.bleed_mm.insert(0,"3"); self.bleed_mm.pack(side="left")

        marks_row = tk.Frame(left, bg=SURFACE)
        marks_row.pack(fill="x", padx=16, pady=(1,2))
        self.marks_var = tk.BooleanVar(value=False)
        tk.Checkbutton(marks_row, text="Add crop marks", variable=self.marks_var,
            bg=SURFACE, fg=TEXT, selectcolor=SURFACE2, activebackground=SURFACE,
            font=("Segoe UI",9)).pack(side="left")
        lbl(marks_row,"Length",8,MUTED).pack(side="left",padx=(8,3))
        self.mark_len = entry(marks_row, w=4); self.mark_len.insert(0,"5"); self.mark_len.pack(side="left")
        lbl(marks_row,"mm",8,MUTED).pack(side="left",padx=(3,0))

        self.gray_var = tk.BooleanVar(value=False)
        tk.Checkbutton(left, text="Grayscale", variable=self.gray_var,
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

    def _toggle_bleed(self):
        if self.bleed_var.get():
            self.bleed_opts.pack(fill="x", padx=16, pady=(0,6), after=self.bleed_row)
        else:
            self.bleed_opts.pack_forget()

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

    def _do_export(self, dest_resolver):
        dpi = self.dpi_var.get()
        tw, th = self._get_target_px()
        fmt = self.fmt_var.get()
        quality = int(self.q_var.get())
        ext = fmt.lower().replace("jpeg","jpg")
        bleed = self.bleed_var.get()
        try: bleed_mm = float(self.bleed_mm.get())
        except: bleed_mm = 3.0
        marks = self.marks_var.get()
        try: mark_len = float(self.mark_len.get())
        except: mark_len = 5.0
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
                    cropped = img.crop((int(fl*iw), int(ft*ih), int(fr*iw), int(fb*ih))).resize((tw, th), Image.LANCZOS)
                    if self.gray_var.get(): cropped = ImageOps.grayscale(cropped).convert("RGB")
                    cropped = apply_resizer_finishing(cropped, dpi, bleed, bleed_mm, marks, mark_len)
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
                cropped = img.crop(box).resize((tw,th), Image.LANCZOS)
                if self.gray_var.get(): cropped = ImageOps.grayscale(cropped).convert("RGB")
                cropped = apply_resizer_finishing(cropped, dpi, bleed, bleed_mm, marks, mark_len)

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
        styled_btn(left,"🗑   Clear All",self._clear,style="secondary").pack(fill="x",padx=16,pady=3)
        sep(left)
        self.conv_btn = styled_btn(left,"🔄  Convert All",self._convert_all,style="success")
        self.conv_btn.pack(fill="x",padx=16,pady=3)
        self.status = lbl(left,"",8,MUTED); self.status.pack(padx=16,pady=6,anchor="w")

        # Right
        right = tk.Frame(self, bg=BG)
        right.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        self.drop_zone = tk.Label(right,
            text="⬇  Drop files here  (images, PDFs, SVGs…)",
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
                    if os.path.splitext(path)[1].lower() == ".pdf":
                        pages.extend(img.convert("RGB") for img in open_pdf_pages(path, dpi=rdpi))
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
                # ── PDF input → rasterise each page ──
                if src_ext == ".pdf":
                    pages = open_pdf_pages(path, dpi=rdpi)
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

class PrintCropCanvas(CropCanvas):
    def __init__(self,*args,**kwargs):
        self.preview_grayscale=False; self.preview_brightness=100
        super().__init__(*args,**kwargs)
    def _fit_to_canvas(self):
        iw,ih=self.pil_img.size; scale=min(self.canvas_w/iw,self.canvas_h/ih)
        dw,dh=max(1,int(iw*scale)),max(1,int(ih*scale))
        self.offset_x=(self.canvas_w-dw)//2; self.offset_y=(self.canvas_h-dh)//2
        self.disp_w,self.disp_h,self.scale=dw,dh,scale
        shown=self.pil_img
        if self.preview_grayscale: shown=ImageOps.grayscale(shown).convert("RGB")
        if self.preview_brightness!=100: shown=ImageEnhance.Brightness(shown).enhance(self.preview_brightness/100)
        self.tk_img=ImageTk.PhotoImage(shown.resize((dw,dh),Image.LANCZOS))
    def set_effects(self,gray,brightness):
        self.preview_grayscale=bool(gray); self.preview_brightness=int(brightness)
        self._fit_to_canvas(); self._draw()

class PrintLayoutTab(tk.Frame,DropMixin):
    def __init__(self,parent):
        super().__init__(parent,bg=BG)
        self.files=[]; self.canvases=[]; self.layouts=load_print_layouts(); self.custom_sizes=load_custom_sizes(); self._preview_job=None
        self.pdf_sources={}; self.display_names={}; self._temp_dir=tempfile.mkdtemp(prefix="copypro_print_")
        self._build(); self._load_layout("Default")
    def on_show(self):
        self.custom_sizes=load_custom_sizes(); self.paper_cb.config(values=self._paper_names())
    def _paper_names(self): return list(PAPER_SIZES_MM)+list(self.custom_sizes)
    def _build(self):
        left=tk.Frame(self,bg=SURFACE,width=285); left.pack(side="left",fill="y",padx=(0,1)); left.pack_propagate(False)
        lbl(left,"PRINT LAYOUT",8,MUTED,True).pack(pady=(12,4),padx=16,anchor="w")
        pr=tk.Frame(left,bg=SURFACE); pr.pack(fill="x",padx=16)
        self.layout_var=tk.StringVar(value="Default")
        self.layout_cb=ttk.Combobox(pr,textvariable=self.layout_var,values=list(self.layouts),state="readonly",width=20); self.layout_cb.pack(side="left",fill="x",expand=True)
        self.layout_cb.bind("<<ComboboxSelected>>",lambda e:self._load_layout(self.layout_var.get()))
        tk.Button(pr,text="Save",command=self._save_layout_dialog,bg=SURFACE2,fg=TEXT,relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="left",padx=(5,0))
        lbl(left,"Paper size").pack(padx=16,anchor="w",pady=(7,0))
        self.paper_var=tk.StringVar(value="A4"); self.paper_cb=ttk.Combobox(left,textvariable=self.paper_var,values=self._paper_names(),state="readonly")
        self.paper_cb.pack(fill="x",padx=16,pady=(2,3)); self.paper_cb.bind("<<ComboboxSelected>>",self._changed)
        rr=tk.Frame(left,bg=SURFACE); rr.pack(fill="x",padx=16)
        self.orientation_var=tk.StringVar(value="Portrait")
        for x in ("Portrait","Landscape"):
            tk.Radiobutton(rr,text=x,variable=self.orientation_var,value=x,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._changed).pack(side="left")
        self.dpi_var=tk.IntVar(value=300); ttk.Combobox(rr,textvariable=self.dpi_var,values=[150,300,600],state="readonly",width=6).pack(side="right"); lbl(rr,"DPI",8,MUTED).pack(side="right",padx=4)
        sep(left)
        grid_hdr=tk.Frame(left,bg=SURFACE); grid_hdr.pack(fill="x",padx=16)
        lbl(grid_hdr,"ITEM & GRID",8,MUTED,True).pack(side="left")
        tk.Button(grid_hdr,text="Auto layout…",command=self._open_layout_calculator,
            bg=SURFACE2,fg=TEXT,relief="flat",font=("Segoe UI",8,"bold"),
            cursor="hand2",padx=7,pady=2).pack(side="right")
        size_row=tk.Frame(left,bg=SURFACE); size_row.pack(fill="x",padx=16,pady=(3,0))
        self.item_w=tk.DoubleVar(value=90); self.item_h=tk.DoubleVar(value=50); self.cols_var=tk.IntVar(value=2); self.rows_var=tk.IntVar(value=5)
        for i,(t,v) in enumerate((("Width",self.item_w),("Height",self.item_h))):
            f=tk.Frame(size_row,bg=SURFACE); f.pack(side="left",fill="x",expand=True,padx=(0 if i==0 else 4,0))
            lbl(f,t+" (mm)",7,MUTED).pack(anchor="w"); e=entry(f,v,7); e.pack(fill="x"); e.bind("<KeyRelease>",self._changed)
        tk.Button(size_row,text="Presets…",command=self._open_item_size_presets,bg=SURFACE2,fg=TEXT,
            relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2",padx=7,pady=4).pack(side="left",padx=(6,0),pady=(12,0))
        grid=tk.Frame(left,bg=SURFACE); grid.pack(fill="x",padx=16,pady=(4,0))
        for i,(t,v) in enumerate((("Columns",self.cols_var),("Rows",self.rows_var))):
            f=tk.Frame(grid,bg=SURFACE); f.grid(row=0,column=i,padx=(0 if i==0 else 4,0),sticky="ew"); grid.grid_columnconfigure(i,weight=1)
            lbl(f,t,7,MUTED).pack(anchor="w"); e=entry(f,v,7); e.pack(fill="x"); e.bind("<KeyRelease>",self._changed)
        mr=tk.Frame(left,bg=SURFACE); mr.pack(fill="x",padx=16,pady=(6,0)); self.mode_var=tk.StringVar(value="Repeat")
        for x in ("Repeat","Cut & stack","In order"):
            tk.Radiobutton(mr,text=x,variable=self.mode_var,value=x,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),command=self._mode_changed).pack(side="left")
        self.duplex_var=tk.BooleanVar(value=True); tk.Checkbutton(left,text="Double-sided (left bind)",variable=self.duplex_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",9),command=self._changed).pack(padx=16,anchor="w")
        sep(left); lbl(left,"SPACING & BLEED",8,MUTED,True).pack(padx=16,anchor="w")
        sr=tk.Frame(left,bg=SURFACE); sr.pack(fill="x",padx=16,pady=(3,0)); lbl(sr,"Trim spacing (mm)",8).pack(side="left")
        self.spacing_var=tk.DoubleVar(value=6); e=entry(sr,self.spacing_var,6); e.pack(side="right"); e.bind("<KeyRelease>",self._changed)
        br=tk.Frame(left,bg=SURFACE); br.pack(fill="x",padx=16,pady=(3,0)); self.bleed_var=tk.BooleanVar(value=True)
        tk.Checkbutton(br,text="Add bleed",variable=self.bleed_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",9),command=self._toggle_bleed).pack(side="left")
        self.bleed_mm=tk.DoubleVar(value=3); self.bleed_entry=entry(br,self.bleed_mm,6); self.bleed_entry.pack(side="right"); self.bleed_entry.bind("<KeyRelease>",self._changed); lbl(br,"mm",8,MUTED).pack(side="right",padx=3)
        cr=tk.Frame(left,bg=SURFACE); cr.pack(fill="x",padx=16,pady=(3,0))
        self.cut_marks_var=tk.BooleanVar(value=False)
        self.cut_labels_var=tk.BooleanVar(value=False)
        tk.Checkbutton(cr,text="Cut marks",variable=self.cut_marks_var,bg=SURFACE,fg=TEXT,
            selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",8),
            command=self._toggle_cut_marks).pack(side="left")
        self.cut_labels_chk=tk.Checkbutton(cr,text="Distance labels",variable=self.cut_labels_var,
            bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,
            font=("Segoe UI",8),command=self._changed)
        self.cut_labels_chk.pack(side="right")
        sep(left); lbl(left,"IMAGE ADJUSTMENTS",8,MUTED,True).pack(padx=16,anchor="w")
        self.gray_var=tk.BooleanVar(value=False); tk.Checkbutton(left,text="Grayscale",variable=self.gray_var,bg=SURFACE,fg=TEXT,selectcolor=SURFACE2,activebackground=SURFACE,font=("Segoe UI",9),command=self._effects_changed).pack(padx=16,anchor="w")
        self.brightness_var=tk.IntVar(value=100); self.brightness_lbl=lbl(left,"Brightness: 100%",8); self.brightness_lbl.pack(padx=16,anchor="w")
        ttk.Scale(left,from_=25,to=175,variable=self.brightness_var,orient="horizontal",command=self._brightness_changed).pack(fill="x",padx=16)
        self.fit_lbl=lbl(left,"",8,MUTED); self.fit_lbl.pack(padx=16,anchor="w",pady=(3,1))
        ar=tk.Frame(left,bg=SURFACE); ar.pack(fill="x",padx=16,pady=2)
        styled_btn(ar,"➕ Add",self._add_files).pack(side="left",fill="x",expand=True,padx=(0,3)); styled_btn(ar,"🗑 Clear",self._clear,style="secondary").pack(side="left",fill="x",expand=True,padx=(3,0))
        self.export_btn=styled_btn(left,"Export PDF",self._export_pdf,style="success"); self.export_btn.pack(fill="x",padx=16,pady=3)
        self.status=lbl(left,"",8,MUTED); self.status.pack(padx=16,anchor="w")
        right=tk.Frame(self,bg=BG); right.pack(side="left",fill="both",expand=True)
        self.drop_zone=tk.Label(right,text="⬇  Drop images or multi-page PDFs here",bg=SURFACE2,fg=MUTED,font=("Segoe UI",10),pady=9,cursor="hand2"); self.drop_zone.pack(fill="x",padx=12,pady=(8,0)); self.drop_zone.bind("<Button-1>",lambda e:self._add_files()); self.setup_drop(self.drop_zone,self._drop_files,IMAGE_EXTS | {".pdf"})
        body=tk.Frame(right,bg=BG); body.pack(fill="both",expand=True,padx=12,pady=8)
        self.sheet_preview=tk.Canvas(body,bg=SURFACE2,highlightthickness=0,width=250); self.sheet_preview.pack(side="left",fill="y",padx=(0,8)); self.sheet_preview.bind("<Configure>",lambda e:self._schedule_preview())
        self.cards_canvas=tk.Canvas(body,bg=BG,highlightthickness=0); self.cards_canvas.pack(side="left",fill="both",expand=True)
        self.cards_frame=tk.Frame(self.cards_canvas,bg=BG); self.cards_win=self.cards_canvas.create_window((0,0),window=self.cards_frame,anchor="nw")
        self.cards_frame.bind("<Configure>",lambda e:self.cards_canvas.configure(scrollregion=self.cards_canvas.bbox("all"))); self.cards_canvas.bind("<Configure>",lambda e:self.cards_canvas.itemconfig(self.cards_win,width=e.width))
        self.cards_canvas.bind("<Enter>",lambda e:self.cards_canvas.bind_all("<MouseWheel>",lambda ev:self.cards_canvas.yview_scroll(int(-ev.delta/120),"units"))); self.cards_canvas.bind("<Leave>",lambda e:self.cards_canvas.unbind_all("<MouseWheel>")); self.setup_drop(self.cards_canvas,self._drop_files,IMAGE_EXTS | {".pdf"})
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
        def apply_size(name,w,h): self.item_w.set(w); self.item_h.set(h); self._rebuild(); popup.destroy()
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
            self.item_w.set(state["w"]); self.item_h.set(state["h"]); self.cols_var.set(state["cols"]); self.rows_var.set(state["rows"])
            self.bleed_mm.set(state["bleed"]); self.bleed_var.set(state["bleed"]>0); self.spacing_var.set(state["gap"]); self._toggle_bleed(); self._rebuild(); popup.destroy()
        actions=tk.Frame(popup,bg=BG); actions.pack(pady=4)
        styled_btn(actions,"Calculate",calculate).pack(side="left",padx=4); styled_btn(actions,"Apply layout",apply_result,style="success").pack(side="left",padx=4); styled_btn(actions,"Cancel",popup.destroy,style="secondary").pack(side="left",padx=4)
        calculate()
    def _brightness_changed(self,v): self.brightness_lbl.config(text=f"Brightness: {int(float(v))}%"); self._effects_changed()
    def _effects_changed(self):
        for _,c in self.canvases:c.set_effects(self.gray_var.get(),self.brightness_var.get())
        self._schedule_preview()
    def _schedule_preview(self):
        if self._preview_job:
            try:self.after_cancel(self._preview_job)
            except:pass
        self._preview_job=self.after(80,self._draw_preview)
    def _draw_preview(self):
        self._preview_job=None; c=self.sheet_preview; c.delete("all")
        try:pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,gw,gh=self._metrics()
        except:return
        cw=max(140,c.winfo_width()); ch=max(200,c.winfo_height()); sc=min((cw-24)/pw,(ch-40)/ph); ox=(cw-pw*sc)/2; oy=(ch-ph*sc)/2
        c.create_rectangle(ox,oy,ox+pw*sc,oy+ph*sc,fill="white",outline=BORDER); fits=gw<=pw+1e-6 and gh<=ph+1e-6; self.fit_lbl.config(text=(f"Centered: {gw:.1f} × {gh:.1f} mm" if fits else f"Does not fit: {gw:.1f} × {gh:.1f} mm"),fg=MUTED if fits else DANGER)
        for r in range(rows):
            for col in range(cols):
                tx=x0+col*(iw+gap); ty=y0+r*(ih+gap); bx=tx-bleed; by=ty-bleed
                c.create_rectangle(ox+bx*sc,oy+by*sc,ox+(bx+iw+2*bleed)*sc,oy+(by+ih+2*bleed)*sc,fill="#D8D8E8",outline=ACCENT if bleed else BORDER)
                c.create_rectangle(ox+tx*sc,oy+ty*sc,ox+(tx+iw)*sc,oy+(ty+ih)*sc,outline="#111")
        if self.cut_marks_var.get():
            xs=sorted(set([x0+cidx*(iw+gap) for cidx in range(cols)]+[x0+cidx*(iw+gap)+iw for cidx in range(cols)]))
            ys=sorted(set([y0+ridx*(ih+gap) for ridx in range(rows)]+[y0+ridx*(ih+gap)+ih for ridx in range(rows)]))
            safe_mm=4.0; desired_mm=5.0; artwork_gap_mm=1.0
            top_limit=max(safe_mm, y0-bleed-artwork_gap_mm)
            bottom_limit=min(ph-safe_mm, y0+gh+artwork_gap_mm)
            left_limit=max(safe_mm, x0-bleed-artwork_gap_mm)
            right_limit=min(pw-safe_mm, x0+gw+artwork_gap_mm)
            top_len=max(0,min(desired_mm,top_limit-safe_mm)); bottom_len=max(0,min(desired_mm,ph-safe_mm-bottom_limit))
            left_len=max(0,min(desired_mm,left_limit-safe_mm)); right_len=max(0,min(desired_mm,pw-safe_mm-right_limit))
            label_font=("Segoe UI",9,"bold"); label_gap=max(3,1.2*sc)
            for x in xs:
                xx=ox+x*sc
                if top_len>0:
                    y1=oy+safe_mm*sc; y2=y1+top_len*sc; c.create_line(xx,y1,xx,y2,fill="#111")
                    if self.cut_labels_var.get(): c.create_text(xx+label_gap,(y1+y2)/2,text=f"{min(x,pw-x):.1f}",font=label_font,fill="#111",angle=90)
                if bottom_len>0:
                    y2=oy+(ph-safe_mm)*sc; y1=y2-bottom_len*sc; c.create_line(xx,y1,xx,y2,fill="#111")
                    if self.cut_labels_var.get(): c.create_text(xx+label_gap,(y1+y2)/2,text=f"{min(x,pw-x):.1f}",font=label_font,fill="#111",angle=90)
            for y in ys:
                yy=oy+y*sc
                if left_len>0:
                    x1=ox+safe_mm*sc; x2=x1+left_len*sc; c.create_line(x1,yy,x2,yy,fill="#111")
                    if self.cut_labels_var.get(): c.create_text((x1+x2)/2,yy-label_gap,text=f"{min(y,ph-y):.1f}",font=label_font,fill="#111")
                if right_len>0:
                    x2=ox+(pw-safe_mm)*sc; x1=x2-right_len*sc; c.create_line(x1,yy,x2,yy,fill="#111")
                    if self.cut_labels_var.get(): c.create_text((x1+x2)/2,yy-label_gap,text=f"{min(y,ph-y):.1f}",font=label_font,fill="#111")
    def _show_empty(self):
        for w in self.cards_frame.winfo_children():w.destroy()
        tk.Label(self.cards_frame,text="No images loaded\n\nDrop images here or click Add",bg=BG,fg=MUTED,font=("Segoe UI",12)).pack(pady=50); self._schedule_preview()
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
                        if key in self.files: continue
                        page=doc[page_no]; pix=page.get_pixmap(matrix=fitz.Matrix(100/72,100/72),alpha=False)
                        img=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                        preview=os.path.join(self._temp_dir,f"pdf_{abs(hash(key))}_{page_no+1}.png"); img.save(preview,"PNG")
                        self.files.append(preview); self.pdf_sources[preview]=(p,page_no); self.display_names[preview]=f"{os.path.basename(p)} — page {page_no+1}"; changed=True
                    doc.close()
                except Exception as ex: messagebox.showerror("PDF import failed",f"{os.path.basename(p)}\n{ex}")
            elif ext in IMAGE_EXTS and p not in self.files:
                self.files.append(p); self.display_names[p]=os.path.basename(p); changed=True
        if changed:self._rebuild()
    def _clear(self):
        self.files.clear(); self.canvases.clear(); self.pdf_sources.clear(); self.display_names.clear(); self._show_empty()
    def _remove(self,path):
        if path in self.files:self.files.remove(path)
        self.pdf_sources.pop(path,None); self.display_names.pop(path,None); self._rebuild()
    def _rebuild(self):
        for w in self.cards_frame.winfo_children():w.destroy()
        self.canvases.clear()
        if not self.files:self._show_empty();return
        dpi=int(self.dpi_var.get()); tw=mm_to_px(max(.1,self._num(self.item_w,90)),dpi); th=mm_to_px(max(.1,self._num(self.item_h,50)),dpi); avail=max(400,self.cards_canvas.winfo_width()); ncols=3 if avail>700 else 2; thumb=max(150,min(230,avail//ncols-28))
        for i,path in enumerate(self.files):
            r,col=divmod(i,ncols); cell=tk.Frame(self.cards_frame,bg=SURFACE2,padx=6,pady=6); cell.grid(row=r,column=col,padx=6,pady=6,sticky="n")
            try:
                cc=PrintCropCanvas(cell,path,tw,th,size=(thumb,thumb)); cc.pack(); cc.set_effects(self.gray_var.get(),self.brightness_var.get()); self.canvases.append((path,cc))
                tk.Label(cell,text=self.display_names.get(path,os.path.basename(path))[:36],bg=SURFACE2,fg=MUTED,font=("Segoe UI",8)).pack(); acts=tk.Frame(cell,bg=SURFACE2); acts.pack(fill="x",pady=(3,0))
                for t,cmd in (("⟲",lambda c=cc:c.rotate_manual(90)),("⟳",lambda c=cc:c.rotate_manual(-90)),("✕",lambda p=path:self._remove(p))): tk.Button(acts,text=t,command=cmd,bg=SURFACE,fg=TEXT,relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="left",fill="x",expand=True,padx=1)
            except Exception as ex: tk.Label(cell,text=str(ex)[:45],bg=SURFACE2,fg=DANGER).pack()
        self._schedule_preview()
    def _current_layout(self):
        return {"paper":self.paper_var.get(),"orientation":self.orientation_var.get(),"item_w":self._num(self.item_w,90),"item_h":self._num(self.item_h,50),"cols":int(self._num(self.cols_var,2)),"rows":int(self._num(self.rows_var,5)),"mode":self.mode_var.get(),"duplex":self.duplex_var.get(),"bleed":self.bleed_var.get(),"bleed_mm":self._num(self.bleed_mm,3),"spacing":self._num(self.spacing_var,6),"grayscale":self.gray_var.get(),"brightness":int(self.brightness_var.get()),"dpi":int(self.dpi_var.get()),"cut_marks":self.cut_marks_var.get(),"cut_labels":self.cut_labels_var.get()}
    def _load_layout(self,name):
        d=self.layouts.get(name)
        if not d:return
        self.layout_var.set(name); self.paper_var.set(d.get("paper","A4")); self.orientation_var.set(d.get("orientation","Portrait")); self.item_w.set(d.get("item_w",90)); self.item_h.set(d.get("item_h",50)); self.cols_var.set(d.get("cols",2)); self.rows_var.set(d.get("rows",5)); self.mode_var.set(d.get("mode","Repeat")); self.duplex_var.set(d.get("duplex",False)); self.bleed_var.set(d.get("bleed",False)); self.bleed_mm.set(d.get("bleed_mm",3)); self.spacing_var.set(d.get("spacing",0)); self.gray_var.set(d.get("grayscale",False)); self.brightness_var.set(d.get("brightness",100)); self.dpi_var.set(d.get("dpi",300)); self.cut_marks_var.set(d.get("cut_marks",False)); self.cut_labels_var.set(d.get("cut_labels",False)); self._toggle_cut_marks(); self.brightness_lbl.config(text=f"Brightness: {self.brightness_var.get()}%"); self._toggle_bleed(); self._mode_changed()
    def _save_layout_dialog(self):
        name=simpledialog.askstring("Save layout","Layout name:",initialvalue=self.layout_var.get(),parent=self)
        if not name:return
        name=name.strip(); self.layouts[name]=self._current_layout(); save_print_layouts(self.layouts); self.layout_cb.config(values=list(self.layouts)); self.layout_var.set(name); self.status.config(text=f"Saved layout: {name}",fg=SUCCESS)
    def _processed(self,path,cc,target_px=None):
        if path in self.pdf_sources:
            # PDFs are resolution-independent when they contain vector artwork. Render
            # each page specifically for its final placed size, with 2x supersampling,
            # instead of reusing the low-resolution preview or a fixed raster DPI.
            pdf_path,page_no=self.pdf_sources[path]
            doc=fitz.open(pdf_path)
            page=doc[page_no]
            l,t,r,b=cc.get_crop_fractions()
            frac_w=max(0.001,r-l); frac_h=max(0.001,b-t)
            page_w,page_h=page.rect.width,page.rect.height
            if cc.total_rotation in (90,270):
                page_w,page_h=page_h,page_w
            if target_px:
                target_w,target_h=target_px
                scale=max((target_w*2)/(page_w*frac_w),(target_h*2)/(page_h*frac_h))
                render_dpi=max(int(self.dpi_var.get()),min(1200,int(math.ceil(scale*72))))
            else:
                render_dpi=max(600,int(self.dpi_var.get())*2)
            pix=page.get_pixmap(matrix=fitz.Matrix(render_dpi/72,render_dpi/72),alpha=False)
            raw=Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
            doc.close()
        else:
            raw=ImageOps.exif_transpose(Image.open(path)).convert("RGB")
        if cc.total_rotation:raw=raw.rotate(cc.total_rotation,expand=True)
        iw,ih=raw.size; l,t,r,b=cc.get_crop_fractions(); raw=raw.crop((int(l*iw),int(t*ih),int(r*iw),int(b*ih)))
        if self.gray_var.get():raw=ImageOps.grayscale(raw).convert("RGB")
        if self.brightness_var.get()!=100:raw=ImageEnhance.Brightness(raw).enhance(self.brightness_var.get()/100)
        return raw
    def _item(self,path,cc,dpi,iw,ih,bleed):
        target=(mm_to_px(iw,dpi),mm_to_px(ih,dpi))
        img=self._processed(path,cc,target_px=target).resize(target,Image.LANCZOS)
        return add_bleed_and_marks(img,bleed,dpi,False) if bleed>0 else img
    def _sheet(self,assignments,back=False):
        pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,gw,gh=self._metrics(); dpi=int(self.dpi_var.get()); sheet=Image.new("RGB",(mm_to_px(pw,dpi),mm_to_px(ph,dpi)),"white"); by={p:c for p,c in self.canvases}
        for slot,path in enumerate(assignments):
            if not path or path not in by:continue
            row=slot//cols; col=slot%cols
            if back:col=cols-1-col
            item=self._item(path,by[path],dpi,iw,ih,bleed); tx=x0+col*(iw+gap)-bleed; ty=y0+row*(ih+gap)-bleed; sheet.paste(item,(mm_to_px(tx,dpi),mm_to_px(ty,dpi)))
        if self.cut_marks_var.get(): self._draw_cut_guides(sheet,pw,ph,iw,ih,cols,rows,gap,bleed,x0,y0,dpi)
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
    def _pages(self):
        slots=max(1,int(self.cols_var.get())*int(self.rows_var.get()))
        mode=self.mode_var.get(); duplex=self.duplex_var.get(); pages=[]
        if mode=="Repeat":
            step=2 if duplex else 1
            for i in range(0,len(self.files),step):
                pages.append(self._sheet([self.files[i]]*slots))
                if duplex: pages.append(self._sheet([self.files[i+1] if i+1<len(self.files) else None]*slots,True))
        elif mode=="Cut & stack":
            # Cut-and-stack is duplex by default: each consecutive pair is front/back.
            pairs=[self.files[i:i+2] for i in range(0,len(self.files),2)]
            ns=math.ceil(len(pairs)/slots)
            for p in range(ns):
                front=[]; back=[]
                for s in range(slots):
                    idx=s*ns+p; pair=pairs[idx] if idx<len(pairs) else []
                    front.append(pair[0] if pair else None)
                    back.append(pair[1] if len(pair)>1 else None)
                pages.append(self._sheet(front))
                pages.append(self._sheet(back,True))
        else:  # In order: simple one-sided placement in queue order.
            for i in range(0,len(self.files),slots):
                chunk=self.files[i:i+slots]+[None]*max(0,slots-len(self.files[i:i+slots]))
                pages.append(self._sheet(chunk))
        return pages
    def _export_pdf(self):
        if not self.canvases:messagebox.showwarning("Nothing to export","Add images first.");return
        pw,ph,*rest=self._metrics(); gw,gh=rest[-2:]
        if gw>pw+1e-6 or gh>ph+1e-6:messagebox.showerror("Layout does not fit","The centered grid is larger than the selected paper.");return
        out=filedialog.asksaveasfilename(title="Export print layout",defaultextension=".pdf",filetypes=[("PDF","*.pdf")],initialfile=(str(self.paper_var.get()).replace(" ","_")+f"_{self.item_w.get():g}x{self.item_h.get():g}.pdf"))
        if not out:return
        self.export_btn.config(state="disabled"); self.status.config(text="Exporting…",fg=MUTED); threading.Thread(target=self._do_export,args=(out,),daemon=True).start()
    def _do_export(self,out):
        try:
            pages=self._pages(); dpi=int(self.dpi_var.get()); pages[0].save(out,"PDF",save_all=True,append_images=pages[1:],resolution=dpi,quality=95,subsampling=0,optimize=True)
            self.after(0,lambda:self.status.config(text=f"Exported {len(pages)} PDF page(s)",fg=SUCCESS)); self.after(0,lambda:messagebox.showinfo("Done",f"Print layout exported to:\n{out}"))
        except Exception as ex:
            self.after(0,lambda e=ex:messagebox.showerror("Export failed",str(e))); self.after(0,lambda:self.status.config(text="Export failed",fg=DANGER))
        finally:self.after(0,lambda:self.export_btn.config(state="normal"))

    def destroy(self):
        try: shutil.rmtree(self._temp_dir,ignore_errors=True)
        except Exception: pass
        super().destroy()


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

# Unit prices in EUR, keyed by the final/resolved PLU code.
# These values were transcribed from the supplied store price sheets.
PLU_PRICES = {
    # Black-and-white copying / printing
    "10": 0.15, "11": 0.10, "12": 0.08, "13": 0.07,
    "20": 0.25, "21": 0.20, "22": 0.15, "23": 0.12,
    "30": 0.30, "31": 0.25, "32": 0.20, "33": 0.15,
    "40": 0.50, "41": 0.45, "42": 0.40, "43": 0.25,
    "50": 0.30, "51": 0.50, "52": 0.40, "53": 0.60,
    "60": 0.20,
    "110": 0.15, "111": 0.10, "112": 0.08, "113": 0.07,
    "120": 0.25, "121": 0.20, "122": 0.15, "123": 0.12,
    "130": 0.30, "131": 0.25, "132": 0.20, "133": 0.15,
    "140": 0.50, "141": 0.45, "142": 0.40, "143": 0.25,
    "310": 0.15, "311": 0.10, "312": 0.08, "313": 0.07,
    "320": 0.25, "321": 0.20, "322": 0.15, "323": 0.12,
    "330": 0.30, "331": 0.25, "332": 0.20, "333": 0.15,
    "340": 0.50, "341": 0.45, "342": 0.40, "343": 0.25,
    "450": 0.50, "451": 0.80, "650": 0.50, "651": 0.80,

    # Colour printing
    "510": 0.95, "511": 0.85, "512": 0.75, "513": 0.65,
    "520": 1.65, "521": 1.50, "522": 1.30, "523": 1.10,
    "530": 1.60, "531": 1.50, "532": 1.35, "533": 1.20,
    "540": 2.90, "541": 2.75, "542": 2.50, "543": 2.00,
    "710": 0.95, "711": 0.85, "712": 0.75, "713": 0.65,
    "720": 1.65, "721": 1.50, "722": 1.30, "723": 1.10,
    "730": 1.60, "731": 1.50, "732": 1.35, "733": 1.20,
    "740": 2.90, "741": 2.75, "742": 2.50, "743": 2.00,

    # Paper and speciality products
    "899": 2.00, "954": 0.70, "956": 0.80, "966": 3.50,
    "988": 1.00, "999": 2.00,
    "1090": 2.50, "1092": 2.50, "1095": 3.50, "1096": 3.20,
    "1184": 7.50, "1319": 2.00,
    "2022": 3.00, "2851": 3.10,
    "9502": 0.15, "9504": 5.50, "9512": 0.20,
    "9582": 1.00, "9592": 1.50,
    "9822": 0.50, "9832": 0.80, "9842": 0.65, "9852": 0.95,
    "9862": 1.10, "9872": 2.10, "9922": 1.50,
    "9972": 0.40, "9982": 0.60,

    # Photo printing
    "1999": 0.60, "1999x2": 1.20, "2000": 2.50, "2001": 5.00,

    # Other visible product codes
    "4820": 1.60, "4840": 2.10, "4860": 2.55, "4893": 3.20,
    "4900": 3.30, "4903": 3.30, "4909": 3.20,
    "4914": 2.30, "4915": 2.80,
}


# Descriptions keyed by the final/resolved PLU code.
PLU_DESCRIPTIONS = {
    # Black-and-white copying
    "10": "A4 B&W copy, one-sided, 1–100",
    "11": "A4 B&W copy, one-sided, 101–500",
    "12": "A4 B&W copy, one-sided, 501–1000",
    "13": "A4 B&W copy, one-sided, 1001+",
    "20": "A4 B&W copy, double-sided, 1–100",
    "21": "A4 B&W copy, double-sided, 101–500",
    "22": "A4 B&W copy, double-sided, 501–1000",
    "23": "A4 B&W copy, double-sided, 1001+",
    "30": "A3 B&W copy, one-sided, 1–100",
    "31": "A3 B&W copy, one-sided, 101–500",
    "32": "A3 B&W copy, one-sided, 501–1000",
    "33": "A3 B&W copy, one-sided, 1001+",
    "40": "A3 B&W copy, double-sided, 1–100",
    "41": "A3 B&W copy, double-sided, 101–500",
    "42": "A3 B&W copy, double-sided, 501–1000",
    "43": "A3 B&W copy, double-sided, 1001+",
    "50": "A4 B&W copy from glass, one-sided",
    "51": "A4 B&W copy from glass, double-sided",
    "52": "A3 B&W copy from glass, one-sided",
    "53": "A3 B&W copy from glass, double-sided",
    "60": "Reduction / enlargement",

    # Black-and-white printing
    "110": "A4 B&W print, one-sided, 1–100",
    "111": "A4 B&W print, one-sided, 101–500",
    "112": "A4 B&W print, one-sided, 501–1000",
    "113": "A4 B&W print, one-sided, 1001+",
    "120": "A4 B&W print, double-sided, 1–100",
    "121": "A4 B&W print, double-sided, 101–500",
    "122": "A4 B&W print, double-sided, 501–1000",
    "123": "A4 B&W print, double-sided, 1001+",
    "130": "A3 B&W print, one-sided, 1–100",
    "131": "A3 B&W print, one-sided, 101–500",
    "132": "A3 B&W print, one-sided, 501–1000",
    "133": "A3 B&W print, one-sided, 1001+",
    "140": "A3 B&W print, double-sided, 1–100",
    "141": "A3 B&W print, double-sided, 101–500",
    "142": "A3 B&W print, double-sided, 501–1000",
    "143": "A3 B&W print, double-sided, 1001+",

    # INEO B&W printing
    "310": "A4 B&W INEO print, one-sided, 1–100",
    "311": "A4 B&W INEO print, one-sided, 101–500",
    "312": "A4 B&W INEO print, one-sided, 501–1000",
    "313": "A4 B&W INEO print, one-sided, 1001+",
    "320": "A4 B&W INEO print, double-sided, 1–100",
    "321": "A4 B&W INEO print, double-sided, 101–500",
    "322": "A4 B&W INEO print, double-sided, 501–1000",
    "323": "A4 B&W INEO print, double-sided, 1001+",
    "330": "A3 B&W INEO print, one-sided, 1–100",
    "331": "A3 B&W INEO print, one-sided, 101–500",
    "332": "A3 B&W INEO print, one-sided, 501–1000",
    "333": "A3 B&W INEO print, one-sided, 1001+",
    "340": "A3 B&W INEO print, double-sided, 1–100",
    "341": "A3 B&W INEO print, double-sided, 101–500",
    "342": "A3 B&W INEO print, double-sided, 501–1000",
    "343": "A3 B&W INEO print, double-sided, 1001+",

    "450": "A4 B&W print on 90 gsm paper",
    "451": "A3 B&W print on 90 gsm paper",
    "650": "A4 B&W print on other paper",
    "651": "A3 B&W print on other paper",

    # Colour printing on 90 gsm paper
    "510": "A4 colour print, one-sided, 1–10",
    "511": "A4 colour print, one-sided, 11–50",
    "512": "A4 colour print, one-sided, 51–100",
    "513": "A4 colour print, one-sided, 101–500",
    "520": "A4 colour print, double-sided, 1–10",
    "521": "A4 colour print, double-sided, 11–50",
    "522": "A4 colour print, double-sided, 51–100",
    "523": "A4 colour print, double-sided, 101–500",
    "530": "A3 colour print, one-sided, 1–10",
    "531": "A3 colour print, one-sided, 11–50",
    "532": "A3 colour print, one-sided, 51–100",
    "533": "A3 colour print, one-sided, 101–500",
    "540": "A3 colour print, double-sided, 1–10",
    "541": "A3 colour print, double-sided, 11–50",
    "542": "A3 colour print, double-sided, 51–100",
    "543": "A3 colour print, double-sided, 101–500",

    # Colour printing on other paper
    "710": "A4 colour print on other paper, one-sided, 1–10",
    "711": "A4 colour print on other paper, one-sided, 11–50",
    "712": "A4 colour print on other paper, one-sided, 51–100",
    "713": "A4 colour print on other paper, one-sided, 101–500",
    "720": "A4 colour print on other paper, double-sided, 1–10",
    "721": "A4 colour print on other paper, double-sided, 11–50",
    "722": "A4 colour print on other paper, double-sided, 51–100",
    "723": "A4 colour print on other paper, double-sided, 101–500",
    "730": "A3 colour print on other paper, one-sided, 1–10",
    "731": "A3 colour print on other paper, one-sided, 11–50",
    "732": "A3 colour print on other paper, one-sided, 51–100",
    "733": "A3 colour print on other paper, one-sided, 101–500",
    "740": "A3 colour print on other paper, double-sided, 1–10",
    "741": "A3 colour print on other paper, double-sided, 11–50",
    "742": "A3 colour print on other paper, double-sided, 51–100",
    "743": "A3 colour print on other paper, double-sided, 101–500",

    # Paper and speciality products
    "899": "A4 decorative paper",
    "954": "80 gsm coloured A4 paper",
    "956": "160 gsm coloured A4 paper",
    "966": "SRA3 adhesive sheet",
    "988": "135–250 gsm A4 photo paper",
    "999": "A4 transparency",
    "1090": "350–400 gsm A3/SRA3 photo paper",
    "1092": "Textured A3/SRA3 paper",
    "1095": "Long-lasting tear-resistant A3 paper",
    "1096": "Magnetic SRA3 paper",
    "1184": "Pioneer Navigator A4 paper pack",
    "1319": "A4 Curious Collection paper",
    "2022": "A3 transparency",
    "2851": "SRA3 transfer paper, 350 gsm",
    "9502": "80 gsm A4 paper",
    "9504": "80 gsm paper pack",
    "9512": "80 gsm A3 paper",
    "9582": "A4 transfer paper",
    "9592": "A3 transfer paper",
    "9822": "160 gsm A4 paper",
    "9832": "160 gsm A3 paper",
    "9842": "200 gsm A4 paper",
    "9852": "200 gsm A3 paper",
    "9862": "250–300 gsm A4 paper",
    "9872": "250–300 gsm A3 paper",
    "9922": "A4 adhesive paper",
    "9972": "120 gsm A4 paper",
    "9982": "120 gsm A3 paper",

    # Photo printing
    "1999": "10 × 15 photo print",
    "1999x2": "A5 photo print",
    "2000": "A4 photo print",
    "2001": "Document photo",

    # Other visible codes; product names were not visible in the supplied photo
    "4820": "Other product", "4840": "Other product", "4860": "Other product",
    "4893": "Other product", "4900": "Other product", "4903": "Other product",
    "4909": "Other product", "4914": "Other product", "4915": "Other product",
}


# Additional confirmed store products. Existing entries above remain authoritative
# when the same PLU appeared on more than one price sheet.
PLU_PRICES.update({
    "70":0.13,"71":0.13,"72":0.25,"73":0.25,"74":0.25,"75":0.25,"76":0.45,"77":0.45,
    "151":0.30,"153":0.40,"156":0.30,"160":2.00,"161":2.00,"162":3.00,"163":3.00,"172":1.00,
    "396":3.50,"397":3.50,"398":2.00,"399":1.50,"470":0.65,"471":0.65,"472":1.30,"473":1.30,
    "474":1.30,"475":1.30,"476":2.60,"477":2.60,"496":6.50,"497":6.50,"498":3.50,"499":2.00,
    "500":7.00,"561":0.80,"563":1.10,
    "751":8.00,"753":8.50,"754":8.60,"755":8.90,"756":8.90,"757":9.00,"758":9.40,"759":9.80,"760":10.20,
    "767":4.50,"768":5.20,"769":4.50,"770":4.90,"771":5.20,"772":5.50,"773":6.00,"774":6.20,"775":6.50,"776":6.80,
    "791":3.00,"796":5.50,
    "800":0.50,"801":0.60,"802":0.70,"803":0.80,"804":0.90,"805":1.00,"806":1.50,"807":2.20,"808":2.70,"809":3.20,
    "810":0.50,"811":0.60,"812":0.70,"820":0.50,"821":0.60,"822":0.70,"824":0.90,"825":1.00,"827":2.00,"828":2.50,"829":3.20,
    "830":0.25,"831":0.30,"832":0.35,"840":0.25,"841":0.30,"842":0.30,"850":3.20,
    "880":3.00,"881":3.50,"882":4.50,"883":5.50,"884":6.00,"891":1.20,"892":1.20,"893":1.20,"894":2.00,"895":1.20,"896":2.00,
    "900":1.00,"901":1.00,"903":1.20,"904":1.20,"905":2.00,"906":2.50,"907":3.50,"908":3.50,"910":2.80,"914":1.50,"918":4.20,
    "920":0.25,"921":0.60,"922":0.80,"924":0.60,"926":0.80,"927":1.50,"928":2.00,"929":2.50,"930":1.00,
    "933":0.10,"934":0.50,"935":1.00,"936":0.50,"941":0.50,"942":1.20,
    "1001":6.50,"1002":1.20,"1003":1.00,"1004":1.30,"1006":0.20,"1007":1.50,"1009":2.00,
    "1011":6.00,"1012":7.00,"1015":0.50,"1016":2.50,"1019":1.00,"1020":0.70,"1021":0.60,"1025":8.00,
    "1033":6.50,"1042":3.00,"1044":4.50,"1062":0.50,"1066":1.20,"1067":1.50,"1068":1.40,"1069":1.80,
    "1070":1.70,"1071":2.70,"1072":2.70,"1073":3.70,"1074":3.60,"1075":5.00,"1076":5.20,"1077":5.00,"1079":7.00,
    "1081":0.70,"1082":1.00,"1083":0.90,"1084":1.30,"1085":0.90,"1086":1.50,"1087":1.60,"1088":1.80,
    "1103":3.50,"1109":7.50,"1113":1.50,"1117":2.90,"1118":2.50,"1120":3.00,"1124":5.00,
    "1137":1.50,"1144":2.50,"1149":10.00,"1153":10.00,"1164":2.00,"1173":7.50,"1182":2.00,"1183":1.00,
    "1214":4.50,"1216":28.00,"1220":9.00,"1223":15.00,"1240":6.00,"1244":3.00,"1245":2.00,"1249":3.00,
    "1253":3.00,"1268":20.00,"1270":2.00,"1274":3.50,"1275":3.00,"1284":2.00,"1289":0.50,"1293":18.00,"1296":15.00,
    "1300":1.00,"1301":1.00,"1302":1.00,"1303":1.00,"1305":5.50,"1317":3.50,"1318":1.50,"1327":3.00,"1345":22.00,"1349":4.00,
    "1350":4.50,"1351":5.00,"1354":8.00,"1355":18.00,"1360":65.00,"1362":2.50,"1364":3.00,
    "2003":1.00,"2006":30.00,"2007":28.00,"2009":5.00,"2012":32.00,"2025":4.00,"2029":2.50,"2032":5.00,"2033":6.50,"2034":10.00,
    "2047":8.00,"2052":2.00,"2054":26.00,"2059":1.00,"2062":3.00,"2070":2.50,"2074":5.00,"2079":2.50,"2080":3.50,"2081":4.50,
    "2088":24.00,"2089":12.00,"2090":6.00,"2091":3.00,
    "8800":5.00,"9302":1.00,"9310":1.00,"9311":1.00,"9312":25.00,"9313":35.00,"9314":3.00,"9315":3.00,
    "9400":0.70,"9503":0.06,"9513":0.07,
})

PLU_DESCRIPTIONS.update({
    "70":"Self-service A4 B&W print, one-sided","71":"Self-service A4 B&W copy, one-sided","72":"Self-service A3 B&W print, one-sided","73":"Self-service A3 B&W copy, one-sided",
    "74":"Self-service A4 B&W print, double-sided","75":"Self-service A4 B&W copy, double-sided","76":"Self-service A3 B&W print, double-sided","77":"Self-service A3 B&W copy, double-sided",
    "151":"A4 B&W scan","153":"A3 B&W scan","156":"Document upload / sending","160":"DVD","161":"CD","162":"Disc recording","163":"Disc copy","172":"Computer use, 15 minutes",
    "396":"Large-format B&W scanning, 1 m²","397":"A0 B&W scanning","398":"A1 B&W scanning","399":"A2 B&W scanning",
    "470":"Self-service A4 colour print, one-sided","471":"Self-service A4 colour copy, one-sided","472":"Self-service A3 colour print, one-sided","473":"Self-service A3 colour copy, one-sided",
    "474":"Self-service A4 colour print, double-sided","475":"Self-service A4 colour copy, double-sided","476":"Self-service A3 colour print, double-sided","477":"Self-service A3 colour copy, double-sided",
    "496":"Large-format colour scanning, 1 m²","497":"A0 colour scanning","498":"A1 colour scanning","499":"A2 colour scanning","500":"Computer editing, 15 minutes",
    "561":"A4 colour scan","563":"A3 colour scan",
    "751":"Leather-look hard cover AA, 20–40 sheets","753":"Leather-look hard cover A, 41–90 sheets","754":"Leather-look hard cover B, 91–120 sheets","755":"Leather-look hard cover C, 121–145 sheets",
    "756":"Leather-look 18 mm hard cover with thermo binding, 130–160 sheets","757":"Leather-look 21 mm hard cover with thermo binding, 160–190 sheets","758":"Leather-look 24 mm hard cover with thermo binding, 190–220 sheets","759":"Leather-look 30 mm hard cover with thermo binding, 220–280 sheets","760":"Leather-look 36 mm hard cover with thermo binding, 280–340 sheets",
    "767":"Matte thermal binding, 3 mm, 1–10 sheets","768":"Matte thermal binding, 5 mm, 25–40 sheets","769":"Clear thermal binding, 1 mm, 1–10 sheets","770":"Clear thermal binding, 3 mm, 10–25 sheets",
    "771":"Clear thermal binding, 5 mm, 25–40 sheets","772":"Clear thermal binding, 7 mm, 40–55 sheets","773":"Clear thermal binding, 9 mm, 55–75 sheets","774":"Clear thermal binding, 12 mm, 75–100 sheets",
    "775":"Clear thermal binding, 15 mm, 100–130 sheets","776":"Clear thermal binding, 18 mm, 130–160 sheets","791":"Binding channels, landscape","796":"White soft cover G/A",
    "800":"White plastic spiral 6–8 mm","801":"White plastic spiral 10 mm","802":"White plastic spiral 12 mm","803":"White plastic spiral 14 mm","804":"Plastic spiral 16 mm, red or black","805":"White plastic spiral 19 mm",
    "806":"White plastic spiral 22–25 mm","807":"White plastic spiral 28–32 mm","808":"White plastic spiral 38–45 mm","809":"White plastic spiral 50 mm or larger",
    "810":"Red plastic spiral 6–8 mm","811":"Red plastic spiral 10 mm","812":"Red plastic spiral 12 mm","820":"Black plastic spiral 6–8 mm","821":"Black plastic spiral 10 mm","822":"Black plastic spiral 12 mm",
    "824":"Red or black plastic spiral 16 mm","825":"Black plastic spiral 19 mm","827":"Black plastic spiral 28–32 mm","828":"Black plastic spiral 38–45 mm","829":"Black plastic spiral 50 mm or larger",
    "830":"Blue plastic spiral 6–8 mm","831":"Blue plastic spiral 10 mm","832":"Blue plastic spiral 12 mm","840":"Green plastic spiral 6–8 mm","841":"Green plastic spiral 10 mm","842":"Green plastic spiral 12 mm","850":"Red screw-binding spiral 22 mm",
    "880":"Plastic spiral binding, 1–50 sheets","881":"Plastic spiral binding, 51–150 sheets","882":"Plastic spiral binding, 151–250 sheets","883":"Plastic spiral binding, 251–340 sheets","884":"Plastic spiral binding, 341+ sheets",
    "891":"A4 clear binding cover","892":"A4 coloured binding cover","893":"A4 matte binding cover","894":"A3 clear binding cover","895":"A4 binding back cover","896":"A3 binding back cover",
    "900":"White screw-binding spiral 8 mm","901":"Black screw-binding spiral 8 mm","903":"White screw-binding spiral 10 mm","904":"Black screw-binding spiral 10 mm",
    "1004":"White screw-binding spiral 12 mm","1068":"White screw-binding spiral 14 mm","1070":"White screw-binding spiral 16 mm","1072":"White screw-binding spiral 20 mm","1074":"White screw-binding spiral 25 mm","1076":"White screw-binding spiral 32 mm",
    "1082":"Red screw-binding spiral 8 mm","1084":"Red screw-binding spiral 10 mm","1086":"Red screw-binding spiral 12 mm","1087":"Red screw-binding spiral 14 mm","1088":"Red screw-binding spiral 16 mm",
    "1003":"Black screw-binding spiral 12 mm","1067":"Black screw-binding spiral 14 mm","1069":"Black screw-binding spiral 16 mm","1071":"Black screw-binding spiral 20 mm","1073":"Black screw-binding spiral 25 mm","1075":"Black screw-binding spiral 32 mm","1077":"Black screw-binding spiral 35 mm","1079":"Black screw-binding spiral 51 mm",
    "1081":"Blue screw-binding spiral 8 mm","1083":"Blue screw-binding spiral 10 mm","1085":"Blue screw-binding spiral 12 mm","1002":"Clear screw-binding spiral 10 mm","1066":"Clear screw-binding spiral 12 mm",
    "905":"A5–A6 lamination, 125 micron","906":"A4 lamination, 100–125 micron","907":"A3 lamination, 100–150 micron","908":"A4 adhesive lamination, 100 micron","910":"A4 thick lamination, 150–175 micron","914":"75 × 105 mm lamination, 175 micron","918":"A3 matte lamination, 125 micron",
    "920":"Single-hole punching","921":"Two-hole punching, up to 50 sheets","922":"Four-hole punching, up to 50 sheets","924":"Corner rounding, one item","926":"Stapling, 1–50 sheets","927":"Stapling, 51–100 sheets","928":"Stapling, 101–150 sheets","929":"Metal/plastic rivets and similar work","930":"Other work",
    "933":"Cutting, one cut","934":"Guillotine cutting","935":"Delivery by post or courier","936":"Rivet insertion","941":"Notary sticker","942":"Notarial document binding",
    "1001":"Hard binding, landscape","1006":"A4/A5 document sleeve","1007":"A3 document sleeve","1009":"Folder with metal fastener","1011":"Lever arch file, 50 mm","1012":"Lever arch file, 70 mm","1015":"Paper CD sleeve","1016":"Adhesive CD sleeve",
    "1019":"Standard A4 envelope","1020":"Standard A5 envelope","1021":"Standard A6 / long envelope","1025":"Lever arch file, 80 mm","1033":"Ring binder","1042":"Small glue stick, 8 g","1044":"Scissors",
    "1062":"Cardboard business-card box","1103":"Instant glue","1109":"Double-sided adhesive tape","1113":"Standard pencil","1117":"Cardboard box","1118":"Automatic pen","1120":"Wide-format tube, narrow","1124":"Thick adhesive tape",
    "1137":"Coloured push pins","1144":"Staples 24/6","1149":"A4 frame","1153":"A4 frame with thick border","1164":"Paper clips","1173":"Mounting squares","1182":"Paper bags","1183":"Brown kraft A4 envelope",
    "1214":"Large coloured sheets","1216":"40 × 60 (A2) or 40 × 50 frame","1220":"15 × 21 (A5) frame with thick border","1223":"A3 frame","1240":"Small hole punch","1244":"PVA glue","1245":"Thin adhesive tape","1249":"Correction fluid",
    "1253":"Pack of binder clips","1268":"16 GB USB flash drive","1270":"CopyPro pen","1274":"Markers","1275":"Adhesive tape with dispenser","1284":"Postcards","1289":"Binder clip","1293":"8 GB USB flash drive","1296":"Large hole punch",
    "1300":"Printer cartridges","1301":"Canvas","1302":"Stamps","1303":"Stamp rubber","1305":"Window boxes","1317":"Glossy A4 envelope","1318":"Sticky notes","1327":"Gift bags","1345":"32 GB USB flash drive","1349":"Marker set",
    "1350":"Wide-format tube, wide","1351":"10 × 15 (A6) frame","1354":"Printing on a bag","1355":"Printing on a T-shirt","1360":"A0 frame","1362":"Textured A5 envelope","1364":"Pack of 20 document sleeves",
    "2003":"Printing on a mug","2006":"60 × 90 frame","2007":"50 × 70 frame","2009":"Satin ribbon","2012":"A1 frame","2025":"Large glossy coloured sheets","2029":"Gift ribbon","2032":"Memo note sheets","2033":"A5/A6 information holder","2034":"A4 information holder",
    "2047":"Clear bags, 100 pcs","2052":"Adhesive ribbon","2054":"64 GB USB flash drive","2059":"Small calendars","2062":"Navigator paper, 150 sheets","2070":"CopyPro A6 / 14 × 18 envelopes","2074":"Decorative elements","2079":"CopyPro C65 envelopes","2080":"CopyPro A5 envelopes","2081":"CopyPro A4 envelopes",
    "2088":"A0 frame backing board","2089":"A1 frame backing board","2090":"A2 frame backing board","2091":"A3 frame backing board",
    "8800":"Rebinding","9302":"Plotting","9310":"Cup/trophy item","9311":"Medals","9312":"One-sided business-card layout design","9313":"Double-sided business-card layout design","9314":"Badge production","9315":"Keychain production",
    "9400":"Notarial thread","9503":"A4 self-service paper, one sheet","9513":"A3 self-service paper, one sheet",
})

# Clear PLU/name pairs whose price is variable or was not unambiguous.
PLU_DESCRIPTIONS.update({
    "1219": "13 × 18 / A5 frame",
    "1237": "Staple remover",
    "1294": "Padded envelope",
    "1991": "Jigsaw puzzle production (price varies by puzzle)",
    "9316": "Magnet production",
})
PLU_PRICES["9316"] = 2.00


# Wide-format colour printing and materials.
# Prices and descriptions are taken from the current in-store wide-format list.
WIDE_FORMAT_ITEMS = {
    "4800": ("80 gsm paprastas / brėžinys", 0.95, "drawing"),
    "4810": ("80 gsm paprastas / tekstas+pav.", 1.40, "partial"),
    "4820": ("80 gsm paprastas / paveikslas", 1.60, "full"),
    "4830": ("120 gsm storesnis / tekstas+pav.", 1.70, "partial"),
    "4840": ("120 gsm storesnis / paveikslas", 2.10, "full"),
    "4913": ("140 gsm storas / tekstas+paveikslas", 2.35, "partial"),
    "4914": ("140 gsm storas / paveikslas", 2.30, "full"),
    "4850": ("180 gsm storas / tekstas+paveikslas", 2.40, "partial"),
    "4860": ("180 gsm storas / paveikslas", 2.55, "full"),
    "4894": ("Satin / fotopopierius / tekstas+pav.", 2.80, "partial"),
    "4893": ("Satin / fotopopierius / paveikslas", 3.20, "full"),
    "4915": ("Sintetinė drobė", 2.80, "fixed"),
    "4900": ("Natūrali drobė", 3.30, "fixed"),
    "4902": ("Film plėvelė / tekstas+paveikslas", 2.90, "partial"),
    "4903": ("Film plėvelė / paveikslas", 3.30, "full"),
    "4908": ("Lipdukas PVC / tekstas+paveikslas", 3.00, "partial"),
    "4909": ("Lipdukas PVC / paveikslas", 3.20, "full"),
    "4910": ("Kalkė / brėžinys", 1.70, "drawing"),
    "4911": ("Kalkė / tekstas+paveikslas", 2.20, "partial"),
    "4912": ("Kalkė / paveikslas", 2.50, "full"),
    "3009": ("Spausdinimas ant vatmano", 14.00, "fixed"),
    "4916": ("Karštas laminatas", 1.80, "fixed"),
    "4917": ("Magnetinė plėvelė", 2.50, "fixed"),
    "3008": ("Vatmanas 90×64 cm (240 gsm)", 3.00, "fixed"),
}
for _code, (_description, _price, _coverage) in WIDE_FORMAT_ITEMS.items():
    PLU_DESCRIPTIONS[_code] = _description
    PLU_PRICES[_code] = _price

# Editable external data. Paths are stored in AppData/CopyPro/settings.json.
_EMBEDDED_CODE_ITEMS = [{'code': '10', 'name': 'A4 nespalvota kopija, vienpusis, 1–100', 'price': 0.15, 'aliases': []},
 {'code': '11', 'name': 'A4 nespalvota kopija, vienpusis, 101–500', 'price': 0.1, 'aliases': []},
 {'code': '12', 'name': 'A4 nespalvota kopija, vienpusis, 501–1000', 'price': 0.08, 'aliases': []},
 {'code': '13', 'name': 'A4 nespalvota kopija, vienpusis, 1001+', 'price': 0.07, 'aliases': []},
 {'code': '20', 'name': 'A4 nespalvota kopija, dvipusis, 1–100', 'price': 0.25, 'aliases': []},
 {'code': '21', 'name': 'A4 nespalvota kopija, dvipusis, 101–500', 'price': 0.2, 'aliases': []},
 {'code': '22', 'name': 'A4 nespalvota kopija, dvipusis, 501–1000', 'price': 0.15, 'aliases': []},
 {'code': '23', 'name': 'A4 nespalvota kopija, dvipusis, 1001+', 'price': 0.12, 'aliases': []},
 {'code': '30', 'name': 'A3 nespalvota kopija, vienpusis, 1–100', 'price': 0.3, 'aliases': []},
 {'code': '31', 'name': 'A3 nespalvota kopija, vienpusis, 101–500', 'price': 0.25, 'aliases': []},
 {'code': '32', 'name': 'A3 nespalvota kopija, vienpusis, 501–1000', 'price': 0.2, 'aliases': []},
 {'code': '33', 'name': 'A3 nespalvota kopija, vienpusis, 1001+', 'price': 0.15, 'aliases': []},
 {'code': '40', 'name': 'A3 nespalvota kopija, dvipusis, 1–100', 'price': 0.5, 'aliases': []},
 {'code': '41', 'name': 'A3 nespalvota kopija, dvipusis, 101–500', 'price': 0.45, 'aliases': []},
 {'code': '42', 'name': 'A3 nespalvota kopija, dvipusis, 501–1000', 'price': 0.4, 'aliases': []},
 {'code': '43', 'name': 'A3 nespalvota kopija, dvipusis, 1001+', 'price': 0.25, 'aliases': []},
 {'code': '50', 'name': 'A4 nespalvota kopija nuo stiklo, vienpusis', 'price': 0.3, 'aliases': []},
 {'code': '51', 'name': 'A4 nespalvota kopija nuo stiklo, dvipusis', 'price': 0.5, 'aliases': []},
 {'code': '52', 'name': 'A3 nespalvota kopija nuo stiklo, vienpusis', 'price': 0.4, 'aliases': []},
 {'code': '53', 'name': 'A3 nespalvota kopija nuo stiklo, dvipusis', 'price': 0.6, 'aliases': []},
 {'code': '60', 'name': 'Mazinimas / didinimas', 'price': 0.2, 'aliases': []},
 {'code': '70', 'name': 'Self-service A4 nespalvota spausdinimas, vienpusis', 'price': 0.13, 'aliases': []},
 {'code': '71', 'name': 'Self-service A4 nespalvota kopija, vienpusis', 'price': 0.13, 'aliases': []},
 {'code': '72', 'name': 'Self-service A3 nespalvota spausdinimas, vienpusis', 'price': 0.25, 'aliases': []},
 {'code': '73', 'name': 'Self-service A3 nespalvota kopija, vienpusis', 'price': 0.25, 'aliases': []},
 {'code': '74', 'name': 'Self-service A4 nespalvota spausdinimas, dvipusis', 'price': 0.25, 'aliases': []},
 {'code': '75', 'name': 'Self-service A4 nespalvota kopija, dvipusis', 'price': 0.25, 'aliases': []},
 {'code': '76', 'name': 'Self-service A3 nespalvota spausdinimas, dvipusis', 'price': 0.45, 'aliases': []},
 {'code': '77', 'name': 'Self-service A3 nespalvota kopija, dvipusis', 'price': 0.45, 'aliases': []},
 {'code': '110', 'name': 'A4 nespalvota spausdinimas, vienpusis, 1–100', 'price': 0.15, 'aliases': []},
 {'code': '111', 'name': 'A4 nespalvota spausdinimas, vienpusis, 101–500', 'price': 0.1, 'aliases': []},
 {'code': '112', 'name': 'A4 nespalvota spausdinimas, vienpusis, 501–1000', 'price': 0.08, 'aliases': []},
 {'code': '113', 'name': 'A4 nespalvota spausdinimas, vienpusis, 1001+', 'price': 0.07, 'aliases': []},
 {'code': '120', 'name': 'A4 nespalvota spausdinimas, dvipusis, 1–100', 'price': 0.25, 'aliases': []},
 {'code': '121', 'name': 'A4 nespalvota spausdinimas, dvipusis, 101–500', 'price': 0.2, 'aliases': []},
 {'code': '122', 'name': 'A4 nespalvota spausdinimas, dvipusis, 501–1000', 'price': 0.15, 'aliases': []},
 {'code': '123', 'name': 'A4 nespalvota spausdinimas, dvipusis, 1001+', 'price': 0.12, 'aliases': []},
 {'code': '130', 'name': 'A3 nespalvota spausdinimas, vienpusis, 1–100', 'price': 0.3, 'aliases': []},
 {'code': '131', 'name': 'A3 nespalvota spausdinimas, vienpusis, 101–500', 'price': 0.25, 'aliases': []},
 {'code': '132', 'name': 'A3 nespalvota spausdinimas, vienpusis, 501–1000', 'price': 0.2, 'aliases': []},
 {'code': '133', 'name': 'A3 nespalvota spausdinimas, vienpusis, 1001+', 'price': 0.15, 'aliases': []},
 {'code': '140', 'name': 'A3 nespalvota spausdinimas, dvipusis, 1–100', 'price': 0.5, 'aliases': []},
 {'code': '141', 'name': 'A3 nespalvota spausdinimas, dvipusis, 101–500', 'price': 0.45, 'aliases': []},
 {'code': '142', 'name': 'A3 nespalvota spausdinimas, dvipusis, 501–1000', 'price': 0.4, 'aliases': []},
 {'code': '143', 'name': 'A3 nespalvota spausdinimas, dvipusis, 1001+', 'price': 0.25, 'aliases': []},
 {'code': '151', 'name': 'A4 nespalvota skenavimas', 'price': 0.3, 'aliases': []},
 {'code': '153', 'name': 'A3 nespalvota skenavimas', 'price': 0.4, 'aliases': []},
 {'code': '156', 'name': 'Document upload / sending', 'price': 0.3, 'aliases': []},
 {'code': '160', 'name': 'DVD', 'price': 2.0, 'aliases': []},
 {'code': '161', 'name': 'CD', 'price': 2.0, 'aliases': []},
 {'code': '162', 'name': 'Disc recording', 'price': 3.0, 'aliases': []},
 {'code': '163', 'name': 'Disc kopija', 'price': 3.0, 'aliases': []},
 {'code': '172', 'name': 'Computer use, 15 minutes', 'price': 1.0, 'aliases': []},
 {'code': '310', 'name': 'A4 nespalvota INEO spausdinimas, vienpusis, 1–100', 'price': 0.15, 'aliases': []},
 {'code': '311', 'name': 'A4 nespalvota INEO spausdinimas, vienpusis, 101–500', 'price': 0.1, 'aliases': []},
 {'code': '312', 'name': 'A4 nespalvota INEO spausdinimas, vienpusis, 501–1000', 'price': 0.08, 'aliases': []},
 {'code': '313', 'name': 'A4 nespalvota INEO spausdinimas, vienpusis, 1001+', 'price': 0.07, 'aliases': []},
 {'code': '320', 'name': 'A4 nespalvota INEO spausdinimas, dvipusis, 1–100', 'price': 0.25, 'aliases': []},
 {'code': '321', 'name': 'A4 nespalvota INEO spausdinimas, dvipusis, 101–500', 'price': 0.2, 'aliases': []},
 {'code': '322', 'name': 'A4 nespalvota INEO spausdinimas, dvipusis, 501–1000', 'price': 0.15, 'aliases': []},
 {'code': '323', 'name': 'A4 nespalvota INEO spausdinimas, dvipusis, 1001+', 'price': 0.12, 'aliases': []},
 {'code': '330', 'name': 'A3 nespalvota INEO spausdinimas, vienpusis, 1–100', 'price': 0.3, 'aliases': []},
 {'code': '331', 'name': 'A3 nespalvota INEO spausdinimas, vienpusis, 101–500', 'price': 0.25, 'aliases': []},
 {'code': '332', 'name': 'A3 nespalvota INEO spausdinimas, vienpusis, 501–1000', 'price': 0.2, 'aliases': []},
 {'code': '333', 'name': 'A3 nespalvota INEO spausdinimas, vienpusis, 1001+', 'price': 0.15, 'aliases': []},
 {'code': '340', 'name': 'A3 nespalvota INEO spausdinimas, dvipusis, 1–100', 'price': 0.5, 'aliases': []},
 {'code': '341', 'name': 'A3 nespalvota INEO spausdinimas, dvipusis, 101–500', 'price': 0.45, 'aliases': []},
 {'code': '342', 'name': 'A3 nespalvota INEO spausdinimas, dvipusis, 501–1000', 'price': 0.4, 'aliases': []},
 {'code': '343', 'name': 'A3 nespalvota INEO spausdinimas, dvipusis, 1001+', 'price': 0.25, 'aliases': []},
 {'code': '396', 'name': 'Large-format nespalvota skenavimasning, 1 m²', 'price': 3.5, 'aliases': []},
 {'code': '397', 'name': 'A0 nespalvota skenavimasning', 'price': 3.5, 'aliases': []},
 {'code': '398', 'name': 'A1 nespalvota skenavimasning', 'price': 2.0, 'aliases': []},
 {'code': '399', 'name': 'A2 nespalvota skenavimasning', 'price': 1.5, 'aliases': []},
 {'code': '450', 'name': 'A4 nespalvota spausdinimas ant 90 gsm popieriaus', 'price': 0.5, 'aliases': []},
 {'code': '451', 'name': 'A3 nespalvota spausdinimas ant 90 gsm popieriaus', 'price': 0.8, 'aliases': []},
 {'code': '470', 'name': 'Self-service A4 spalvota spausdinimas, vienpusis', 'price': 0.65, 'aliases': []},
 {'code': '471', 'name': 'Self-service A4 spalvota kopija, vienpusis', 'price': 0.65, 'aliases': []},
 {'code': '472', 'name': 'Self-service A3 spalvota spausdinimas, vienpusis', 'price': 1.3, 'aliases': []},
 {'code': '473', 'name': 'Self-service A3 spalvota kopija, vienpusis', 'price': 1.3, 'aliases': []},
 {'code': '474', 'name': 'Self-service A4 spalvota spausdinimas, dvipusis', 'price': 1.3, 'aliases': []},
 {'code': '475', 'name': 'Self-service A4 spalvota kopija, dvipusis', 'price': 1.3, 'aliases': []},
 {'code': '476', 'name': 'Self-service A3 spalvota spausdinimas, dvipusis', 'price': 2.6, 'aliases': []},
 {'code': '477', 'name': 'Self-service A3 spalvota kopija, dvipusis', 'price': 2.6, 'aliases': []},
 {'code': '496', 'name': 'Large-format spalvota skenavimasning, 1 m²', 'price': 6.5, 'aliases': []},
 {'code': '497', 'name': 'A0 spalvota skenavimasning', 'price': 6.5, 'aliases': []},
 {'code': '498', 'name': 'A1 spalvota skenavimasning', 'price': 3.5, 'aliases': []},
 {'code': '499', 'name': 'A2 spalvota skenavimasning', 'price': 2.0, 'aliases': []},
 {'code': '500', 'name': 'Computer editing, 15 minutes', 'price': 7.0, 'aliases': []},
 {'code': '510', 'name': 'A4 spalvota spausdinimas, vienpusis, 1–10', 'price': 0.95, 'aliases': []},
 {'code': '511', 'name': 'A4 spalvota spausdinimas, vienpusis, 11–50', 'price': 0.85, 'aliases': []},
 {'code': '512', 'name': 'A4 spalvota spausdinimas, vienpusis, 51–100', 'price': 0.75, 'aliases': []},
 {'code': '513', 'name': 'A4 spalvota spausdinimas, vienpusis, 101–500', 'price': 0.65, 'aliases': []},
 {'code': '520', 'name': 'A4 spalvota spausdinimas, dvipusis, 1–10', 'price': 1.65, 'aliases': []},
 {'code': '521', 'name': 'A4 spalvota spausdinimas, dvipusis, 11–50', 'price': 1.5, 'aliases': []},
 {'code': '522', 'name': 'A4 spalvota spausdinimas, dvipusis, 51–100', 'price': 1.3, 'aliases': []},
 {'code': '523', 'name': 'A4 spalvota spausdinimas, dvipusis, 101–500', 'price': 1.1, 'aliases': []},
 {'code': '530', 'name': 'A3 spalvota spausdinimas, vienpusis, 1–10', 'price': 1.6, 'aliases': []},
 {'code': '531', 'name': 'A3 spalvota spausdinimas, vienpusis, 11–50', 'price': 1.5, 'aliases': []},
 {'code': '532', 'name': 'A3 spalvota spausdinimas, vienpusis, 51–100', 'price': 1.35, 'aliases': []},
 {'code': '533', 'name': 'A3 spalvota spausdinimas, vienpusis, 101–500', 'price': 1.2, 'aliases': []},
 {'code': '540', 'name': 'A3 spalvota spausdinimas, dvipusis, 1–10', 'price': 2.9, 'aliases': []},
 {'code': '541', 'name': 'A3 spalvota spausdinimas, dvipusis, 11–50', 'price': 2.75, 'aliases': []},
 {'code': '542', 'name': 'A3 spalvota spausdinimas, dvipusis, 51–100', 'price': 2.5, 'aliases': []},
 {'code': '543', 'name': 'A3 spalvota spausdinimas, dvipusis, 101–500', 'price': 2.0, 'aliases': []},
 {'code': '561', 'name': 'A4 spalvota skenavimas', 'price': 0.8, 'aliases': []},
 {'code': '563', 'name': 'A3 spalvota skenavimas', 'price': 1.1, 'aliases': []},
 {'code': '650', 'name': 'A4 nespalvota spausdinimas ant kito popieriaus', 'price': 0.5, 'aliases': []},
 {'code': '651', 'name': 'A3 nespalvota spausdinimas ant kito popieriaus', 'price': 0.8, 'aliases': []},
 {'code': '710', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, vienpusis, 1–10', 'price': 0.95, 'aliases': []},
 {'code': '711', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, vienpusis, 11–50', 'price': 0.85, 'aliases': []},
 {'code': '712', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, vienpusis, 51–100', 'price': 0.75, 'aliases': []},
 {'code': '713', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, vienpusis, 101–500', 'price': 0.65, 'aliases': []},
 {'code': '720', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, dvipusis, 1–10', 'price': 1.65, 'aliases': []},
 {'code': '721', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, dvipusis, 11–50', 'price': 1.5, 'aliases': []},
 {'code': '722', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, dvipusis, 51–100', 'price': 1.3, 'aliases': []},
 {'code': '723', 'name': 'A4 spalvota spausdinimas ant kito popieriaus, dvipusis, 101–500', 'price': 1.1, 'aliases': []},
 {'code': '730', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, vienpusis, 1–10', 'price': 1.6, 'aliases': []},
 {'code': '731', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, vienpusis, 11–50', 'price': 1.5, 'aliases': []},
 {'code': '732', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, vienpusis, 51–100', 'price': 1.35, 'aliases': []},
 {'code': '733', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, vienpusis, 101–500', 'price': 1.2, 'aliases': []},
 {'code': '740', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, dvipusis, 1–10', 'price': 2.9, 'aliases': []},
 {'code': '741', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, dvipusis, 11–50', 'price': 2.75, 'aliases': []},
 {'code': '742', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, dvipusis, 51–100', 'price': 2.5, 'aliases': []},
 {'code': '743', 'name': 'A3 spalvota spausdinimas ant kito popieriaus, dvipusis, 101–500', 'price': 2.0, 'aliases': []},
 {'code': '751', 'name': 'Odos imitacijos kietas virselis AA, 20–40 lapai', 'price': 8.0, 'aliases': []},
 {'code': '753', 'name': 'Odos imitacijos kietas virselis A, 41–90 lapai', 'price': 8.5, 'aliases': []},
 {'code': '754', 'name': 'Odos imitacijos kietas virselis B, 91–120 lapai', 'price': 8.6, 'aliases': []},
 {'code': '755', 'name': 'Odos imitacijos kietas virselis C, 121–145 lapai', 'price': 8.9, 'aliases': []},
 {'code': '756', 'name': 'Odos imitacijos 18 mm kietas virselis with terminis irisimas, 130–160 lapai', 'price': 8.9, 'aliases': []},
 {'code': '757', 'name': 'Odos imitacijos 21 mm kietas virselis with terminis irisimas, 160–190 lapai', 'price': 9.0, 'aliases': []},
 {'code': '758', 'name': 'Odos imitacijos 24 mm kietas virselis with terminis irisimas, 190–220 lapai', 'price': 9.4, 'aliases': []},
 {'code': '759', 'name': 'Odos imitacijos 30 mm kietas virselis with terminis irisimas, 220–280 lapai', 'price': 9.8, 'aliases': []},
 {'code': '760', 'name': 'Odos imitacijos 36 mm kietas virselis with terminis irisimas, 280–340 lapai', 'price': 10.2, 'aliases': []},
 {'code': '767', 'name': 'Matinis terminis irisimas, 3 mm, 1–10 lapai', 'price': 4.5, 'aliases': []},
 {'code': '768', 'name': 'Matinis terminis irisimas, 5 mm, 25–40 lapai', 'price': 5.2, 'aliases': []},
 {'code': '769', 'name': 'Skaidrus terminis irisimas, 1 mm, 1–10 lapai', 'price': 4.5, 'aliases': []},
 {'code': '770', 'name': 'Skaidrus terminis irisimas, 3 mm, 10–25 lapai', 'price': 4.9, 'aliases': []},
 {'code': '771', 'name': 'Skaidrus terminis irisimas, 5 mm, 25–40 lapai', 'price': 5.2, 'aliases': []},
 {'code': '772', 'name': 'Skaidrus terminis irisimas, 7 mm, 40–55 lapai', 'price': 5.5, 'aliases': []},
 {'code': '773', 'name': 'Skaidrus terminis irisimas, 9 mm, 55–75 lapai', 'price': 6.0, 'aliases': []},
 {'code': '774', 'name': 'Skaidrus terminis irisimas, 12 mm, 75–100 lapai', 'price': 6.2, 'aliases': []},
 {'code': '775', 'name': 'Skaidrus terminis irisimas, 15 mm, 100–130 lapai', 'price': 6.5, 'aliases': []},
 {'code': '776', 'name': 'Skaidrus terminis irisimas, 18 mm, 130–160 lapai', 'price': 6.8, 'aliases': []},
 {'code': '791', 'name': 'Irisimas channels, gulscias', 'price': 3.0, 'aliases': []},
 {'code': '796', 'name': 'Baltas minkstas virselis G/A', 'price': 5.5, 'aliases': []},
 {'code': '800', 'name': 'Baltas plastikine spirale 6–8 mm', 'price': 0.5, 'aliases': []},
 {'code': '801', 'name': 'Baltas plastikine spirale 10 mm', 'price': 0.6, 'aliases': []},
 {'code': '802', 'name': 'Baltas plastikine spirale 12 mm', 'price': 0.7, 'aliases': []},
 {'code': '803', 'name': 'Baltas plastikine spirale 14 mm', 'price': 0.8, 'aliases': []},
 {'code': '804', 'name': 'Plastikine spirale 16 mm, raudona arba juoda', 'price': 0.9, 'aliases': []},
 {'code': '805', 'name': 'Baltas plastikine spirale 19 mm', 'price': 1.0, 'aliases': []},
 {'code': '806', 'name': 'Baltas plastikine spirale 22–25 mm', 'price': 1.5, 'aliases': []},
 {'code': '807', 'name': 'Baltas plastikine spirale 28–32 mm', 'price': 2.2, 'aliases': []},
 {'code': '808', 'name': 'Baltas plastikine spirale 38–45 mm', 'price': 2.7, 'aliases': []},
 {'code': '809', 'name': 'Baltas plastikine spirale 50 mm arba didesne', 'price': 3.2, 'aliases': []},
 {'code': '810', 'name': 'Raudona plastikine spirale 6–8 mm', 'price': 0.5, 'aliases': []},
 {'code': '811', 'name': 'Raudona plastikine spirale 10 mm', 'price': 0.6, 'aliases': []},
 {'code': '812', 'name': 'Raudona plastikine spirale 12 mm', 'price': 0.7, 'aliases': []},
 {'code': '820', 'name': 'Juodas plastikine spirale 6–8 mm', 'price': 0.5, 'aliases': []},
 {'code': '821', 'name': 'Juodas plastikine spirale 10 mm', 'price': 0.6, 'aliases': []},
 {'code': '822', 'name': 'Juodas plastikine spirale 12 mm', 'price': 0.7, 'aliases': []},
 {'code': '824', 'name': 'Raudona arba juoda plastikine spirale 16 mm', 'price': 0.9, 'aliases': []},
 {'code': '825', 'name': 'Juodas plastikine spirale 19 mm', 'price': 1.0, 'aliases': []},
 {'code': '827', 'name': 'Juodas plastikine spirale 28–32 mm', 'price': 2.0, 'aliases': []},
 {'code': '828', 'name': 'Juodas plastikine spirale 38–45 mm', 'price': 2.5, 'aliases': []},
 {'code': '829', 'name': 'Juodas plastikine spirale 50 mm arba didesne', 'price': 3.2, 'aliases': []},
 {'code': '830', 'name': 'Melyna plastikine spirale 6–8 mm', 'price': 0.25, 'aliases': []},
 {'code': '831', 'name': 'Melyna plastikine spirale 10 mm', 'price': 0.3, 'aliases': []},
 {'code': '832', 'name': 'Melyna plastikine spirale 12 mm', 'price': 0.35, 'aliases': []},
 {'code': '840', 'name': 'Zalia plastikine spirale 6–8 mm', 'price': 0.25, 'aliases': []},
 {'code': '841', 'name': 'Zalia plastikine spirale 10 mm', 'price': 0.3, 'aliases': []},
 {'code': '842', 'name': 'Zalia plastikine spirale 12 mm', 'price': 0.3, 'aliases': []},
 {'code': '850', 'name': 'Raudona sukama spirale 22 mm', 'price': 3.2, 'aliases': []},
 {'code': '880', 'name': 'Irisimas plastikine spirale, 1–50 lapai', 'price': 3.0, 'aliases': []},
 {'code': '881', 'name': 'Irisimas plastikine spirale, 51–150 lapai', 'price': 3.5, 'aliases': []},
 {'code': '882', 'name': 'Irisimas plastikine spirale, 151–250 lapai', 'price': 4.5, 'aliases': []},
 {'code': '883', 'name': 'Irisimas plastikine spirale, 251–340 lapai', 'price': 5.5, 'aliases': []},
 {'code': '884', 'name': 'Irisimas plastikine spirale, 341+ lapai', 'price': 6.0, 'aliases': []},
 {'code': '891', 'name': 'A4 skaidrus irisimo virselis', 'price': 1.2, 'aliases': []},
 {'code': '892', 'name': 'A4 spalvotaed irisimo virselis', 'price': 1.2, 'aliases': []},
 {'code': '893', 'name': 'A4 matinis irisimo virselis', 'price': 1.2, 'aliases': []},
 {'code': '894', 'name': 'A3 skaidrus irisimo virselis', 'price': 2.0, 'aliases': []},
 {'code': '895', 'name': 'A4 irisimo galinis virselis', 'price': 1.2, 'aliases': []},
 {'code': '896', 'name': 'A3 irisimo galinis virselis', 'price': 2.0, 'aliases': []},
 {'code': '899', 'name': 'A4 dekoratyvinis popierius', 'price': 2.0, 'aliases': []},
 {'code': '900', 'name': 'Baltas sukama spirale 8 mm', 'price': 1.0, 'aliases': []},
 {'code': '901', 'name': 'Juodas sukama spirale 8 mm', 'price': 1.0, 'aliases': []},
 {'code': '903', 'name': 'Baltas sukama spirale 10 mm', 'price': 1.2, 'aliases': []},
 {'code': '904', 'name': 'Juodas sukama spirale 10 mm', 'price': 1.2, 'aliases': []},
 {'code': '905', 'name': 'A5–A6 laminavimas, 125 micron', 'price': 2.0, 'aliases': []},
 {'code': '906', 'name': 'A4 laminavimas, 100–125 micron', 'price': 2.5, 'aliases': []},
 {'code': '907', 'name': 'A3 laminavimas, 100–150 micron', 'price': 3.5, 'aliases': []},
 {'code': '908', 'name': 'A4 lipnus laminavimas, 100 micron', 'price': 3.5, 'aliases': []},
 {'code': '910', 'name': 'A4 thick laminavimas, 150–175 micron', 'price': 2.8, 'aliases': []},
 {'code': '914', 'name': '75 × 105 mm laminavimas, 175 micron', 'price': 1.5, 'aliases': []},
 {'code': '918', 'name': 'A3 matinis laminavimas, 125 micron', 'price': 4.2, 'aliases': []},
 {'code': '920', 'name': 'Single-hole punching', 'price': 0.25, 'aliases': []},
 {'code': '921', 'name': 'Two-hole punching, up to 50 lapai', 'price': 0.6, 'aliases': []},
 {'code': '922', 'name': 'Four-hole punching, up to 50 lapai', 'price': 0.8, 'aliases': []},
 {'code': '924', 'name': 'Corner rounding, one item', 'price': 0.6, 'aliases': []},
 {'code': '926', 'name': 'Stapling, 1–50 lapai', 'price': 0.8, 'aliases': []},
 {'code': '927', 'name': 'Stapling, 51–100 lapai', 'price': 1.5, 'aliases': []},
 {'code': '928', 'name': 'Stapling, 101–150 lapai', 'price': 2.0, 'aliases': []},
 {'code': '929', 'name': 'Metalines/plastikines kniedes ir panasus darbai', 'price': 2.5, 'aliases': []},
 {'code': '930', 'name': 'Kiti darbai', 'price': 1.0, 'aliases': []},
 {'code': '933', 'name': 'Pjaustymas, vienas pjuvis', 'price': 0.1, 'aliases': []},
 {'code': '934', 'name': 'Guillotine pjaustymas', 'price': 0.5, 'aliases': []},
 {'code': '935', 'name': 'Pristatymas pastu arba kurjeriu', 'price': 1.0, 'aliases': []},
 {'code': '936', 'name': 'Kniedes idejimas', 'price': 0.5, 'aliases': []},
 {'code': '941', 'name': 'Notary lipdukas', 'price': 0.5, 'aliases': []},
 {'code': '942', 'name': 'Notarinio dokumento surisimas', 'price': 1.2, 'aliases': []},
 {'code': '954', 'name': '80 gsm spalvotaed A4 popierius', 'price': 0.7, 'aliases': []},
 {'code': '956', 'name': '160 gsm spalvotaed A4 popierius', 'price': 0.8, 'aliases': []},
 {'code': '966', 'name': 'SRA3 lipnus lapas', 'price': 3.5, 'aliases': []},
 {'code': '988', 'name': '135–250 gsm A4 foto popierius', 'price': 1.0, 'aliases': []},
 {'code': '999', 'name': 'A4 skaidre', 'price': 2.0, 'aliases': []},
 {'code': '1001', 'name': 'Kietas irisimas, gulscias', 'price': 6.5, 'aliases': []},
 {'code': '1002', 'name': 'Skaidrus sukama spirale 10 mm', 'price': 1.2, 'aliases': []},
 {'code': '1003', 'name': 'Juodas sukama spirale 12 mm', 'price': 1.0, 'aliases': []},
 {'code': '1004', 'name': 'Baltas sukama spirale 12 mm', 'price': 1.3, 'aliases': []},
 {'code': '1006', 'name': 'A4/A5 document sleeve', 'price': 0.2, 'aliases': []},
 {'code': '1007', 'name': 'A3 document sleeve', 'price': 1.5, 'aliases': []},
 {'code': '1009', 'name': 'Folder with metal fastener', 'price': 2.0, 'aliases': []},
 {'code': '1011', 'name': 'Lever arch file, 50 mm', 'price': 6.0, 'aliases': []},
 {'code': '1012', 'name': 'Lever arch file, 70 mm', 'price': 7.0, 'aliases': []},
 {'code': '1015', 'name': 'Popierius CD sleeve', 'price': 0.5, 'aliases': []},
 {'code': '1016', 'name': 'Adhesive CD sleeve', 'price': 2.5, 'aliases': []},
 {'code': '1019', 'name': 'Standard A4 vokas', 'price': 1.0, 'aliases': []},
 {'code': '1020', 'name': 'Standard A5 vokas', 'price': 0.7, 'aliases': []},
 {'code': '1021', 'name': 'Standard A6 / long vokas', 'price': 0.6, 'aliases': []},
 {'code': '1025', 'name': 'Lever arch file, 80 mm', 'price': 8.0, 'aliases': []},
 {'code': '1033', 'name': 'Ring binder', 'price': 6.5, 'aliases': []},
 {'code': '1042', 'name': 'Small glue stick, 8 g', 'price': 3.0, 'aliases': []},
 {'code': '1044', 'name': 'Scissors', 'price': 4.5, 'aliases': []},
 {'code': '1062', 'name': 'Korteleboard vizitiniu korteliu dezute', 'price': 0.5, 'aliases': []},
 {'code': '1066', 'name': 'Skaidrus sukama spirale 12 mm', 'price': 1.2, 'aliases': []},
 {'code': '1067', 'name': 'Juodas sukama spirale 14 mm', 'price': 1.5, 'aliases': []},
 {'code': '1068', 'name': 'Baltas sukama spirale 14 mm', 'price': 1.4, 'aliases': []},
 {'code': '1069', 'name': 'Juodas sukama spirale 16 mm', 'price': 1.8, 'aliases': []},
 {'code': '1070', 'name': 'Baltas sukama spirale 16 mm', 'price': 1.7, 'aliases': []},
 {'code': '1071', 'name': 'Juodas sukama spirale 20 mm', 'price': 2.7, 'aliases': []},
 {'code': '1072', 'name': 'Baltas sukama spirale 20 mm', 'price': 2.7, 'aliases': []},
 {'code': '1073', 'name': 'Juodas sukama spirale 25 mm', 'price': 3.7, 'aliases': []},
 {'code': '1074', 'name': 'Baltas sukama spirale 25 mm', 'price': 3.6, 'aliases': []},
 {'code': '1075', 'name': 'Juodas sukama spirale 32 mm', 'price': 5.0, 'aliases': []},
 {'code': '1076', 'name': 'Baltas sukama spirale 32 mm', 'price': 5.2, 'aliases': []},
 {'code': '1077', 'name': 'Juodas sukama spirale 35 mm', 'price': 5.0, 'aliases': []},
 {'code': '1079', 'name': 'Juodas sukama spirale 51 mm', 'price': 7.0, 'aliases': []},
 {'code': '1081', 'name': 'Melyna sukama spirale 8 mm', 'price': 0.7, 'aliases': []},
 {'code': '1082', 'name': 'Raudona sukama spirale 8 mm', 'price': 1.0, 'aliases': []},
 {'code': '1083', 'name': 'Melyna sukama spirale 10 mm', 'price': 0.9, 'aliases': []},
 {'code': '1084', 'name': 'Raudona sukama spirale 10 mm', 'price': 1.3, 'aliases': []},
 {'code': '1085', 'name': 'Melyna sukama spirale 12 mm', 'price': 0.9, 'aliases': []},
 {'code': '1086', 'name': 'Raudona sukama spirale 12 mm', 'price': 1.5, 'aliases': []},
 {'code': '1087', 'name': 'Raudona sukama spirale 14 mm', 'price': 1.6, 'aliases': []},
 {'code': '1088', 'name': 'Raudona sukama spirale 16 mm', 'price': 1.8, 'aliases': []},
 {'code': '1090', 'name': '350–400 gsm A3/SRA3 foto popierius', 'price': 2.5, 'aliases': []},
 {'code': '1092', 'name': 'Teksturinis A3/SRA3 popierius', 'price': 2.5, 'aliases': []},
 {'code': '1095', 'name': 'Ilgaamzis neplystantis A3 popierius', 'price': 3.5, 'aliases': []},
 {'code': '1096', 'name': 'Magnetinis SRA3 popierius', 'price': 3.2, 'aliases': []},
 {'code': '1103', 'name': 'Instant glue', 'price': 3.5, 'aliases': []},
 {'code': '1109', 'name': 'Dvipuse lipni juosta', 'price': 7.5, 'aliases': []},
 {'code': '1113', 'name': 'Standard pencil', 'price': 1.5, 'aliases': []},
 {'code': '1117', 'name': 'Korteleboard box', 'price': 2.9, 'aliases': []},
 {'code': '1118', 'name': 'Automatic pen', 'price': 2.5, 'aliases': []},
 {'code': '1120', 'name': 'Wide-format tube, narrow', 'price': 3.0, 'aliases': []},
 {'code': '1124', 'name': 'Thick lipnus tape', 'price': 5.0, 'aliases': []},
 {'code': '1137', 'name': 'Colouraudona push pins', 'price': 1.5, 'aliases': []},
 {'code': '1144', 'name': 'Staples 24/6', 'price': 2.5, 'aliases': []},
 {'code': '1149', 'name': 'A4 remelis', 'price': 10.0, 'aliases': []},
 {'code': '1153', 'name': 'A4 remelis with thick border', 'price': 10.0, 'aliases': []},
 {'code': '1164', 'name': 'Popierius clips', 'price': 2.0, 'aliases': []},
 {'code': '1173', 'name': 'Mounting squares', 'price': 7.5, 'aliases': []},
 {'code': '1182', 'name': 'Popierius bags', 'price': 2.0, 'aliases': []},
 {'code': '1183', 'name': 'Brown kraft A4 vokas', 'price': 1.0, 'aliases': []},
 {'code': '1184', 'name': 'Pioneer Navigator A4 popieriaus pakuote', 'price': 7.5, 'aliases': []},
 {'code': '1214', 'name': 'Large spalvotaed lapai', 'price': 4.5, 'aliases': []},
 {'code': '1216', 'name': '40 × 60 (A2) arba 40 × 50 remelis', 'price': 28.0, 'aliases': []},
 {'code': '1219', 'name': '13 × 18 / A5 remelis', 'price': None, 'aliases': []},
 {'code': '1220', 'name': '15 × 21 (A5) remelis with thick border', 'price': 9.0, 'aliases': []},
 {'code': '1223', 'name': 'A3 remelis', 'price': 15.0, 'aliases': []},
 {'code': '1237', 'name': 'Staple remover', 'price': None, 'aliases': []},
 {'code': '1240', 'name': 'Small hole punch', 'price': 6.0, 'aliases': []},
 {'code': '1244', 'name': 'PVA glue', 'price': 3.0, 'aliases': []},
 {'code': '1245', 'name': 'Thin lipnus tape', 'price': 2.0, 'aliases': []},
 {'code': '1249', 'name': 'Correction fluid', 'price': 3.0, 'aliases': []},
 {'code': '1253', 'name': 'Pack of binder clips', 'price': 3.0, 'aliases': []},
 {'code': '1268', 'name': '16 GB USB atmintine', 'price': 20.0, 'aliases': []},
 {'code': '1270', 'name': 'KopijaPro pen', 'price': 2.0, 'aliases': []},
 {'code': '1274', 'name': 'Markers', 'price': 3.5, 'aliases': []},
 {'code': '1275', 'name': 'Adhesive tape with dispenser', 'price': 3.0, 'aliases': []},
 {'code': '1284', 'name': 'Postkorteles', 'price': 2.0, 'aliases': []},
 {'code': '1289', 'name': 'Binder clip', 'price': 0.5, 'aliases': []},
 {'code': '1293', 'name': '8 GB USB atmintine', 'price': 18.0, 'aliases': []},
 {'code': '1294', 'name': 'Padded vokas', 'price': None, 'aliases': []},
 {'code': '1296', 'name': 'Large hole punch', 'price': 15.0, 'aliases': []},
 {'code': '1300', 'name': 'Spausdinimaser cartridges', 'price': 1.0, 'aliases': []},
 {'code': '1301', 'name': 'Canvas', 'price': 1.0, 'aliases': []},
 {'code': '1302', 'name': 'Stamps', 'price': 1.0, 'aliases': []},
 {'code': '1303', 'name': 'Stamp rubber', 'price': 1.0, 'aliases': []},
 {'code': '1305', 'name': 'Window boxes', 'price': 5.5, 'aliases': []},
 {'code': '1317', 'name': 'Glossy A4 vokas', 'price': 3.5, 'aliases': []},
 {'code': '1318', 'name': 'Sticky notes', 'price': 1.5, 'aliases': []},
 {'code': '1319', 'name': 'A4 Curious Collection popierius', 'price': 2.0, 'aliases': []},
 {'code': '1327', 'name': 'Gift bags', 'price': 3.0, 'aliases': []},
 {'code': '1345', 'name': '32 GB USB atmintine', 'price': 22.0, 'aliases': []},
 {'code': '1349', 'name': 'Marker set', 'price': 4.0, 'aliases': []},
 {'code': '1350', 'name': 'Wide-format tube, wide', 'price': 4.5, 'aliases': []},
 {'code': '1351', 'name': '10 × 15 (A6) remelis', 'price': 5.0, 'aliases': []},
 {'code': '1354', 'name': 'Spausdinimasing on a bag', 'price': 8.0, 'aliases': []},
 {'code': '1355', 'name': 'Spausdinimasing on a T-shirt', 'price': 18.0, 'aliases': []},
 {'code': '1360', 'name': 'A0 remelis', 'price': 65.0, 'aliases': []},
 {'code': '1362', 'name': 'Teksturinis A5 vokas', 'price': 2.5, 'aliases': []},
 {'code': '1364', 'name': 'Pack of 20 document sleeves', 'price': 3.0, 'aliases': []},
 {'code': '1991', 'name': 'Jigsaw puzzle production (price varies by puzzle)', 'price': None, 'aliases': []},
 {'code': '1999', 'name': '10 × 15 foto spausdinimas', 'price': 0.6, 'aliases': []},
 {'code': '2000', 'name': 'A4 foto spausdinimas', 'price': 2.5, 'aliases': []},
 {'code': '2001', 'name': 'Document foto', 'price': 5.0, 'aliases': []},
 {'code': '2003', 'name': 'Spausdinimasing on a mug', 'price': 1.0, 'aliases': []},
 {'code': '2006', 'name': '60 × 90 remelis', 'price': 30.0, 'aliases': []},
 {'code': '2007', 'name': '50 × 70 remelis', 'price': 28.0, 'aliases': []},
 {'code': '2009', 'name': 'Satin ribbon', 'price': 5.0, 'aliases': []},
 {'code': '2012', 'name': 'A1 remelis', 'price': 32.0, 'aliases': []},
 {'code': '2022', 'name': 'A3 skaidre', 'price': 3.0, 'aliases': []},
 {'code': '2025', 'name': 'Large blizgus spalvotaed lapai', 'price': 4.0, 'aliases': []},
 {'code': '2029', 'name': 'Gift ribbon', 'price': 2.5, 'aliases': []},
 {'code': '2032', 'name': 'Memo note lapai', 'price': 5.0, 'aliases': []},
 {'code': '2033', 'name': 'A5/A6 information holder', 'price': 6.5, 'aliases': []},
 {'code': '2034', 'name': 'A4 information holder', 'price': 10.0, 'aliases': []},
 {'code': '2047', 'name': 'Skaidrus bags, 100 pcs', 'price': 8.0, 'aliases': []},
 {'code': '2052', 'name': 'Adhesive ribbon', 'price': 2.0, 'aliases': []},
 {'code': '2054', 'name': '64 GB USB atmintine', 'price': 26.0, 'aliases': []},
 {'code': '2059', 'name': 'Small calendars', 'price': 1.0, 'aliases': []},
 {'code': '2062', 'name': 'Navigator popierius, 150 lapai', 'price': 3.0, 'aliases': []},
 {'code': '2070', 'name': 'KopijaPro A6 / 14 × 18 vokass', 'price': 2.5, 'aliases': []},
 {'code': '2074', 'name': 'Decorative elements', 'price': 5.0, 'aliases': []},
 {'code': '2079', 'name': 'KopijaPro C65 vokass', 'price': 2.5, 'aliases': []},
 {'code': '2080', 'name': 'KopijaPro A5 vokass', 'price': 3.5, 'aliases': []},
 {'code': '2081', 'name': 'KopijaPro A4 vokass', 'price': 4.5, 'aliases': []},
 {'code': '2088', 'name': 'A0 remelis backing board', 'price': 24.0, 'aliases': []},
 {'code': '2089', 'name': 'A1 remelis backing board', 'price': 12.0, 'aliases': []},
 {'code': '2090', 'name': 'A2 remelis backing board', 'price': 6.0, 'aliases': []},
 {'code': '2091', 'name': 'A3 remelis backing board', 'price': 3.0, 'aliases': []},
 {'code': '2851', 'name': 'SRA3 transfer popierius, 350 gsm', 'price': 3.1, 'aliases': []},
 {'code': '3008', 'name': 'Vatmanas 90×64 cm (240 gsm)', 'price': 3.0, 'aliases': []},
 {'code': '3009', 'name': 'Spausdinimas ant vatmano', 'price': 14.0, 'aliases': []},
 {'code': '4800', 'name': '80 gsm paprastas / brėžinys', 'price': 0.95, 'aliases': []},
 {'code': '4810', 'name': '80 gsm paprastas / tekstas+pav.', 'price': 1.4, 'aliases': []},
 {'code': '4820', 'name': '80 gsm paprastas / paveikslas', 'price': 1.6, 'aliases': []},
 {'code': '4830', 'name': '120 gsm storesnis / tekstas+pav.', 'price': 1.7, 'aliases': []},
 {'code': '4840', 'name': '120 gsm storesnis / paveikslas', 'price': 2.1, 'aliases': []},
 {'code': '4850', 'name': '180 gsm storas / tekstas+paveikslas', 'price': 2.4, 'aliases': []},
 {'code': '4860', 'name': '180 gsm storas / paveikslas', 'price': 2.55, 'aliases': []},
 {'code': '4893', 'name': 'Satin / fotopopierius / paveikslas', 'price': 3.2, 'aliases': []},
 {'code': '4894', 'name': 'Satin / fotopopierius / tekstas+pav.', 'price': 2.8, 'aliases': []},
 {'code': '4900', 'name': 'Natūrali drobė', 'price': 3.3, 'aliases': []},
 {'code': '4902', 'name': 'Ple vele plėvelė / tekstas+paveikslas', 'price': 2.9, 'aliases': []},
 {'code': '4903', 'name': 'Ple vele plėvelė / paveikslas', 'price': 3.3, 'aliases': []},
 {'code': '4908', 'name': 'Lipdukas PVC / tekstas+paveikslas', 'price': 3.0, 'aliases': []},
 {'code': '4909', 'name': 'Lipdukas PVC / paveikslas', 'price': 3.2, 'aliases': []},
 {'code': '4910', 'name': 'Kalkė / brėžinys', 'price': 1.7, 'aliases': []},
 {'code': '4911', 'name': 'Kalkė / tekstas+paveikslas', 'price': 2.2, 'aliases': []},
 {'code': '4912', 'name': 'Kalkė / paveikslas', 'price': 2.5, 'aliases': []},
 {'code': '4913', 'name': '140 gsm storas / tekstas+paveikslas', 'price': 2.35, 'aliases': []},
 {'code': '4914', 'name': '140 gsm storas / paveikslas', 'price': 2.3, 'aliases': []},
 {'code': '4915', 'name': 'Sintetinė drobė', 'price': 2.8, 'aliases': []},
 {'code': '4916', 'name': 'Karštas laminatas', 'price': 1.8, 'aliases': []},
 {'code': '4917', 'name': 'Magnetinė plėvelė', 'price': 2.5, 'aliases': []},
 {'code': '8800', 'name': 'Reirisimo', 'price': 5.0, 'aliases': []},
 {'code': '9302', 'name': 'Plotting', 'price': 1.0, 'aliases': []},
 {'code': '9310', 'name': 'Cup/trophy item', 'price': 1.0, 'aliases': []},
 {'code': '9311', 'name': 'Medals', 'price': 1.0, 'aliases': []},
 {'code': '9312', 'name': 'One-sided vizitines korteles maketavimas', 'price': 25.0, 'aliases': []},
 {'code': '9313', 'name': 'Double-sided vizitines korteles maketavimas', 'price': 35.0, 'aliases': []},
 {'code': '9314', 'name': 'Badge production', 'price': 3.0, 'aliases': []},
 {'code': '9315', 'name': 'Keychain production', 'price': 3.0, 'aliases': []},
 {'code': '9316', 'name': 'Magnet production', 'price': 2.0, 'aliases': []},
 {'code': '9400', 'name': 'Notarial thread', 'price': 0.7, 'aliases': []},
 {'code': '9502', 'name': '80 gsm A4 popierius', 'price': 0.15, 'aliases': []},
 {'code': '9503', 'name': 'A4 self-service popierius, one lapas', 'price': 0.06, 'aliases': []},
 {'code': '9504', 'name': '80 gsm popieriaus pakuote', 'price': 5.5, 'aliases': []},
 {'code': '9512', 'name': '80 gsm A3 popierius', 'price': 0.2, 'aliases': []},
 {'code': '9513', 'name': 'A3 self-service popierius, one lapas', 'price': 0.07, 'aliases': []},
 {'code': '9582', 'name': 'A4 transfer popierius', 'price': 1.0, 'aliases': []},
 {'code': '9592', 'name': 'A3 transfer popierius', 'price': 1.5, 'aliases': []},
 {'code': '9822', 'name': '160 gsm A4 popierius', 'price': 0.5, 'aliases': []},
 {'code': '9832', 'name': '160 gsm A3 popierius', 'price': 0.8, 'aliases': []},
 {'code': '9842', 'name': '200 gsm A4 popierius', 'price': 0.65, 'aliases': []},
 {'code': '9852', 'name': '200 gsm A3 popierius', 'price': 0.95, 'aliases': []},
 {'code': '9862', 'name': '250–300 gsm A4 popierius', 'price': 1.1, 'aliases': []},
 {'code': '9872', 'name': '250–300 gsm A3 popierius', 'price': 2.1, 'aliases': []},
 {'code': '9922', 'name': 'A4 lipnus popierius', 'price': 1.5, 'aliases': []},
 {'code': '9972', 'name': '120 gsm A4 popierius', 'price': 0.4, 'aliases': []},
 {'code': '9982', 'name': '120 gsm A3 popierius', 'price': 0.6, 'aliases': []},
 {'code': '1999x2', 'name': 'A5 foto spausdinimas', 'price': 1.2, 'aliases': []}]
_EMBEDDED_PAPER_SIZES = [{'name': 'A0', 'width_mm': 841, 'height_mm': 1189, 'custom': False},
 {'name': 'A1', 'width_mm': 594, 'height_mm': 841, 'custom': False},
 {'name': 'A2', 'width_mm': 420, 'height_mm': 594, 'custom': False},
 {'name': 'A3', 'width_mm': 297, 'height_mm': 420, 'custom': False},
 {'name': 'A4', 'width_mm': 210, 'height_mm': 297, 'custom': False},
 {'name': 'A5', 'width_mm': 148, 'height_mm': 210, 'custom': False},
 {'name': 'A6', 'width_mm': 105, 'height_mm': 148, 'custom': False},
 {'name': 'A7', 'width_mm': 74, 'height_mm': 105, 'custom': False},
 {'name': 'SRA3', 'width_mm': 320, 'height_mm': 450, 'custom': False},
 {'name': 'SRA4', 'width_mm': 225, 'height_mm': 320, 'custom': False},
 {'name': '10×15 cm (102×152)', 'width_mm': 102, 'height_mm': 152, 'custom': False},
 {'name': '21×15 cm (210×152)', 'width_mm': 210, 'height_mm': 152, 'custom': False},
 {'name': 'Square 10×10', 'width_mm': 100, 'height_mm': 100, 'custom': False},
 {'name': 'Square 15×15', 'width_mm': 150, 'height_mm': 150, 'custom': False}]
_EMBEDDED_WIDE_ITEMS = [{'code': '4800', 'name': '80 gsm paprastas / brėžinys', 'price': 0.95, 'coverage': 'drawing'},
 {'code': '4810', 'name': '80 gsm paprastas / tekstas+pav.', 'price': 1.4, 'coverage': 'partial'},
 {'code': '4820', 'name': '80 gsm paprastas / paveikslas', 'price': 1.6, 'coverage': 'full'},
 {'code': '4830', 'name': '120 gsm storesnis / tekstas+pav.', 'price': 1.7, 'coverage': 'partial'},
 {'code': '4840', 'name': '120 gsm storesnis / paveikslas', 'price': 2.1, 'coverage': 'full'},
 {'code': '4913', 'name': '140 gsm storas / tekstas+paveikslas', 'price': 2.35, 'coverage': 'partial'},
 {'code': '4914', 'name': '140 gsm storas / paveikslas', 'price': 2.3, 'coverage': 'full'},
 {'code': '4850', 'name': '180 gsm storas / tekstas+paveikslas', 'price': 2.4, 'coverage': 'partial'},
 {'code': '4860', 'name': '180 gsm storas / paveikslas', 'price': 2.55, 'coverage': 'full'},
 {'code': '4894', 'name': 'Satin / fotopopierius / tekstas+pav.', 'price': 2.8, 'coverage': 'partial'},
 {'code': '4893', 'name': 'Satin / fotopopierius / paveikslas', 'price': 3.2, 'coverage': 'full'},
 {'code': '4915', 'name': 'Sintetinė drobė', 'price': 2.8, 'coverage': 'fixed'},
 {'code': '4900', 'name': 'Natūrali drobė', 'price': 3.3, 'coverage': 'fixed'},
 {'code': '4902', 'name': 'Film plėvelė / tekstas+paveikslas', 'price': 2.9, 'coverage': 'partial'},
 {'code': '4903', 'name': 'Film plėvelė / paveikslas', 'price': 3.3, 'coverage': 'full'},
 {'code': '4908', 'name': 'Lipdukas PVC / tekstas+paveikslas', 'price': 3.0, 'coverage': 'partial'},
 {'code': '4909', 'name': 'Lipdukas PVC / paveikslas', 'price': 3.2, 'coverage': 'full'},
 {'code': '4910', 'name': 'Kalkė / brėžinys', 'price': 1.7, 'coverage': 'drawing'},
 {'code': '4911', 'name': 'Kalkė / tekstas+paveikslas', 'price': 2.2, 'coverage': 'partial'},
 {'code': '4912', 'name': 'Kalkė / paveikslas', 'price': 2.5, 'coverage': 'full'},
 {'code': '3009', 'name': 'Spausdinimas ant vatmano', 'price': 14.0, 'coverage': 'fixed'},
 {'code': '4916', 'name': 'Karštas laminatas', 'price': 1.8, 'coverage': 'fixed'},
 {'code': '4917', 'name': 'Magnetinė plėvelė', 'price': 2.5, 'coverage': 'fixed'},
 {'code': '3008', 'name': 'Vatmanas 90×64 cm (240 gsm)', 'price': 3.0, 'coverage': 'fixed'}]

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
    if os.path.exists(bundled): shutil.copy2(bundled,destination)

def _ensure_editable_files():
    paths=_editable_paths()
    _copy_default_if_missing(paths["codes_file"],"copypro_kodai.csv")
    _copy_default_if_missing(paths["paper_sizes_file"],"copypro_popieriaus_dydziai.csv")

def _csv_rows(path):
    with open(path,"r",encoding="utf-8-sig",newline="") as f:
        return list(csv.DictReader(f))

def reload_editable_data():
    global PAPER_SIZES_MM, WIDE_FORMAT_ITEMS, PLU_SEARCH_ALIASES
    _ensure_editable_files(); paths=_editable_paths()
    PLU_DESCRIPTIONS.clear(); PLU_PRICES.clear(); WIDE_FORMAT_ITEMS={}; PLU_SEARCH_ALIASES={}
    coverage_to_internal={"dalinis":"partial","pilnas":"full","brėžinys":"drawing","brezinys":"drawing","fiksuotas":"fixed"}
    for item in _csv_rows(paths["codes_file"]):
        if str(item.get("active","taip")).strip().lower() not in ("taip","yes","1","true"): continue
        code=str(item.get("code","")).strip(); name=str(item.get("name_lt","")).strip()
        if not code or not name: continue
        PLU_DESCRIPTIONS[code]=name
        PLU_SEARCH_ALIASES[code]=str(item.get("search_aliases_lt","")).strip()
        raw=str(item.get("price_eur","")).strip().replace(",",".")
        if raw:
            try: PLU_PRICES[code]=float(raw)
            except ValueError: pass
        raw_coverage=str(item.get("wide_format_coverage","")).strip().lower()
        coverage=coverage_to_internal.get(raw_coverage,raw_coverage)
        label=str(item.get("wide_format_label_lt","")).strip()
        if coverage or label:
            WIDE_FORMAT_ITEMS[code]=(name,PLU_PRICES.get(code,0.0),coverage or "fixed",label or name)
    PAPER_SIZES_MM={}
    for item in _csv_rows(paths["paper_sizes_file"]):
        if str(item.get("active","taip")).strip().lower() not in ("taip","yes","1","true"): continue
        if str(item.get("is_custom","ne")).strip().lower() in ("taip","yes","1","true"): continue
        try: PAPER_SIZES_MM[str(item["name"]).strip()]=(float(str(item["width_mm"]).replace(",",".")),float(str(item["height_mm"]).replace(",",".")))
        except Exception: continue

def load_custom_sizes():
    try:
        result={}
        for item in _csv_rows(_editable_paths()["paper_sizes_file"]):
            if str(item.get("is_custom","ne")).strip().lower() not in ("taip","yes","1","true"): continue
            result[str(item["name"]).strip()]=(float(str(item["width_mm"]).replace(",",".")),float(str(item["height_mm"]).replace(",",".")))
        return result
    except Exception:return {}

def save_custom_sizes(sizes_dict):
    path=_editable_paths()["paper_sizes_file"]
    try: rows=_csv_rows(path)
    except Exception: rows=[]
    rows=[x for x in rows if str(x.get("is_custom","ne")).strip().lower() not in ("taip","yes","1","true")]
    for name,(w,h) in sizes_dict.items(): rows.append({"group":"Pasirinktiniai","name":name,"width_mm":w,"height_mm":h,"active":"taip","is_custom":"taip","notes_lt":""})
    fields=["group","name","width_mm","height_mm","active","is_custom","notes_lt"]
    with open(path,"w",encoding="utf-8-sig",newline="") as f:
        writer=csv.DictWriter(f,fieldnames=fields);writer.writeheader();writer.writerows(rows)

reload_editable_data()

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

    def _build(self):
        toolbar = tk.Frame(self, bg=SURFACE, pady=8)
        toolbar.pack(fill="x")
        styled_btn(toolbar, "➕  Add Row", self._add_row).pack(side="left", padx=12)
        styled_btn(toolbar, "🗑  Clear All", self._clear_all, style="secondary").pack(side="left", padx=4)
        styled_btn(toolbar, "📋  Copy Summary", self._copy_summary, style="secondary").pack(side="left", padx=4)
        styled_btn(toolbar, "↔  Wide format quote", self._open_wide_format_quote, style="secondary").pack(side="left", padx=4)
        tk.Label(toolbar,
                 text="Enter amount*code, for example 8*710 — Enter adds the next row",
                 bg=SURFACE, fg=MUTED, font=("Segoe UI", 8)).pack(side="right", padx=12)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)

        left = tk.Frame(body, bg=BG, width=480)
        left.pack(side="left", fill="both", expand=False)
        left.pack_propagate(False)

        search_box = tk.Frame(left, bg=SURFACE, padx=12, pady=10)
        search_box.pack(fill="x")
        tk.Label(search_box, text="SEARCH KNOWN ITEMS", bg=SURFACE, fg=MUTED,
                 font=("Segoe UI", 8, "bold"), anchor="w").pack(fill="x", pady=(0, 4))
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_box, textvariable=self.search_var,
            bg=SURFACE2, fg=TEXT, insertbackground=TEXT, relief="flat",
            font=("Segoe UI", 10))
        self.search_entry.pack(fill="x", ipady=5)
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
            for code, (_description, price, item_coverage, item_label) in WIDE_FORMAT_ITEMS.items():
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

    def _open_data_settings(self):
        popup=tk.Toplevel(self); popup.title("Duomenu failu nustatymai"); popup.configure(bg=BG); popup.geometry("690x300"); popup.grab_set()
        paths=_editable_paths(); vars_={k:tk.StringVar(value=v) for k,v in paths.items()}
        labels=(("codes_file","Prekiu ir paslaugu kodai"),("paper_sizes_file","Popieriaus dydziai"))
        lbl(popup,"DUOMENU FAILAI",11,TEXT,True,bg=BG).pack(anchor="w",padx=18,pady=(16,10))
        body=tk.Frame(popup,bg=BG); body.pack(fill="both",expand=True,padx=18)
        def browse(key):
            p=filedialog.askopenfilename(parent=popup,title="Pasirinkti CSV faila",filetypes=[("CSV","*.csv"),("Visi failai","*.*")])
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
                reload_editable_data()
                for tab in self.tabs.values():
                    if hasattr(tab,"on_show"): tab.on_show()
                messagebox.showinfo("Issaugota","Failu vietos issaugotos ir duomenys perkrauti.",parent=popup); popup.destroy()
            except Exception as ex: messagebox.showerror("Klaida",str(ex),parent=popup)
        footer=tk.Frame(popup,bg=BG); footer.pack(fill="x",padx=18,pady=14)
        styled_btn(footer,"Atidaryti AppData aplanka",lambda:subprocess.Popen(["explorer",COPYPRO_DATA_DIR]) if sys.platform.startswith("win") else None,style="secondary").pack(side="left")
        styled_btn(footer,"Tikrinti atnaujinimus",lambda:self._check_for_updates(manual=True),style="secondary").pack(side="left",padx=8)
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
            ("pdf",       "📄  PDF Optimiser",      PdfToolsTab),
            ("pages",     "📊  Page Counter",       PageCounterTab),
            ("organiser", "🗃  Batch Organiser",    BatchOrganiserTab),
            ("counter",   "🖨  Print Counter",      PrintCounterTab),
        ]
        self.tabs = {}; self.tab_btns = {}
        for key,_,cls in tab_defs:
            self.tabs[key] = cls(content)

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
