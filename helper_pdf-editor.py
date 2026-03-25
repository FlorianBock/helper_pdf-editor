"""
Helper: PDF Editor
==================
A GUI tool that lets you:
  • Click anywhere on a PDF page to place a text label
  • Drag the grip handle (⠿) to reposition text
  • Click × to delete a text item
  • Adjust font size and bold before placing text
  • Save the result — text is burned permanently into a new PDF

Also supports interactive AcroForm fields when present.

Requirements:  pip install PyMuPDF Pillow
Run:           python helper_pdf-editor.py
"""

import dataclasses
import ctypes
import io
import os
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from tkinter import filedialog, messagebox, ttk
from typing import Optional

try:
    from _version import VERSION, BUILD_DATE
except ImportError:
    VERSION    = "1.0"
    BUILD_DATE = None

try:
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk
except ImportError as exc:
    import sys
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "Missing Dependencies",
        f"Required package not found: {exc}\n\n"
        "Please install them:\n    pip install PyMuPDF Pillow",
    )
    sys.exit(1)


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclasses.dataclass
class TextPlacement:
    """One piece of free-form text placed on a PDF page."""
    page_idx: int
    x_pdf: float        # PDF-space coords (top-left origin, y downward)
    y_pdf: float
    text: str
    font_size: float
    bold: bool = False
    kind: str = "text"       # "text" or "check"
    width_pdf: float = 120.0  # widget width in PDF units (ignored for kind=check)
    # Runtime UI state — not persisted to disk
    canvas_win_id: int = 0
    frame: Optional[tk.Frame] = None
    var: Optional[tk.StringVar] = None
    entry: Optional[tk.Entry] = None


@dataclasses.dataclass
class EraserRect:
    """A white filled rectangle (or oval/polygon) burned over a region to erase content."""
    page_idx: int
    x0: float; y0: float; x1: float; y1: float
    kind: str = "rect"         # "rect" | "circle" | "lasso" | "free"
    points: list = dataclasses.field(default_factory=list)  # PDF-space points for lasso/free
    pencil_size: float = 10.0  # stroke width in PDF points (free mode only)


@dataclasses.dataclass
class ShapeRect:
    """A geometric shape drawn on a PDF page (rect, rounded_rect, ellipse, line, arrow)."""
    page_idx: int
    x0: float; y0: float; x1: float; y1: float   # PDF-space bounding box
    kind: str = "rect"          # "rect" | "rounded_rect" | "ellipse" | "line" | "arrow"
    stroke_color: tuple = (0.0, 0.0, 0.0)         # RGB 0-1
    fill_color: tuple | None = None               # None = no fill
    line_width: float = 1.5
    # Runtime: canvas item id for the preview overlay
    canvas_id: int = 0


@dataclasses.dataclass
class RegionClip:
    """A raster copy of a page region that can be pasted back as an image."""
    page_idx: int
    x_pdf: float; y_pdf: float          # paste destination (PDF-space top-left)
    w_pdf: float; h_pdf: float          # dimensions in PDF units
    image: object = dataclasses.field(repr=False, default=None)   # PIL.Image
    # Runtime
    canvas_win_id: int = 0
    frame: object = dataclasses.field(repr=False, default=None)   # tk.Frame


def _field_is_checked(value: str) -> bool:
    return str(value).lower() in ("yes", "true", "on", "1", "/yes")


def _sep(parent: tk.Widget) -> None:
    ttk.Separator(parent, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=2)


# ---------------------------------------------------------------------------
# Printing helpers (pure ctypes / GDI32 -- no third-party packages required)
# ---------------------------------------------------------------------------

_gdi32    = ctypes.windll.gdi32
_winspool = ctypes.WinDLL("winspool.drv")

try:
    from tkinterdnd2 import DND_FILES as _DND_FILES, TkinterDnD as _TkinterDnD
    _AppBase = _TkinterDnD.Tk
except ImportError:
    _AppBase = tk.Tk
    _DND_FILES = None

# GDI GetDeviceCaps indices
_HORZRES    = 8
_VERTRES    = 10
_LOGPIXELSX = 88
_LOGPIXELSY = 90

# StretchDIBits constants
_BI_RGB         = 0
_DIB_RGB_COLORS = 0
_SRCCOPY        = 0x00CC0020


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize",          ctypes.c_uint32),
        ("biWidth",         ctypes.c_int32),
        ("biHeight",        ctypes.c_int32),
        ("biPlanes",        ctypes.c_uint16),
        ("biBitCount",      ctypes.c_uint16),
        ("biCompression",   ctypes.c_uint32),
        ("biSizeImage",     ctypes.c_uint32),
        ("biXPelsPerMeter", ctypes.c_int32),
        ("biYPelsPerMeter", ctypes.c_int32),
        ("biClrUsed",       ctypes.c_uint32),
        ("biClrImportant",  ctypes.c_uint32),
    ]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", _BITMAPINFOHEADER),
        ("bmiColors", ctypes.c_uint32 * 1),
    ]


class _DOCINFOW(ctypes.Structure):
    _fields_ = [
        ("cbSize",       ctypes.c_int),
        ("lpszDocName",  ctypes.c_wchar_p),
        ("lpszOutput",   ctypes.c_wchar_p),
        ("lpszDatatype", ctypes.c_wchar_p),
        ("fwType",       ctypes.c_uint),
    ]


class _PRINTER_INFO_4(ctypes.Structure):
    _fields_ = [
        ("pPrinterName", ctypes.c_wchar_p),
        ("pServerName",  ctypes.c_wchar_p),
        ("Attributes",   ctypes.c_uint),
    ]


def _find_unicode_font(bold: bool) -> str | None:
    """Return a path to a Windows TrueType font that covers the full Unicode BMP.

    Tries several common system font families in order.  Returns None when no
    suitable file is found (caller falls back to the PDF built-in fonts).
    """
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    candidates = (
        ("arialbd.ttf",   "arial.ttf"  ),   # Arial
        ("calibrib.ttf",  "calibri.ttf" ),   # Calibri
        ("verdanab.ttf",  "verdana.ttf" ),   # Verdana
        ("tahomabd.ttf",  "tahoma.ttf"  ),   # Tahoma
        ("segoeui.ttf",   "segoeui.ttf" ),   # Segoe UI (no separate bold in older Windows)
    )
    for bold_file, reg_file in candidates:
        name = bold_file if bold else reg_file
        path = os.path.join(fonts_dir, name)
        if os.path.isfile(path):
            return path
    return None


def _get_printers() -> list:
    """Return available printer display names via winspool, falling back to PowerShell."""
    PRINTER_ENUM_LOCAL       = 0x02
    PRINTER_ENUM_CONNECTIONS = 0x04
    flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS
    needed   = ctypes.c_ulong(0)
    returned = ctypes.c_ulong(0)
    # First call: get required buffer size
    _winspool.EnumPrintersW(flags, None, 4, None, 0,
                             ctypes.byref(needed), ctypes.byref(returned))
    if needed.value:
        buf = (ctypes.c_byte * needed.value)()
        if _winspool.EnumPrintersW(flags, None, 4, buf, needed.value,
                                   ctypes.byref(needed), ctypes.byref(returned)):
            entry_size = ctypes.sizeof(_PRINTER_INFO_4)
            names = []
            for i in range(returned.value):
                info = _PRINTER_INFO_4.from_buffer(buf, i * entry_size)
                if info.pPrinterName:
                    names.append(info.pPrinterName)
            if names:
                return names
    # Fallback: PowerShell
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command",
             "Get-Printer | Select-Object -ExpandProperty Name"],
            text=True, timeout=10, stderr=subprocess.DEVNULL,
            creationflags=0x08000000,
        )
        return [p.strip() for p in out.strip().splitlines() if p.strip()]
    except Exception:
        return []


def _get_default_printer() -> str:
    """Return the system default printer name via winspool, falling back to PowerShell."""
    buf  = ctypes.create_unicode_buffer(512)
    size = ctypes.c_ulong(512)
    if _winspool.GetDefaultPrinterW(buf, ctypes.byref(size)):
        return buf.value
    # Fallback: PowerShell
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command",
             "(Get-CimInstance Win32_Printer | Where-Object Default).Name"],
            text=True, timeout=10, stderr=subprocess.DEVNULL,
            creationflags=0x08000000,
        )
        return out.strip()
    except Exception:
        return ""


class _PrintDialog(tk.Toplevel):
    """Modal print dialog: choose printer, page range, and number of copies."""

    def __init__(self, parent: tk.Tk, n_pages: int, current_page: int) -> None:
        super().__init__(parent)
        self.title("Print")
        self.resizable(False, False)
        self.grab_set()
        self.result = None          # (printer_name, page_list, copies) or None
        self._n_pages = n_pages
        self._current_page = current_page

        printers = _get_printers()
        default  = _get_default_printer()

        # ── Printer ──────────────────────────────────────────────────────
        pf = ttk.LabelFrame(self, text="Printer", padding=8)
        pf.pack(fill=tk.X, padx=12, pady=6)

        ttk.Label(pf, text="Name:").grid(row=0, column=0, sticky=tk.W)
        self._printer_var = tk.StringVar(value=default)
        self._printer_cb  = ttk.Combobox(
            pf, textvariable=self._printer_var,
            values=printers, width=44, state="readonly")
        self._printer_cb.grid(row=0, column=1, padx=6, sticky=tk.W)
        if not printers:
            self._printer_cb.config(state="normal")
            self._printer_var.set("(no printers found)")
        ttk.Button(pf, text="Properties\u2026", command=self._open_printer_props
                   ).grid(row=0, column=2, padx=6)

        # ── Page range ───────────────────────────────────────────────────
        rf = ttk.LabelFrame(self, text="Print range", padding=8)
        rf.pack(fill=tk.X, padx=12, pady=6)

        self._range_var = tk.StringVar(value="all")
        ttk.Radiobutton(rf, text="All pages",
                        variable=self._range_var, value="all"
                        ).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(rf, text=f"Current page  ({current_page + 1})",
                        variable=self._range_var, value="current"
                        ).grid(row=1, column=0, columnspan=3, sticky=tk.W)
        ttk.Radiobutton(rf, text="Pages:",
                        variable=self._range_var, value="range"
                        ).grid(row=2, column=0, sticky=tk.W)
        self._range_entry = ttk.Entry(rf, width=16)
        self._range_entry.insert(0, f"1-{n_pages}")
        self._range_entry.grid(row=2, column=1, padx=4, sticky=tk.W)
        ttk.Label(rf, text=f"e.g. 1-3, 5   (1 \u2013 {n_pages})",
                  foreground="#666").grid(row=2, column=2, padx=4, sticky=tk.W)
        # Switch to 'range' mode automatically when the entry is clicked
        self._range_entry.bind("<FocusIn>", lambda _e: self._range_var.set("range"))

        # ── Copies ───────────────────────────────────────────────────────
        cf = ttk.LabelFrame(self, text="Copies", padding=8)
        cf.pack(fill=tk.X, padx=12, pady=6)

        ttk.Label(cf, text="Number of copies:").grid(row=0, column=0, sticky=tk.W)
        self._copies_var = tk.StringVar(value="1")
        tk.Spinbox(cf, from_=1, to=99, textvariable=self._copies_var,
                   width=4).grid(row=0, column=1, padx=6)

        # ── Buttons ──────────────────────────────────────────────────────
        bf = tk.Frame(self)
        bf.pack(pady=10)
        ttk.Button(bf, text="Print",  command=self._on_print,  width=10).pack(side=tk.LEFT, padx=6)
        ttk.Button(bf, text="Cancel", command=self.destroy,    width=8 ).pack(side=tk.LEFT, padx=6)

        self.transient(parent)
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")
        self.bind("<Return>", lambda _e: self._on_print())
        self.bind("<Escape>", lambda _e: self.destroy())

    def _on_print(self) -> None:
        """Validate inputs, store result, and close."""
        printer = self._printer_var.get().strip()
        if not printer or printer.startswith("("):
            messagebox.showwarning("No printer", "Please select a printer.", parent=self)
            return
        try:
            copies = max(1, int(self._copies_var.get()))
        except ValueError:
            copies = 1
        mode = self._range_var.get()
        if mode == "all":
            pages = list(range(self._n_pages))
        elif mode == "current":
            pages = [self._current_page]
        else:
            pages = self._parse_range(self._range_entry.get())
            if pages is None:
                messagebox.showwarning(
                    "Invalid range",
                    f"Enter a valid page range (e.g. 1-3, 5).\n"
                    f"Valid page numbers: 1 \u2013 {self._n_pages}.",
                    parent=self)
                return
        self.result = (printer, pages, copies)
        self.destroy()

    def _open_printer_props(self) -> None:
        """Open the Windows printer properties dialog for the selected printer."""
        printer = self._printer_var.get().strip()
        if not printer or printer.startswith("("):
            messagebox.showwarning("No printer", "Please select a printer first.", parent=self)
            return
        try:
            subprocess.Popen(["rundll32.exe", "printui.dll,PrintUIEntry", "/p", "/n", printer])
        except Exception as exc:
            messagebox.showerror("Error", str(exc), parent=self)

    def _parse_range(self, text: str):
        """Parse '1-3,5' into a zero-based page list, or return None on error."""
        pages = []
        try:
            for part in text.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    a, b = part.split("-", 1)
                    a, b = int(a.strip()) - 1, int(b.strip()) - 1
                    if a < 0 or b >= self._n_pages or a > b:
                        return None
                    pages.extend(range(a, b + 1))
                else:
                    p = int(part) - 1
                    if p < 0 or p >= self._n_pages:
                        return None
                    pages.append(p)
            return pages or None
        except ValueError:
            return None


# ---------------------------------------------------------------------------
# About dialog
# ---------------------------------------------------------------------------

class _AboutDialog(tk.Toplevel):
    """Modal 'About' dialog showing version, build date, repo, and author info."""

    _GITHUB_URL  = "https://github.com/FlorianBock/helper_pdf-editor"
    _AUTHOR_NAME  = "Florian Bock"
    _AUTHOR_EMAIL = "florian.bock.mobile@googlemail.com"

    def __init__(self, parent: tk.Tk) -> None:
        super().__init__(parent)
        self.title("About Helper: PDF Editor")
        self.resizable(False, False)
        self.grab_set()

        pad = dict(padx=16, pady=6)

        # ── App name ──────────────────────────────────────────────────────
        tk.Label(self, text="Helper: PDF Editor",
                 font=("Arial", 14, "bold")).pack(padx=16, pady=(16, 2))

        # ── Version / build date ──────────────────────────────────────────
        if BUILD_DATE:
            ver_text = f"Version {VERSION}   ·   Built {BUILD_DATE}"
        else:
            ver_text = f"Version {VERSION}   ·   (running from source)"
        tk.Label(self, text=ver_text, fg="#444").pack(padx=16, pady=(0, 10))

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=16)

        # ── GitHub repo ───────────────────────────────────────────────────
        repo_frame = tk.Frame(self)
        repo_frame.pack(**pad)
        tk.Label(repo_frame, text="GitHub:", width=9,
                 anchor=tk.E, fg="#444").pack(side=tk.LEFT)
        repo_link = tk.Label(repo_frame, text=self._GITHUB_URL,
                             fg="#0066cc", cursor="hand2", font=("Arial", 9, "underline"))
        repo_link.pack(side=tk.LEFT, padx=(4, 0))
        repo_link.bind("<Button-1>",
                       lambda _e: webbrowser.open(self._GITHUB_URL))

        # ── Author ────────────────────────────────────────────────────────
        auth_frame = tk.Frame(self)
        auth_frame.pack(padx=16, pady=(2, 6))
        tk.Label(auth_frame, text="Author:", width=9,
                 anchor=tk.E, fg="#444").pack(side=tk.LEFT)
        tk.Label(auth_frame, text=f"{self._AUTHOR_NAME}  ",
                 fg="#222").pack(side=tk.LEFT)
        mail_link = tk.Label(auth_frame, text=self._AUTHOR_EMAIL,
                             fg="#0066cc", cursor="hand2", font=("Arial", 9, "underline"))
        mail_link.pack(side=tk.LEFT)
        mail_link.bind("<Button-1>",
                       lambda _e: webbrowser.open(f"mailto:{self._AUTHOR_EMAIL}"))

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=16, pady=(6, 0))

        # ── Close button ──────────────────────────────────────────────────
        ttk.Button(self, text="Close", command=self.destroy,
                   width=10).pack(pady=12)

        self.transient(parent)
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")
        self.bind("<Escape>", lambda _e: self.destroy())
        self.bind("<Return>", lambda _e: self.destroy())


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class PDFFormFiller(_AppBase):

    _ZOOM_OPTIONS = ["50%", "75%", "100%", "125%", "150%", "175%", "200%", "250%", "300%"]
    _DEFAULT_ZOOM = "150%"
    _PAD = 16           # canvas padding around the page image (px)
    _DEFAULT_FS = 11.0  # default font size (pt)

    def __init__(self) -> None:
        super().__init__()
        self.title("Helper: PDF Editor")
        self.geometry("1100x820")
        self.minsize(700, 500)

        # Document state
        self._doc: fitz.Document | None = None
        self._pdf_path: str = ""
        self._current_page: int = 0
        self._zoom: float = 1.5

        # Free-text placements (persist across page navigation)
        self._placements: list[TextPlacement] = []

        # AcroForm bookkeeping for current page
        self._acro_widgets: dict[str, tuple] = {}

        # Drag / resize state
        self._drag_data: dict = {}
        self._resize_data: dict = {}

        # Page-op button refs (created in _build_page_toolbar, needed by _refresh_controls)
        self._btn_rot_ccw = self._btn_rot_cw = None
        self._btn_mir_h = self._btn_mir_v = None
        self._btn_pg_up = self._btn_pg_dn = None
        self._btn_del_pg = self._btn_ins_pg = self._btn_dup_pg = None
        self._grip_w_cache: int = 0        # lazily measured grip pixel width
        self._right_ctrl_w_cache: int = 0   # lazily measured right-side controls width

        # Undo stack — each entry is either a TextPlacement, EraserRect, ShapeRect, or RegionClip
        self._undo_stack: list[TextPlacement | EraserRect | ShapeRect | RegionClip] = []

        # Eraser rectangles (white boxes painted over content)
        self._erasers: list[EraserRect] = []

        # Shapes drawn on pages
        self._shapes: list[ShapeRect] = []

        # Pasted region clips
        self._clips: list[RegionClip] = []

        # Clipboard: the most-recently copied region (PIL Image + PDF-unit size)
        self._clipboard_image: object = None   # PIL.Image or None
        self._clipboard_size_pdf: tuple[float, float] = (0.0, 0.0)  # (w_pdf, h_pdf)

        # Continuous scroll state (canvas item ids for page images, keyed by page index)
        self._page_offsets: dict[int, tuple[int, int]] = {}  # page_idx -> (cx, cy)

        # Eraser sub-mode
        self._eraser_kind_var = tk.StringVar(value="rect")
        self._eraser_pencil_size_var = tk.IntVar(value=10)

        # Eraser drag state
        self._eraser_start: tuple[float, float] | None = None
        self._eraser_rect_id: int = 0
        self._eraser_lasso_points: list[tuple[float, float]] = []   # lasso/free polyline

        # Shape drag state
        self._shape_start: tuple[float, float] | None = None
        self._shape_preview_id: int = 0

        # Copy-region drag state
        self._copy_start: tuple[float, float] | None = None
        self._copy_preview_id: int = 0

        # Clip drag state (moving a pasted clipboard image)
        self._clip_drag_data: dict = {}

        # Placement mode (set before _build_ui so toolbar can bind to it)
        self._mode_var = tk.StringVar(value="text")

        # Shape style vars (populated in _build_shape_toolbar)
        self._shape_kind_var    = tk.StringVar(value="rect")
        self._shape_stroke_var  = tk.StringVar(value="#000000")
        self._shape_fill_var    = tk.StringVar(value="")        # "" = no fill
        self._shape_lw_var      = tk.StringVar(value="1.5")
        self._shape_filled_var  = tk.BooleanVar(value=False)    # fill checkbox

        # Continuous scroll mode
        self._continuous_var = tk.BooleanVar(value=True)

        # Prevent GC of photo image
        self._photo_image: ImageTk.PhotoImage | None = None
        # List of photo images used in continuous-scroll mode (prevents GC)
        self._photo_images: list[ImageTk.PhotoImage] = []

        self._build_ui()
        self._mode_var.trace_add("write", self._on_mode_change)
        self._setup_drop_target()
        self.bind_all("<Control-z>", lambda _e: self._undo())
        self.bind_all("<Control-Z>", lambda _e: self._undo())
        self._refresh_controls()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        """Assemble all UI regions: toolbars, canvas area, and status bar."""
        self._build_toolbar()          # row 1 – file, navigation, zoom
        self._build_page_toolbar()     # row 2 – page-level operations
        self._build_editing_toolbar()  # row 3 – text/mark editing controls
        self._build_eraser_toolbar()   # row 4 – eraser sub-mode controls (hidden by default)
        self._build_shape_toolbar()    # row 5 – shape style controls (hidden by default)
        self._build_canvas_area()      # main scrollable PDF canvas
        self._build_statusbar()        # bottom status line

    def _build_toolbar(self) -> None:
        # Wrap toolbar in a horizontally-scrollable container so buttons
        # are never cropped when the window is narrow.
        outer = tk.Frame(self, bd=1, relief=tk.RAISED)
        outer.pack(side=tk.TOP, fill=tk.X)

        tb_canvas = tk.Canvas(outer, height=36, highlightthickness=0, bd=0)
        h_scroll = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=tb_canvas.xview)
        tb_canvas.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        tb_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)

        bar = tk.Frame(tb_canvas, padx=4, pady=3)
        bar_win = tb_canvas.create_window(0, 0, anchor=tk.NW, window=bar)

        def _on_bar_configure(event):
            tb_canvas.configure(scrollregion=tb_canvas.bbox("all"))
            # Make the canvas tall enough for the inner frame
            tb_canvas.configure(height=bar.winfo_reqheight())
        bar.bind("<Configure>", _on_bar_configure)

        # --- File ---
        tk.Button(bar, text="Open PDF…", command=self._open_pdf, width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text="Save PDF…", command=self._save_pdf, width=10).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text="Print…", command=self._print_pdf, width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(bar, text="↩ Undo", command=self._undo, width=7).pack(side=tk.LEFT, padx=2)
        _sep(bar)
        tk.Button(bar, text="About", command=self._show_about, width=6).pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Navigation ---
        self._btn_prev = tk.Button(bar, text="◀ Prev", command=self._prev_page, width=8)
        self._btn_prev.pack(side=tk.LEFT, padx=2)
        self._lbl_page = tk.Label(bar, text="—", width=14)
        self._lbl_page.pack(side=tk.LEFT)
        self._btn_next = tk.Button(bar, text="Next ▶", command=self._next_page, width=8)
        self._btn_next.pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Zoom ---
        tk.Label(bar, text="Zoom:").pack(side=tk.LEFT)
        self._zoom_var = tk.StringVar(value=self._DEFAULT_ZOOM)
        combo = ttk.Combobox(bar, textvariable=self._zoom_var, values=self._ZOOM_OPTIONS,
                             width=6, state="readonly")
        combo.pack(side=tk.LEFT, padx=4)
        combo.bind("<<ComboboxSelected>>", self._on_zoom_change)

    def _build_page_toolbar(self) -> None:
        """Build the page operations toolbar (row 2): rotate, mirror, reorder, insert/delete."""
        bar2 = tk.Frame(self, bd=1, relief=tk.RAISED, padx=4, pady=2)
        bar2.pack(side=tk.TOP, fill=tk.X)

        tk.Label(bar2, text="Page:", font=("Arial", 8, "bold")).pack(side=tk.LEFT, padx=(4, 2))

        self._btn_rot_ccw = tk.Button(bar2, text="↺ Rotate 90° CCW",
                                      command=lambda: self._rotate_page(-90), width=16)
        self._btn_rot_ccw.pack(side=tk.LEFT, padx=2)

        self._btn_rot_cw = tk.Button(bar2, text="↻ Rotate 90° CW",
                                     command=lambda: self._rotate_page(90), width=15)
        self._btn_rot_cw.pack(side=tk.LEFT, padx=2)
        _sep(bar2)

        self._btn_mir_h = tk.Button(bar2, text="⇔ Mirror H",
                                    command=lambda: self._mirror_page(horizontal=True), width=10)
        self._btn_mir_h.pack(side=tk.LEFT, padx=2)

        self._btn_mir_v = tk.Button(bar2, text="⇕ Mirror V",
                                    command=lambda: self._mirror_page(horizontal=False), width=10)
        self._btn_mir_v.pack(side=tk.LEFT, padx=2)
        _sep(bar2)

        self._btn_pg_up = tk.Button(bar2, text="▲ Move page up",
                                    command=self._move_page_up, width=14)
        self._btn_pg_up.pack(side=tk.LEFT, padx=2)

        self._btn_pg_dn = tk.Button(bar2, text="▼ Move page down",
                                    command=self._move_page_down, width=15)
        self._btn_pg_dn.pack(side=tk.LEFT, padx=2)
        _sep(bar2)

        self._btn_del_pg = tk.Button(bar2, text="🗑 Delete page",
                                     command=self._delete_page, fg="#aa0000", width=12)
        self._btn_del_pg.pack(side=tk.LEFT, padx=2)

        self._btn_ins_pg = tk.Button(bar2, text="+ Insert blank page",
                                     command=self._insert_blank_page, width=17)
        self._btn_ins_pg.pack(side=tk.LEFT, padx=2)
        _sep(bar2)

        self._btn_dup_pg = tk.Button(bar2, text="⎘ Duplicate page",
                                     command=self._duplicate_page, width=16)
        self._btn_dup_pg.pack(side=tk.LEFT, padx=2)

    def _build_editing_toolbar(self) -> None:
        """Build the editing toolbar (row 3): font, placement mode, eraser, scroll toggle.

        Placed on its own scrollable row so these controls are always fully visible
        without scrolling the file / navigation bar.
        """
        outer = tk.Frame(self, bd=1, relief=tk.RAISED)
        outer.pack(side=tk.TOP, fill=tk.X)

        tb_canvas = tk.Canvas(outer, height=36, highlightthickness=0, bd=0)
        h_scroll = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=tb_canvas.xview)
        tb_canvas.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        tb_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)

        bar = tk.Frame(tb_canvas, padx=4, pady=3)
        tb_canvas.create_window(0, 0, anchor=tk.NW, window=bar)

        def _on_bar_configure(event):
            # Keep scroll region and canvas height in sync with the inner frame.
            tb_canvas.configure(scrollregion=tb_canvas.bbox("all"),
                                height=bar.winfo_reqheight())
        bar.bind("<Configure>", _on_bar_configure)

        # --- Font size spinner ---
        tk.Label(bar, text="Font size:").pack(side=tk.LEFT)
        self._fs_var = tk.StringVar(value=str(self._DEFAULT_FS))
        spin = tk.Spinbox(bar, from_=6, to=72, increment=1, textvariable=self._fs_var,
                          width=4)
        spin.pack(side=tk.LEFT, padx=2)

        # --- Bold toggle ---
        self._bold_var = tk.BooleanVar(value=False)
        tk.Checkbutton(bar, text="Bold", variable=self._bold_var).pack(side=tk.LEFT, padx=4)
        _sep(bar)

        # --- Clear all text/marks on the current page ---
        tk.Button(bar, text="Clear page texts", command=self._clear_page,
                  fg="#aa0000").pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Placement mode: free text, tick mark, cross mark, or eraser rectangle ---
        tk.Label(bar, text="Mode:").pack(side=tk.LEFT)
        tk.Radiobutton(bar, text="Text", variable=self._mode_var, value="text",
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(bar, text="✓", variable=self._mode_var, value="check_v",
                       fg="#006600", font=("Arial", 12, "bold"),
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(bar, text="✗", variable=self._mode_var, value="check_x",
                       fg="#990000", font=("Arial", 12, "bold"),
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Eraser mode (drag a rectangle to white-out content on save) ---
        tk.Radiobutton(bar, text="⬜ Eraser", variable=self._mode_var, value="eraser",
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Shape draw mode ---
        tk.Radiobutton(bar, text="⬛ Shape", variable=self._mode_var, value="shape",
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Copy region / Paste modes ---
        tk.Radiobutton(bar, text="⧉ Copy region", variable=self._mode_var, value="copy_region",
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(bar, text="📋 Paste", variable=self._mode_var, value="paste",
                       indicatoron=True).pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # --- Continuous-scroll toggle ---
        tk.Checkbutton(bar, text="Continuous scroll", variable=self._continuous_var,
                       command=self._on_continuous_toggle).pack(side=tk.LEFT, padx=4)
        _sep(bar)

        # --- Quick-reference hint label ---
        tk.Label(bar,
                 text="Click to add  •  Drag ⠿ to move  •  × to delete  •  Drag to erase/draw",
                 fg="#555", font=("Arial", 8, "italic")).pack(side=tk.LEFT, padx=6)

    def _build_eraser_toolbar(self) -> None:
        """Build the eraser sub-mode toolbar (hidden until eraser mode is active)."""
        outer = tk.Frame(self, bd=1, relief=tk.RAISED)
        # Do NOT pack yet — shown/hidden by _on_mode_change
        self._eraser_toolbar_outer = outer

        tb_canvas = tk.Canvas(outer, height=36, highlightthickness=0, bd=0)
        h_scroll = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=tb_canvas.xview)
        tb_canvas.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        tb_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)

        bar = tk.Frame(tb_canvas, padx=4, pady=3)
        tb_canvas.create_window(0, 0, anchor=tk.NW, window=bar)

        def _on_bar_configure(event):
            tb_canvas.configure(scrollregion=tb_canvas.bbox("all"),
                                height=bar.winfo_reqheight())
        bar.bind("<Configure>", _on_bar_configure)

        tk.Label(bar, text="Eraser:", font=("Arial", 8, "bold")).pack(side=tk.LEFT, padx=(2, 4))

        for label, val in (("⬜ Rectangle", "rect"), ("⭕ Circle", "circle"),
                           ("⚓ Lasso", "lasso"), ("✏ Free", "free")):
            tk.Radiobutton(bar, text=label, variable=self._eraser_kind_var, value=val,
                           indicatoron=True).pack(side=tk.LEFT, padx=2)

        _sep(bar)

        tk.Label(bar, text="Pencil size:").pack(side=tk.LEFT)
        self._pencil_size_spin = tk.Spinbox(
            bar, from_=1, to=200, increment=1,
            textvariable=self._eraser_pencil_size_var, width=4)
        self._pencil_size_spin.pack(side=tk.LEFT, padx=2)
        tk.Label(bar, text="px", fg="#555").pack(side=tk.LEFT)

        def _update_pencil_spin(*_):
            state = tk.NORMAL if self._eraser_kind_var.get() == "free" else tk.DISABLED
            self._pencil_size_spin.config(state=state)
        self._eraser_kind_var.trace_add("write", _update_pencil_spin)
        _update_pencil_spin()  # apply initial state

    def _build_shape_toolbar(self) -> None:
        """Build the shape style toolbar (hidden until shape mode is active): kind, stroke colour, fill, line width."""
        outer = tk.Frame(self, bd=1, relief=tk.RAISED)
        # Do NOT pack yet — shown/hidden by _on_mode_change
        self._shape_toolbar_outer = outer

        tb_canvas = tk.Canvas(outer, height=36, highlightthickness=0, bd=0)
        h_scroll = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=tb_canvas.xview)
        tb_canvas.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        tb_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)

        bar = tk.Frame(tb_canvas, padx=4, pady=3)
        tb_canvas.create_window(0, 0, anchor=tk.NW, window=bar)

        def _on_bar_configure(event):
            tb_canvas.configure(scrollregion=tb_canvas.bbox("all"),
                                height=bar.winfo_reqheight())
        bar.bind("<Configure>", _on_bar_configure)

        tk.Label(bar, text="Shape:", font=("Arial", 8, "bold")).pack(side=tk.LEFT, padx=(2, 4))

        # Shape kind
        for label, val in (("Rectangle", "rect"), ("Rounded rect", "rounded_rect"),
                           ("Ellipse", "ellipse"), ("Line", "line"), ("Arrow", "arrow")):
            tk.Radiobutton(bar, text=label, variable=self._shape_kind_var, value=val,
                           indicatoron=True).pack(side=tk.LEFT, padx=2)
        _sep(bar)

        # Stroke colour swatch
        tk.Label(bar, text="Stroke:").pack(side=tk.LEFT)
        self._stroke_swatch = tk.Label(bar, bg=self._shape_stroke_var.get(),
                                       width=3, relief=tk.SUNKEN, cursor="hand2")
        self._stroke_swatch.pack(side=tk.LEFT, padx=2)
        self._stroke_swatch.bind("<Button-1>", lambda _e: self._pick_color("stroke"))

        _sep(bar)

        # Fill toggle + colour swatch
        tk.Checkbutton(bar, text="Fill", variable=self._shape_filled_var,
                       command=self._on_fill_toggle).pack(side=tk.LEFT)
        self._fill_swatch = tk.Label(bar, bg="#ffffff" if not self._shape_filled_var.get() else self._shape_fill_var.get() or "#ffffff",
                                     width=3, relief=tk.SUNKEN, cursor="hand2")
        self._fill_swatch.pack(side=tk.LEFT, padx=2)
        self._fill_swatch.bind("<Button-1>", lambda _e: self._pick_color("fill"))

        _sep(bar)

        # Line width
        tk.Label(bar, text="Width:").pack(side=tk.LEFT)
        tk.Spinbox(bar, from_=0.5, to=20.0, increment=0.5,
                   textvariable=self._shape_lw_var, width=4,
                   format="%.1f").pack(side=tk.LEFT, padx=2)

    def _pick_color(self, target: str) -> None:
        """Open the Tk colour chooser and apply the result to stroke or fill."""
        from tkinter import colorchooser
        initial = self._shape_stroke_var.get() if target == "stroke" else (self._shape_fill_var.get() or "#ffffff")
        _, hexcol = colorchooser.askcolor(color=initial, title=f"Choose {target} colour", parent=self)
        if hexcol:
            if target == "stroke":
                self._shape_stroke_var.set(hexcol)
                self._stroke_swatch.config(bg=hexcol)
            else:
                self._shape_fill_var.set(hexcol)
                self._fill_swatch.config(bg=hexcol)

    def _on_fill_toggle(self) -> None:
        """Grey-out fill swatch when fill is disabled."""
        if self._shape_filled_var.get():
            col = self._shape_fill_var.get() or "#ffffff"
            self._fill_swatch.config(bg=col)
        else:
            self._fill_swatch.config(bg="#d9d9d9")

    def _on_mode_change(self, *_) -> None:
        """Show/hide the eraser and shape sub-toolbars based on the active mode."""
        mode = self._mode_var.get()
        if mode == "eraser":
            self._eraser_toolbar_outer.pack(side=tk.TOP, fill=tk.X,
                                            before=self._canvas_area_outer)
            self._shape_toolbar_outer.pack_forget()
        elif mode == "shape":
            self._shape_toolbar_outer.pack(side=tk.TOP, fill=tk.X,
                                           before=self._canvas_area_outer)
            self._eraser_toolbar_outer.pack_forget()
        else:
            self._eraser_toolbar_outer.pack_forget()
            self._shape_toolbar_outer.pack_forget()

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> tuple[float, float, float]:
        """Convert '#rrggbb' to a (r, g, b) tuple with values 0.0–1.0."""
        h = hex_color.lstrip("#")
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        r = int(h[0:2], 16) / 255.0
        g = int(h[2:4], 16) / 255.0
        b = int(h[4:6], 16) / 255.0
        return (r, g, b)

    def _build_canvas_area(self) -> None:
        """Build the main PDF canvas with horizontal and vertical scrollbars."""
        outer = tk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True)
        self._canvas_area_outer = outer

        self._canvas = tk.Canvas(outer, bg="#606060", cursor="crosshair", highlightthickness=0)
        vs = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=self._canvas.yview)
        hs = ttk.Scrollbar(outer, orient=tk.HORIZONTAL, command=self._canvas.xview)
        self._canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        vs.pack(side=tk.RIGHT, fill=tk.Y)
        hs.pack(side=tk.BOTTOM, fill=tk.X)
        self._canvas.pack(fill=tk.BOTH, expand=True)

        self._canvas.bind("<Button-1>", self._on_canvas_click)
        self._canvas.bind("<B1-Motion>", self._on_canvas_drag)
        self._canvas.bind("<ButtonRelease-1>", self._on_canvas_release)
        self._canvas.bind("<MouseWheel>", self._on_mousewheel)
        self._canvas.bind("<Button-4>", lambda _: self._canvas.yview_scroll(-1, "units"))
        self._canvas.bind("<Button-5>", lambda _: self._canvas.yview_scroll(1, "units"))

    def _build_statusbar(self) -> None:
        """Build the single-line status bar shown at the very bottom of the window."""
        self._status_var = tk.StringVar(value="Open a PDF file to begin.")
        tk.Label(
            self,
            textvariable=self._status_var,
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=6,
        ).pack(side=tk.BOTTOM, fill=tk.X)

    # ------------------------------------------------------------------
    # File operations
    # ------------------------------------------------------------------

    def _open_pdf(self) -> None:
        """Prompt the user to choose a PDF file and load it for editing."""
        path = filedialog.askopenfilename(
            title="Open PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self._open_dropped_pdf(path)

    def _has_unsaved_changes(self) -> bool:
        """Return True if there are any annotations/edits not yet saved to disk."""
        return bool(self._doc and (
            self._placements or self._erasers or self._shapes or self._clips))

    def _open_dropped_pdf(self, path: str) -> None:
        """Open *path*, asking about unsaved changes first if a document is already loaded."""
        if self._has_unsaved_changes():
            answer = messagebox.askyesnocancel(
                "Unsaved Changes",
                f"The current document has unsaved changes.\n\n"
                f"Save before opening '{os.path.basename(path)}'?",
                parent=self,
            )
            if answer is None:   # Cancel — abort
                return
            if answer:           # Yes — save first
                self._save_pdf()
        self._load_pdf(path)

    def _setup_drop_target(self) -> None:
        """Register the window and all child widgets to accept drag-and-drop PDF files via tkinterdnd2."""
        if _DND_FILES is None:
            return  # tkinterdnd2 not installed — drag-and-drop unavailable
        # Defer until the first idle cycle so every widget has a real HWND
        # (Tk only assigns Win32 window handles once it renders the widget tree;
        # calling drop_target_register before that silently does nothing).
        self.after_idle(self._do_register_drop_targets)

    def _do_register_drop_targets(self) -> None:
        """Actually register all widgets; called one idle cycle after startup."""
        self.update_idletasks()   # ensure all widgets are realised
        self._register_drop_recursive(self)

    def _register_drop_recursive(self, widget) -> None:
        """Register *widget* and all its descendants as drop targets for PDF files."""
        if _DND_FILES is None:
            return
        try:
            widget.drop_target_register(_DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_file_drop)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._register_drop_recursive(child)

    def _on_file_drop(self, event) -> None:
        """Handle a file-drop event from tkinterdnd2."""
        raw = event.data.strip()
        # tkinterdnd2 wraps paths that contain spaces in curly braces, and
        # may deliver multiple files as '{path1} {path2}' — pick the first PDF.
        import re as _re
        if raw.startswith("{"):
            candidates = _re.findall(r"\{([^}]+)\}", raw)
            if not candidates:
                candidates = [raw[1:-1]]
        else:
            candidates = raw.split()
        for p in candidates:
            if p.lower().endswith(".pdf") and os.path.isfile(p):
                self.after(0, lambda path=p: self._open_dropped_pdf(path))
                break

    def _load_pdf(self, path: str) -> None:
        """Open *path* and replace the current document with it.

        Called by _open_pdf (via dialog) and from main() when a file path is
        supplied as a command-line argument (e.g. drag-and-drop onto the exe).
        """
        try:
            doc = fitz.open(path)
        except Exception as exc:
            messagebox.showerror("Error opening PDF", str(exc))
            return

        if self._doc:
            self._doc.close()
        self._doc = doc
        self._pdf_path = path
        self._current_page = 0
        self._placements.clear()
        self._undo_stack.clear()
        self._acro_widgets.clear()
        self._erasers.clear()
        self._shapes.clear()
        self._clips.clear()
        self._clipboard_image = None
        self._page_offsets.clear()
        self._photo_images.clear()
        self.title(f"Helper: PDF Editor — {os.path.basename(path)}")
        self._render_page()
        self._refresh_controls()
        n_fields = sum(len(list(doc[i].widgets())) for i in range(len(doc)))
        if n_fields:
            self._status_var.set(f"Opened: {path}  •  {n_fields} form field(s) found")
        else:
            self._status_var.set(
                f"Opened: {path}  •  No form fields — click anywhere to add text")

    def _print_pdf(self) -> None:
        """Open the print dialog; on confirmation, render pages to the selected printer."""
        if not self._doc:
            messagebox.showwarning("No document", "Please open a PDF file first.")
            return
        self._flush_page()
        dlg = _PrintDialog(self, len(self._doc), self._current_page)
        self.wait_window(dlg)
        if dlg.result is None:
            return
        printer_name, pages, copies = dlg.result
        self._do_print(printer_name, pages, copies)

    def _do_print(self, printer_name: str, pages: list, copies: int) -> None:
        """Render *pages* of the current document and spool them to *printer_name*.

        Uses GDI32 (CreateDCW / StretchDIBits) via ctypes so no third-party
        packages are required beyond PyMuPDF and Pillow.
        """
        # Build an edited in-memory copy of the document.
        buf = io.BytesIO()
        self._doc.save(buf)
        buf.seek(0)
        out_doc = fitz.open("pdf", buf)
        self._apply_placements_to(out_doc)

        hdc = None
        doc_started = False
        try:
            hdc = _gdi32.CreateDCW("WINSPOOL", printer_name, None, None)
            if not hdc:
                raise RuntimeError(f'Could not open printer "{printer_name}".\n'
                                   'Check that the printer name is correct and the '
                                   'printer is installed.')

            pr_w  = _gdi32.GetDeviceCaps(hdc, _HORZRES)
            pr_h  = _gdi32.GetDeviceCaps(hdc, _VERTRES)
            dpi_x = _gdi32.GetDeviceCaps(hdc, _LOGPIXELSX) or 300
            dpi_y = _gdi32.GetDeviceCaps(hdc, _LOGPIXELSY) or 300

            doc_name = os.path.basename(self._pdf_path) if self._pdf_path else "PDF"
            di = _DOCINFOW()
            di.cbSize      = ctypes.sizeof(_DOCINFOW)
            di.lpszDocName = doc_name
            di.lpszOutput  = None
            di.lpszDatatype = None
            di.fwType      = 0
            if _gdi32.StartDocW(hdc, ctypes.byref(di)) <= 0:
                raise RuntimeError("StartDoc failed.")
            doc_started = True

            for _ in range(copies):
                for pg_idx in pages:
                    page = out_doc[pg_idx]
                    # Render at printer DPI, capped at 300 dpi to keep memory sane.
                    rx  = min(dpi_x, 300)
                    ry  = min(dpi_y, 300)
                    mat = fitz.Matrix(rx / 72.0, ry / 72.0)
                    pix = page.get_pixmap(matrix=mat, alpha=False, colorspace=fitz.csRGB)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

                    # Scale to fill the printable area, maintaining aspect ratio.
                    scale = min(pr_w / pix.width, pr_h / pix.height)
                    dst_w = int(pix.width  * scale)
                    dst_h = int(pix.height * scale)
                    left  = (pr_w - dst_w) // 2
                    top   = (pr_h - dst_h) // 2

                    if _gdi32.StartPage(hdc) <= 0:
                        raise RuntimeError("StartPage failed.")

                    # Build BITMAPINFO for StretchDIBits.
                    # Windows DIB scanlines must be DWORD-aligned and in BGR order.
                    stride = (pix.width * 3 + 3) & ~3  # bytes per (padded) row
                    raw = img.tobytes("raw", "BGR", stride, 1)

                    bmi = _BITMAPINFO()
                    bmi.bmiHeader.biSize        = ctypes.sizeof(_BITMAPINFOHEADER)
                    bmi.bmiHeader.biWidth       = pix.width
                    bmi.bmiHeader.biHeight      = -pix.height  # negative = top-down
                    bmi.bmiHeader.biPlanes      = 1
                    bmi.bmiHeader.biBitCount    = 24
                    bmi.bmiHeader.biCompression = _BI_RGB
                    bmi.bmiHeader.biSizeImage   = len(raw)

                    _gdi32.StretchDIBits(
                        hdc,
                        left, top, dst_w, dst_h,
                        0, 0, pix.width, pix.height,
                        raw, ctypes.byref(bmi),
                        _DIB_RGB_COLORS, _SRCCOPY,
                    )
                    _gdi32.EndPage(hdc)

            _gdi32.EndDoc(hdc)
            doc_started = False
        except Exception as exc:
            if doc_started and hdc:
                try:
                    _gdi32.AbortDoc(hdc)
                except Exception:
                    pass
            messagebox.showerror("Print error", str(exc))
        finally:
            if hdc:
                _gdi32.DeleteDC(hdc)
            out_doc.close()

    def _save_pdf(self) -> None:
        """Burn all placements/erasers into a fresh copy of the PDF and save to disk."""
        if not self._doc:
            messagebox.showwarning("No document", "Please open a PDF file first.")
            return

        self._flush_page()   # capture latest widget values

        base = os.path.splitext(os.path.basename(self._pdf_path))[0] if self._pdf_path else "output"
        path = filedialog.asksaveasfilename(
            title="Save Filled PDF",
            initialfile=base + "_filled.pdf",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return

        # Always work on a fresh in-memory copy so the live document is never
        # mutated — repeated saves will not stack up duplicate marks.
        buf = io.BytesIO()
        self._doc.save(buf)
        buf.seek(0)
        out_doc = fitz.open("pdf", buf)
        self._apply_placements_to(out_doc)
        try:
            out_doc.save(path, garbage=4, deflate=True)
        except Exception as exc:
            out_doc.close()
            messagebox.showerror("Error saving PDF", str(exc))
            return
        out_doc.close()

        self._status_var.set(f"Saved: {path}")
        messagebox.showinfo("Saved", f"PDF saved to:\n{path}")

    def _apply_placements_to(self, doc: fitz.Document) -> None:
        """Burn all free-text/check placements, erasers, shapes, and clips into *doc* (a copy)."""
        # First burn white eraser rectangles so they sit beneath new content
        for er in self._erasers:
            page: fitz.Page = doc[er.page_idx]
            shape = page.new_shape()
            if er.kind == "circle":
                shape.draw_oval(fitz.Rect(er.x0, er.y0, er.x1, er.y1))
                shape.finish(color=(1, 1, 1), fill=(1, 1, 1), width=0)
            elif er.kind == "lasso" and len(er.points) >= 3:
                shape.draw_polyline([fitz.Point(p[0], p[1]) for p in er.points])
                shape.finish(color=(1, 1, 1), fill=(1, 1, 1), width=0, closePath=True)
            elif er.kind == "free" and len(er.points) >= 2:
                shape.draw_squiggle([fitz.Point(p[0], p[1]) for p in er.points], curviness=0)
                shape.finish(color=(1, 1, 1), fill=None, width=er.pencil_size,
                             closePath=False, lineJoin=1)
            else:  # rect (default) or degenerate lasso
                shape.draw_rect(fitz.Rect(er.x0, er.y0, er.x1, er.y1))
                shape.finish(color=(1, 1, 1), fill=(1, 1, 1), width=0)
            shape.commit()

        # Burn geometric shapes
        for sh in self._shapes:
            page: fitz.Page = doc[sh.page_idx]
            s = page.new_shape()
            fill = sh.fill_color if sh.fill_color else None
            lw   = sh.line_width
            if sh.kind == "rect":
                s.draw_rect(fitz.Rect(sh.x0, sh.y0, sh.x1, sh.y1))
                s.finish(color=sh.stroke_color, fill=fill, width=lw)
            elif sh.kind == "rounded_rect":
                r = min(abs(sh.x1 - sh.x0), abs(sh.y1 - sh.y0)) * 0.15
                x0r, y0r, x1r, y1r = sh.x0, sh.y0, sh.x1, sh.y1
                P = fitz.Point
                # 8 path segments: 4 straight sides + 4 cubic-bezier corners
                s.draw_line(P(x0r + r, y0r), P(x1r - r, y0r))          # top
                s.draw_curve(P(x1r - r, y0r), P(x1r, y0r), P(x1r, y0r + r))  # top-right
                s.draw_line(P(x1r, y0r + r), P(x1r, y1r - r))          # right
                s.draw_curve(P(x1r, y1r - r), P(x1r, y1r), P(x1r - r, y1r))  # bottom-right
                s.draw_line(P(x1r - r, y1r), P(x0r + r, y1r))          # bottom
                s.draw_curve(P(x0r + r, y1r), P(x0r, y1r), P(x0r, y1r - r))  # bottom-left
                s.draw_line(P(x0r, y1r - r), P(x0r, y0r + r))          # left
                s.draw_curve(P(x0r, y0r + r), P(x0r, y0r), P(x0r + r, y0r))  # top-left
                s.finish(color=sh.stroke_color, fill=fill, width=lw, closePath=True)
            elif sh.kind == "ellipse":
                s.draw_oval(fitz.Rect(sh.x0, sh.y0, sh.x1, sh.y1))
                s.finish(color=sh.stroke_color, fill=fill, width=lw)
            elif sh.kind == "line":
                s.draw_line(fitz.Point(sh.x0, sh.y0), fitz.Point(sh.x1, sh.y1))
                s.finish(color=sh.stroke_color, width=lw, closePath=False)
            elif sh.kind == "arrow":
                # Draw line + filled arrowhead
                s.draw_line(fitz.Point(sh.x0, sh.y0), fitz.Point(sh.x1, sh.y1))
                s.finish(color=sh.stroke_color, width=lw, closePath=False)
                # Arrowhead
                import math as _math
                dx = sh.x1 - sh.x0
                dy = sh.y1 - sh.y0
                length = _math.hypot(dx, dy) or 1
                ux, uy = dx / length, dy / length
                head  = lw * 4
                wing  = lw * 2.5
                tip   = fitz.Point(sh.x1, sh.y1)
                base  = fitz.Point(sh.x1 - ux * head, sh.y1 - uy * head)
                left  = fitz.Point(base.x + uy * wing, base.y - ux * wing)
                right = fitz.Point(base.x - uy * wing, base.y + ux * wing)
                s2 = page.new_shape()
                s2.draw_polyline([tip, left, right])
                s2.finish(color=sh.stroke_color, fill=sh.stroke_color,
                           width=0, closePath=True)
                s2.commit()
            s.commit()

        # Burn pasted region clips
        for clip in self._clips:
            page: fitz.Page = doc[clip.page_idx]
            buf = io.BytesIO()
            clip.image.save(buf, format="PNG")
            buf.seek(0)
            rect = fitz.Rect(clip.x_pdf, clip.y_pdf,
                             clip.x_pdf + clip.w_pdf,
                             clip.y_pdf + clip.h_pdf)
            page.insert_image(rect, stream=buf.read())

        # Burn text placements
        errors = []
        for pl in self._placements:
            # Prefer live var value over cached pl.text (belt-and-suspenders)
            text = pl.var.get() if pl.var is not None else pl.text
            if not text.strip():
                continue
            page: fitz.Page = doc[pl.page_idx]
            try:
                if pl.kind == "check":
                    self._draw_check_shape(page, pl)
                else:
                    rect = fitz.Rect(
                        pl.x_pdf,
                        pl.y_pdf,
                        pl.x_pdf + pl.width_pdf,
                        pl.y_pdf + pl.font_size * 2.0,
                    )
                    font_path = _find_unicode_font(pl.bold)
                    if font_path:
                        fn = "UtextBold" if pl.bold else "UtextReg"
                        page.insert_textbox(
                            rect, text,
                            fontname=fn, fontfile=font_path,
                            fontsize=pl.font_size,
                            color=(0, 0, 0), align=0,
                        )
                    else:
                        fontname = "hebo" if pl.bold else "helv"
                        page.insert_textbox(
                            rect, text,
                            fontname=fontname,
                            fontsize=pl.font_size,
                            color=(0, 0, 0), align=0,
                        )
            except Exception as exc:
                errors.append(str(exc))
        if errors:
            messagebox.showwarning("Export warning",
                                   f"{len(errors)} item(s) could not be written:\n" +
                                   "\n".join(errors[:5]))

    def _draw_check_shape(self, page: fitz.Page, pl: "TextPlacement") -> None:
        """Draw a geometric ✓ or ✗ centered on pl.x_pdf / pl.y_pdf."""
        s = pl.font_size * 0.85          # overall symbol size in PDF points
        lw = max(1.2, pl.font_size * 0.1)  # stroke width
        # x_pdf/y_pdf is the CENTER of the symbol on click; shift to top-left
        x = pl.x_pdf - s / 2
        y = pl.y_pdf - s / 2
        shape = page.new_shape()
        if pl.text == "✓":
            shape.draw_polyline([
                fitz.Point(x,            y + s * 0.50),
                fitz.Point(x + s * 0.35, y + s),
                fitz.Point(x + s,        y + s * 0.10),
            ])
            shape.finish(color=(0, 0.55, 0), width=lw, closePath=False)
        else:  # ✗
            shape.draw_line(fitz.Point(x,     y),     fitz.Point(x + s, y + s))
            shape.draw_line(fitz.Point(x + s, y),     fitz.Point(x,     y + s))
            shape.finish(color=(0.75, 0, 0), width=lw, closePath=False)
        shape.commit()

    # ------------------------------------------------------------------
    # Page rendering
    # ------------------------------------------------------------------

    def _render_page(self) -> None:
        """Re-render the canvas — delegates to single-page or continuous mode."""
        if not self._doc:
            return

        if self._continuous_var.get():
            self._render_continuous()
        else:
            self._render_single()

    def _render_single(self) -> None:
        """Render only the current page (classic single-page view)."""
        self._destroy_all_overlays()
        self._canvas.delete("all")
        self._acro_widgets.clear()
        self._page_offsets.clear()
        self._photo_images.clear()

        page: fitz.Page = self._doc[self._current_page]
        mat = fitz.Matrix(self._zoom, self._zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self._photo_image = ImageTk.PhotoImage(img)

        p = self._PAD
        self._canvas.create_image(p, p, anchor=tk.NW, image=self._photo_image)
        self._canvas.configure(scrollregion=(0, 0, pix.width + p * 2, pix.height + p * 2))
        self._page_offsets[self._current_page] = (p, p)

        # AcroForm widgets
        self._overlay_acro_fields(page, p)

        # Restore free-text placements for this page
        for pl in self._placements:
            if pl.page_idx == self._current_page:
                self._create_text_widget(pl)

        # Restore shape overlays for this page
        for sh in self._shapes:
            if sh.page_idx == self._current_page:
                self._draw_shape_preview(sh, p, p)

        # Restore pasted clip overlays for this page
        for cl in self._clips:
            if cl.page_idx == self._current_page:
                self._create_clip_widget(cl, p, p)

    def _render_continuous(self) -> None:
        """Render all pages stacked vertically in a single scrollable canvas."""
        self._destroy_all_overlays()
        self._canvas.delete("all")
        self._acro_widgets.clear()
        self._page_offsets.clear()
        self._photo_images.clear()

        p = self._PAD
        y_cursor = p
        max_w = 0

        for pg_idx in range(len(self._doc)):
            page: fitz.Page = self._doc[pg_idx]
            mat = fitz.Matrix(self._zoom, self._zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            photo = ImageTk.PhotoImage(img)
            self._photo_images.append(photo)

            self._canvas.create_image(p, y_cursor, anchor=tk.NW, image=photo)
            self._page_offsets[pg_idx] = (p, y_cursor)
            max_w = max(max_w, pix.width)

            # AcroForm widgets for this page
            self._overlay_acro_fields(page, p, y_offset=y_cursor)

            # Free-text placements for this page
            for pl in self._placements:
                if pl.page_idx == pg_idx:
                    self._create_text_widget_at(pl, p, y_cursor)

            # Shape overlays for this page
            for sh in self._shapes:
                if sh.page_idx == pg_idx:
                    self._draw_shape_preview(sh, p, y_cursor)

            # Pasted clip overlays for this page
            for cl in self._clips:
                if cl.page_idx == pg_idx:
                    self._create_clip_widget(cl, p, y_cursor)

            y_cursor += pix.height + p  # gap between pages

        self._canvas.configure(
            scrollregion=(0, 0, max_w + p * 2, y_cursor)
        )

        # Scroll so the current page is visible
        if self._current_page in self._page_offsets:
            _, cy = self._page_offsets[self._current_page]
            total_h = self._canvas.bbox("all")
            if total_h:
                total_height = total_h[3]
                frac = cy / total_height
                self._canvas.yview_moveto(max(0.0, frac - 0.02))

    def _overlay_acro_fields(self, page: fitz.Page, offset: int, y_offset: int | None = None) -> None:
        """Overlay interactive Entry/checkbox widgets for any AcroForm fields."""
        x_off = offset
        y_off = y_offset if y_offset is not None else offset
        for w in page.widgets():
            r = w.rect
            x0 = r.x0 * self._zoom + x_off
            y0 = r.y0 * self._zoom + y_off
            x1 = r.x1 * self._zoom + x_off
            y1 = r.y1 * self._zoom + y_off
            W = max(int(x1 - x0), 22)
            H = max(int(y1 - y0), 16)
            ftype = w.field_type
            fname = w.field_name or f"_anon_{id(w)}"
            fval = w.field_value or ""
            font = ("Arial", max(7, int(9 * self._zoom)))
            bg = "#FFFACD"

            if ftype == fitz.PDF_WIDGET_TYPE_TEXT:
                if H >= int(38 * self._zoom):
                    frame = tk.Frame(self._canvas, bg=bg, bd=0)
                    txt = tk.Text(frame, font=font, relief=tk.FLAT, bg=bg,
                                  wrap=tk.WORD, bd=0, highlightthickness=0)
                    txt.insert("1.0", str(fval))
                    txt.pack(fill=tk.BOTH, expand=True)
                    self._canvas.create_window(x0, y0, anchor=tk.NW, window=frame, width=W, height=H)
                    self._acro_widgets[fname] = (w, txt, "multitext")
                else:
                    var = tk.StringVar(value=str(fval))
                    entry = tk.Entry(self._canvas, textvariable=var, font=font,
                                     relief=tk.FLAT, bg=bg, bd=0, highlightthickness=0)
                    self._canvas.create_window(x0, y0, anchor=tk.NW, window=entry, width=W, height=H)
                    self._acro_widgets[fname] = (w, var, "text")
            elif ftype == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                var = tk.BooleanVar(value=_field_is_checked(fval))
                cb = tk.Checkbutton(self._canvas, variable=var, bg=bg,
                                    activebackground=bg, relief=tk.FLAT, bd=0, highlightthickness=0)
                self._canvas.create_window(x0, y0, anchor=tk.NW, window=cb, width=W, height=H)
                self._acro_widgets[fname] = (w, var, "checkbox")
            elif ftype == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                on_val = "Yes"
                try:
                    on_val = w.button_states().get("normal", "Yes")
                except Exception:
                    pass
                var = tk.BooleanVar(value=(str(fval) == on_val or _field_is_checked(fval)))
                rb = tk.Checkbutton(self._canvas, variable=var, bg=bg,
                                    activebackground=bg, relief=tk.FLAT, bd=0, highlightthickness=0)
                self._canvas.create_window(x0, y0, anchor=tk.NW, window=rb, width=W, height=H)
                self._acro_widgets[fname] = (w, var, "radio")
            elif ftype in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX):
                choices = w.choice_values or []
                var = tk.StringVar(value=str(fval))
                combo = ttk.Combobox(self._canvas, textvariable=var, values=choices, font=font)
                self._canvas.create_window(x0, y0, anchor=tk.NW, window=combo, width=W, height=H)
                self._acro_widgets[fname] = (w, var, "combo")

    # ------------------------------------------------------------------
    # Free-text placement
    # ------------------------------------------------------------------

    def _page_at_canvas(self, cx: float, cy: float) -> tuple[int, float, float] | None:
        """Return (page_idx, x_pdf, y_pdf) for the given canvas coords, or None.

        Iterates over the rendered page bounding boxes (_page_offsets) and returns
        the first match.  Works for both single-page and continuous-scroll modes.
        """
        for pg_idx, (ox, oy) in self._page_offsets.items():
            page = self._doc[pg_idx]
            pw = page.rect.width * self._zoom
            ph = page.rect.height * self._zoom
            if ox <= cx <= ox + pw and oy <= cy <= oy + ph:
                return pg_idx, (cx - ox) / self._zoom, (cy - oy) / self._zoom
        return None

    def _on_canvas_click(self, event: tk.Event) -> None:
        """Handle a left-click on the canvas: start an eraser/shape/copy drag or place a new item."""
        if not self._doc:
            return
        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)

        mode = self._mode_var.get()

        # Eraser mode: start drag, don't place text
        if mode == "eraser":
            self._eraser_start = (cx, cy)
            self._eraser_lasso_points = [(cx, cy)]
            return

        # Shape draw mode: start drag
        if mode == "shape":
            self._shape_start = (cx, cy)
            return

        # Copy-region mode: start drag
        if mode == "copy_region":
            self._copy_start = (cx, cy)
            return

        # Paste mode: place clipboard image at click position
        if mode == "paste":
            self._paste_at(cx, cy)
            return

        # Determine which page was clicked
        hit = self._page_at_canvas(cx, cy)
        if hit is None:
            return
        pg_idx, x_pdf, y_pdf = hit

        # If click landed on an existing canvas window widget, don't add new text
        for item in self._canvas.find_overlapping(cx - 2, cy - 2, cx + 2, cy + 2):
            if self._canvas.type(item) == "window":
                return

        try:
            font_size = float(self._fs_var.get())
        except ValueError:
            font_size = self._DEFAULT_FS

        if mode == "check_v":
            kind, text = "check", "✓"
        elif mode == "check_x":
            kind, text = "check", "✗"
        else:
            kind, text = "text", ""

        pl = TextPlacement(
            page_idx=pg_idx,
            x_pdf=x_pdf,
            y_pdf=y_pdf,
            text=text,
            font_size=font_size,
            bold=self._bold_var.get(),
            kind=kind,
        )
        self._placements.append(pl)
        self._undo_stack.append(pl)
        # Update current_page to follow the click (for AcroForm flush etc.)
        self._current_page = pg_idx
        self._refresh_controls()
        ox, oy = self._page_offsets[pg_idx]
        self._create_text_widget_at(pl, ox, oy)

        # Auto-focus new text entries (not checkmarks)
        if kind == "text" and pl.frame:
            for child in pl.frame.winfo_children():
                if isinstance(child, tk.Entry):
                    child.focus_set()
                    break

    def _get_grip_w(self) -> int:
        """Return the pixel width of the grip label (measured once, then cached)."""
        if self._grip_w_cache == 0:
            tmp = tk.Label(self, text="⠿", font=("Arial", 9), padx=2, pady=0)
            self._grip_w_cache = tmp.winfo_reqwidth() + 1  # +1 for frame bd=1
            tmp.destroy()
        return self._grip_w_cache

    def _get_right_ctrl_w(self) -> int:
        """Return the combined pixel width of the three right-side controls (×, ⠿, ⇔)."""
        if self._right_ctrl_w_cache == 0:
            d = tk.Label(self, text="×",  font=("Arial", 10, "bold"), padx=3, pady=0)
            g = tk.Label(self, text="⠿",  font=("Arial", 9),          padx=2, pady=0)
            r = tk.Label(self, text="⇔",  font=("Arial", 9),          padx=2, pady=0)
            self.update_idletasks()
            self._right_ctrl_w_cache = (
                d.winfo_reqwidth() + g.winfo_reqwidth() + r.winfo_reqwidth() + 2
            )
            d.destroy(); g.destroy(); r.destroy()
        return self._right_ctrl_w_cache

    def _create_text_widget(self, pl: TextPlacement) -> None:
        """Create a draggable overlay widget on the canvas for a placement.

        Uses the stored page offset from _page_offsets when available, so it
        works correctly for both single-page and continuous-scroll modes.
        """
        if pl.page_idx in self._page_offsets:
            ox, oy = self._page_offsets[pl.page_idx]
        else:
            ox = oy = self._PAD
        self._create_text_widget_at(pl, ox, oy)

    def _create_text_widget_at(self, pl: TextPlacement, ox: int, oy: int) -> None:
        """Create a draggable overlay widget using explicit page origin (ox, oy)."""
        # content position in canvas coordinates
        cx = pl.x_pdf * self._zoom + ox
        cy = pl.y_pdf * self._zoom + oy
        font_px = max(8, int(pl.font_size * self._zoom * 0.75))

        if pl.kind == "check":
            self._create_check_widget(pl, cx, cy, font_px)
        else:
            self._create_entry_widget(pl, cx, cy, font_px)

    def _create_entry_widget(self, pl: TextPlacement, cx: float, cy: float,
                              font_px: int) -> None:
        """Build the draggable Entry widget.

        Layout (left to right): [Entry text field] [× delete] [⣿ drag] [⇔ resize]
        The frame is placed with its NW corner at (cx, cy).  The background is
        made transparent via a Win32 layered-window colour key so the PDF shows
        through beneath the text.
        """
        _TKEY = "#FEFFFE"  # near-white colour used as Win32 transparency key

        rcw        = self._get_right_ctrl_w()
        content_w  = max(60, int(pl.width_pdf * self._zoom))
        font_spec  = ("Arial", font_px, "bold" if pl.bold else "normal")
        fnt        = tkfont.Font(family="Arial", size=font_px,
                                 weight="bold" if pl.bold else "normal")

        frame = tk.Frame(self._canvas, bg=_TKEY, bd=0, relief=tk.FLAT, cursor="arrow")

        var   = tk.StringVar(value=pl.text)
        entry = tk.Entry(frame, textvariable=var, font=font_spec,
                         relief=tk.FLAT, bg=_TKEY, fg="#000000",
                         bd=0, highlightthickness=1,
                         highlightbackground="#aaaaaa", highlightcolor="#0066cc",
                         width=1, insertwidth=2)
        entry.pack(side=tk.LEFT, padx=0, fill=tk.X, expand=True)

        del_btn = tk.Label(frame, text="×", fg="#cc0000", bg=_TKEY, cursor="hand2",
                           font=("Arial", 10, "bold"), padx=3, pady=0)
        del_btn.pack(side=tk.LEFT)

        grip = tk.Label(frame, text="⣿", bg="#FFD966", cursor="fleur",
                        font=("Arial", 9), padx=2, pady=0)
        grip.pack(side=tk.LEFT)

        resize_grip = tk.Label(frame, text="⇔", bg="#b0c4de",
                               cursor="sb_h_double_arrow",
                               font=("Arial", 9), padx=2, pady=0)
        resize_grip.pack(side=tk.LEFT)

        # Frame NW is placed exactly at the click point (cx, cy).
        win_id = self._canvas.create_window(cx, cy, anchor=tk.NW,
                                            window=frame, width=content_w + rcw)
        pl.canvas_win_id = win_id
        pl.frame  = frame
        pl.var    = var
        pl.entry  = entry

        # Auto-expand to fit typed text.
        def _on_text_change(*_):
            pl.text = var.get()
            text_w  = fnt.measure(var.get() or "") + 10
            new_cw  = max(40, text_w)
            self._canvas.itemconfigure(pl.canvas_win_id, width=new_cw + rcw)
            pl.width_pdf = new_cw / self._zoom

        var.trace_add("write", _on_text_change)

        # Apply Win32 colour-key transparency once the widget is realised.
        def _apply_transparency():
            try:
                GWL_EXSTYLE   = -20
                WS_EX_LAYERED = 0x00080000
                LWA_COLORKEY  = 0x00000001
                r_c, g_c, b_c = [c >> 8 for c in frame.winfo_rgb(_TKEY)]
                cref = (b_c << 16) | (g_c << 8) | r_c
                for w in (frame, entry, del_btn):
                    hwnd  = w.winfo_id()
                    style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
                    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                                                        style | WS_EX_LAYERED)
                    ctypes.windll.user32.SetLayeredWindowAttributes(
                        hwnd, cref, 255, LWA_COLORKEY)
            except Exception:
                pass

        frame.after(100, _apply_transparency)
        # Register the new overlay frame (and its children) as a drop target so
        # dragging a PDF over a text widget still opens it.
        self._register_drop_recursive(frame)

        del_btn.bind("<Button-1>", lambda _e, _pl=pl: self._delete_placement(_pl))
        for widget in (grip, frame):
            widget.bind("<ButtonPress-1>", lambda e, _pl=pl: self._drag_start(e, _pl))
            widget.bind("<B1-Motion>",     lambda e, _pl=pl: self._drag_move(e,  _pl))
        resize_grip.bind("<ButtonPress-1>", lambda e, _pl=pl: self._resize_start(e, _pl))
        resize_grip.bind("<B1-Motion>",     lambda e, _pl=pl: self._resize_move(e,  _pl))

    def _create_check_widget(self, pl: TextPlacement, cx: float, cy: float,
                              font_px: int) -> None:
        """Place a ✓ or ✗ symbol directly as a canvas text item (no tk widget).

        Unlike text placements there is no draggable frame — use Ctrl+Z to undo.
        The symbol is centred on the click point (cx, cy).
        """
        # Draw directly as a canvas text item — no widget, no background.
        fill = "#006600" if pl.text == "✓" else "#990000"
        item_id = self._canvas.create_text(
            cx, cy, text=pl.text, fill=fill,
            font=("Arial", font_px, "bold"), anchor=tk.CENTER,
        )
        pl.canvas_win_id = item_id
        pl.frame = None
        pl.var = None
        pl.entry = None

    def _delete_placement(self, pl: TextPlacement) -> None:
        """Remove a placement from the canvas and from the internal lists."""
        if pl.canvas_win_id:
            self._canvas.delete(pl.canvas_win_id)
            pl.canvas_win_id = 0
        if pl.frame:
            pl.frame.destroy()
            pl.frame = None
        try:
            self._placements.remove(pl)
        except ValueError:
            pass
        # Also remove from undo stack so Ctrl+Z won't try to re-delete it
        try:
            self._undo_stack.remove(pl)
        except ValueError:
            pass

    def _undo(self) -> None:
        """Remove the most recently added placement, eraser, shape, or clip."""
        if not self._undo_stack:
            return
        item = self._undo_stack.pop()
        if isinstance(item, EraserRect):
            try:
                self._erasers.remove(item)
            except ValueError:
                pass
            # Remove visual preview rectangle for this eraser (if still on canvas)
            ox = oy = self._PAD
            if item.page_idx in self._page_offsets:
                ox, oy = self._page_offsets[item.page_idx]
            z = self._zoom
            rx0 = item.x0 * z + ox
            ry0 = item.y0 * z + oy
            rx1 = item.x1 * z + ox
            ry1 = item.y1 * z + oy
            for cid in self._canvas.find_withtag("eraser_preview"):
                coords = self._canvas.coords(cid)
                if coords and abs(coords[0] - rx0) < 2 and abs(coords[1] - ry0) < 2:
                    self._canvas.delete(cid)
                    break
        elif isinstance(item, ShapeRect):
            try:
                self._shapes.remove(item)
            except ValueError:
                pass
            if item.canvas_id:
                self._canvas.delete(item.canvas_id)
                item.canvas_id = 0
        elif isinstance(item, RegionClip):
            self._delete_clip(item)
        else:
            self._delete_placement(item)

    def _drag_start(self, event: tk.Event, pl: TextPlacement) -> None:
        """Record the starting position of a drag operation on a text placement."""
        coords = self._canvas.coords(pl.canvas_win_id) if pl.canvas_win_id else []
        win_x = coords[0] if coords else pl.x_pdf * self._zoom + self._PAD
        win_y = coords[1] if coords else pl.y_pdf * self._zoom + self._PAD
        self._drag_data = {
            "pl": pl,
            "mouse_x": event.x_root,
            "mouse_y": event.y_root,
            "win_x": win_x,
            "win_y": win_y,
        }

    def _drag_move(self, event: tk.Event, pl: TextPlacement) -> None:
        """Move a text placement in response to a mouse drag on its grip handle."""
        dd = self._drag_data
        if not dd or dd.get("pl") is not pl:
            return
        new_x = dd["win_x"] + (event.x_root - dd["mouse_x"])
        new_y = dd["win_y"] + (event.y_root - dd["mouse_y"])
        self._canvas.coords(pl.canvas_win_id, new_x, new_y)
        # Frame NW is now at the content left edge (entry is leftmost element).
        ox, oy = self._page_offsets.get(pl.page_idx, (self._PAD, self._PAD))
        pl.x_pdf = (new_x - ox) / self._zoom
        pl.y_pdf = (new_y - oy) / self._zoom

    def _resize_start(self, event: tk.Event, pl: TextPlacement) -> None:
        """Record the starting width for a horizontal resize drag on a text placement."""
        self._resize_data = {
            "pl": pl,
            "mouse_x": event.x_root,
            "start_w": max(60, int(pl.width_pdf * self._zoom)),
        }

    def _resize_move(self, event: tk.Event, pl: TextPlacement) -> None:
        """Expand or shrink the text entry widget as the resize handle is dragged."""
        rd = self._resize_data
        if not rd or rd.get("pl") is not pl:
            return
        rcw = self._get_right_ctrl_w()
        new_content_w = max(60, rd["start_w"] + (event.x_root - rd["mouse_x"]))
        self._canvas.itemconfigure(pl.canvas_win_id, width=new_content_w + rcw)
        pl.width_pdf = new_content_w / self._zoom

    # ------------------------------------------------------------------
    # Overlay / flush lifecycle
    # ------------------------------------------------------------------

    def _destroy_all_overlays(self) -> None:
        """Destroy all live tk widget overlays (used before a full re-render)."""
        for pl in self._placements:
            if pl.frame:
                pl.frame.destroy()
                pl.frame = None
                pl.canvas_win_id = 0
                pl.var = None
                pl.entry = None
        # Reset shape canvas ids (the canvas is cleared, items no longer exist)
        for sh in self._shapes:
            sh.canvas_id = 0
        # Destroy clip overlay frames
        for cl in self._clips:
            if cl.frame:
                cl.frame.destroy()
                cl.frame = None
                cl.canvas_win_id = 0

    def _destroy_current_overlays(self) -> None:
        """Destroy live tk widgets for the current page before re-rendering."""
        for pl in self._placements:
            if pl.page_idx == self._current_page and pl.frame:
                pl.frame.destroy()
                pl.frame = None
                pl.canvas_win_id = 0
                pl.var = None
                pl.entry = None

    def _flush_page(self) -> None:
        """Persist Entry values into placements; write AcroForm field changes."""
        for pl in self._placements:
            if pl.var is not None:   # flush all pages that still have live vars
                pl.text = pl.var.get()

        if not self._doc or not self._acro_widgets:
            return
        page: fitz.Page = self._doc[self._current_page]
        for w in page.widgets():
            fname = w.field_name or f"_anon_{id(w)}"
            if fname not in self._acro_widgets:
                continue
            _, var, kind = self._acro_widgets[fname]
            try:
                if kind == "text":
                    w.field_value = var.get()
                elif kind == "multitext":
                    w.field_value = var.get("1.0", tk.END).rstrip("\n")
                elif kind in ("checkbox", "radio"):
                    w.field_value = "Yes" if var.get() else "Off"
                elif kind == "combo":
                    w.field_value = var.get()
                w.update()
            except Exception:
                pass

    def _clear_page(self) -> None:
        """Remove all free-text placements, shapes, and clips from the current page."""
        for pl in [p for p in self._placements if p.page_idx == self._current_page]:
            if pl.canvas_win_id:
                self._canvas.delete(pl.canvas_win_id)
            if pl.frame:
                pl.frame.destroy()
                pl.frame = None
            self._placements.remove(pl)
            try:
                self._undo_stack.remove(pl)
            except ValueError:
                pass
        # Remove erasers on this page
        removed_erasers = [e for e in self._erasers if e.page_idx == self._current_page]
        self._erasers = [e for e in self._erasers if e.page_idx != self._current_page]
        for er in removed_erasers:
            try:
                self._undo_stack.remove(er)
            except ValueError:
                pass
        # Remove shapes on this page
        removed_shapes = [s for s in self._shapes if s.page_idx == self._current_page]
        self._shapes = [s for s in self._shapes if s.page_idx != self._current_page]
        for sh in removed_shapes:
            if sh.canvas_id:
                self._canvas.delete(sh.canvas_id)
            try:
                self._undo_stack.remove(sh)
            except ValueError:
                pass
        # Remove pasted clips on this page
        removed_clips = [c for c in self._clips if c.page_idx == self._current_page]
        self._clips = [c for c in self._clips if c.page_idx != self._current_page]
        for cl in removed_clips:
            if cl.canvas_win_id:
                self._canvas.delete(cl.canvas_win_id)
            if cl.frame:
                cl.frame.destroy()
            try:
                self._undo_stack.remove(cl)
            except ValueError:
                pass

    # ------------------------------------------------------------------
    # Eraser canvas events
    # ------------------------------------------------------------------

    def _on_canvas_drag(self, event: tk.Event) -> None:
        """Draw a rubber-band overlay while the user drags in eraser/shape/copy mode."""
        mode = self._mode_var.get()
        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)

        if mode == "eraser":
            if self._eraser_start is None:
                return
            x0, y0 = self._eraser_start
            ekind = self._eraser_kind_var.get()
            if ekind == "circle":
                if self._eraser_rect_id:
                    self._canvas.coords(self._eraser_rect_id, x0, y0, cx, cy)
                else:
                    self._eraser_rect_id = self._canvas.create_oval(
                        x0, y0, cx, cy, outline="#0066cc", width=2, dash=(4, 3))
            elif ekind in ("lasso", "free"):
                # Accumulate points for a freehand polygon / stroke
                self._eraser_lasso_points.append((cx, cy))
                # Redraw the polyline each motion event
                self._canvas.delete("eraser_lasso_preview")
                pts = self._eraser_lasso_points
                if len(pts) >= 2:
                    flat = [v for p in pts for v in p]
                    if ekind == "free":
                        px_w = max(2, self._eraser_pencil_size_var.get())
                        self._canvas.create_line(*flat, fill="#99ccff", width=px_w,
                                                 capstyle=tk.ROUND, joinstyle=tk.ROUND,
                                                 tags="eraser_lasso_preview")
                    else:
                        self._canvas.create_line(*flat, fill="#0066cc", width=2,
                                                 dash=(4, 3), tags="eraser_lasso_preview")
            else:  # rect (default)
                if self._eraser_rect_id:
                    self._canvas.coords(self._eraser_rect_id, x0, y0, cx, cy)
                else:
                    self._eraser_rect_id = self._canvas.create_rectangle(
                        x0, y0, cx, cy, outline="#0066cc", width=2, dash=(4, 3))

        elif mode == "shape":
            if self._shape_start is None:
                return
            x0, y0 = self._shape_start
            kind = self._shape_kind_var.get()
            stroke = self._shape_stroke_var.get()
            if self._shape_preview_id:
                self._canvas.coords(self._shape_preview_id, x0, y0, cx, cy)
            else:
                if kind == "ellipse":
                    self._shape_preview_id = self._canvas.create_oval(
                        x0, y0, cx, cy, outline=stroke, width=2, dash=(4, 3))
                elif kind in ("line", "arrow"):
                    self._shape_preview_id = self._canvas.create_line(
                        x0, y0, cx, cy, fill=stroke, width=2, dash=(4, 3),
                        arrow=tk.LAST if kind == "arrow" else tk.NONE)
                else:
                    self._shape_preview_id = self._canvas.create_rectangle(
                        x0, y0, cx, cy, outline=stroke, width=2, dash=(4, 3))

        elif mode == "copy_region":
            if self._copy_start is None:
                return
            x0, y0 = self._copy_start
            if self._copy_preview_id:
                self._canvas.coords(self._copy_preview_id, x0, y0, cx, cy)
            else:
                self._copy_preview_id = self._canvas.create_rectangle(
                    x0, y0, cx, cy, outline="#ff6600", width=2, dash=(6, 3))

    def _on_canvas_release(self, event: tk.Event) -> None:
        """Finalise an eraser/shape/copy-region drag on mouse-button release."""
        mode = self._mode_var.get()
        cx = self._canvas.canvasx(event.x)
        cy = self._canvas.canvasy(event.y)

        if mode == "eraser" and self._eraser_start is not None:
            self._finish_eraser(cx, cy)
        elif mode == "shape" and self._shape_start is not None:
            self._finish_shape(cx, cy)
        elif mode == "copy_region" and self._copy_start is not None:
            self._finish_copy_region(cx, cy)

    # ------------------------------------------------------------------
    # Eraser finalise
    # ------------------------------------------------------------------

    def _finish_eraser(self, cx: float, cy: float) -> None:
        """Commit the eraser shape (rect, circle, lasso, or free)."""
        x0, y0 = self._eraser_start
        self._eraser_start = None
        ekind = self._eraser_kind_var.get()

        # Clean up rubber-band / lasso preview
        if self._eraser_rect_id:
            self._canvas.delete(self._eraser_rect_id)
            self._eraser_rect_id = 0
        self._canvas.delete("eraser_lasso_preview")

        z = self._zoom

        if ekind in ("lasso", "free"):
            pts_canvas = self._eraser_lasso_points
            self._eraser_lasso_points = []
            min_pts = 2 if ekind == "free" else 3
            if len(pts_canvas) < min_pts:
                return
            # Determine which page contains the centroid of the path
            xs = [p[0] for p in pts_canvas]
            ys = [p[1] for p in pts_canvas]
            min_x, max_x = min(xs), max(xs)
            min_y, max_y = min(ys), max(ys)
            mid_x = (min_x + max_x) / 2
            mid_y = (min_y + max_y) / 2
            pg_idx = self._current_page
            for pidx, (ox, oy) in self._page_offsets.items():
                page_obj = self._doc[pidx]
                pw = page_obj.rect.width * z
                ph = page_obj.rect.height * z
                if ox <= mid_x <= ox + pw and oy <= mid_y <= oy + ph:
                    pg_idx = pidx
                    break
            ox, oy = self._page_offsets.get(pg_idx, (self._PAD, self._PAD))
            pdf_pts = [((p[0] - ox) / z, (p[1] - oy) / z) for p in pts_canvas]
            x0_pdf = min(p[0] for p in pdf_pts)
            y0_pdf = min(p[1] for p in pdf_pts)
            x1_pdf = max(p[0] for p in pdf_pts)
            y1_pdf = max(p[1] for p in pdf_pts)
            if ekind == "free":
                px_w = max(1, self._eraser_pencil_size_var.get())
                pencil_pdf = px_w / z
                er = EraserRect(
                    page_idx=pg_idx,
                    x0=x0_pdf, y0=y0_pdf, x1=x1_pdf, y1=y1_pdf,
                    kind="free",
                    points=pdf_pts,
                    pencil_size=pencil_pdf,
                )
                self._erasers.append(er)
                self._undo_stack.append(er)
                # Canvas preview: thick white round-capped stroke
                flat = [v for p in pts_canvas for v in p]
                self._canvas.create_line(*flat, fill="white", width=px_w,
                                         capstyle=tk.ROUND, joinstyle=tk.ROUND,
                                         tags="eraser_preview")
            else:  # lasso
                er = EraserRect(
                    page_idx=pg_idx,
                    x0=x0_pdf, y0=y0_pdf, x1=x1_pdf, y1=y1_pdf,
                    kind="lasso",
                    points=pdf_pts,
                )
                self._erasers.append(er)
                self._undo_stack.append(er)
                # Canvas preview: white filled polygon
                flat = [v for p in pts_canvas for v in p]
                self._canvas.create_polygon(*flat, fill="white", outline="",
                                            tags="eraser_preview")
            return

        # rect / circle: use bounding box
        rx0, rx1 = sorted([x0, cx])
        ry0, ry1 = sorted([y0, cy])
        if rx1 - rx0 < 4 or ry1 - ry0 < 4:
            return
        mid_x = (rx0 + rx1) / 2
        mid_y = (ry0 + ry1) / 2
        pg_idx = self._current_page
        for pidx, (ox, oy) in self._page_offsets.items():
            page_obj = self._doc[pidx]
            pw = page_obj.rect.width * z
            ph = page_obj.rect.height * z
            if ox <= mid_x <= ox + pw and oy <= mid_y <= oy + ph:
                pg_idx = pidx
                break
        ox, oy = self._page_offsets.get(pg_idx, (self._PAD, self._PAD))
        er = EraserRect(
            page_idx=pg_idx,
            x0=(rx0 - ox) / z, y0=(ry0 - oy) / z,
            x1=(rx1 - ox) / z, y1=(ry1 - oy) / z,
            kind=ekind,
        )
        self._erasers.append(er)
        self._undo_stack.append(er)
        if ekind == "circle":
            self._canvas.create_oval(rx0, ry0, rx1, ry1,
                                     fill="white", outline="", tags="eraser_preview")
        else:
            self._canvas.create_rectangle(rx0, ry0, rx1, ry1,
                                           fill="white", outline="", tags="eraser_preview")

    # ------------------------------------------------------------------
    # Shape finalise
    # ------------------------------------------------------------------

    def _finish_shape(self, cx: float, cy: float) -> None:
        """Commit the drawn shape."""
        x0, y0 = self._shape_start
        self._shape_start = None
        if self._shape_preview_id:
            self._canvas.delete(self._shape_preview_id)
            self._shape_preview_id = 0
        z = self._zoom
        rx0, rx1 = sorted([x0, cx])
        ry0, ry1 = sorted([y0, cy])
        kind = self._shape_kind_var.get()
        # For lines/arrows a zero-area rect is fine; otherwise enforce minimum
        if kind not in ("line", "arrow") and (rx1 - rx0 < 4 or ry1 - ry0 < 4):
            return
        if kind in ("line", "arrow") and (abs(cx - x0) < 2 and abs(cy - y0) < 2):
            return
        mid_x = (rx0 + rx1) / 2
        mid_y = (ry0 + ry1) / 2
        pg_idx = self._current_page
        for pidx, (ox, oy) in self._page_offsets.items():
            page = self._doc[pidx]
            pw = page.rect.width * z
            ph = page.rect.height * z
            if ox <= mid_x <= ox + pw and oy <= mid_y <= oy + ph:
                pg_idx = pidx
                break
        ox, oy = self._page_offsets.get(pg_idx, (self._PAD, self._PAD))
        # For line/arrow keep original direction, not sorted
        if kind in ("line", "arrow"):
            lx0, ly0, lx1, ly1 = x0, y0, cx, cy
        else:
            lx0, ly0, lx1, ly1 = rx0, ry0, rx1, ry1
        stroke = self._hex_to_rgb(self._shape_stroke_var.get())
        fill   = self._hex_to_rgb(self._shape_fill_var.get()) if (
            self._shape_filled_var.get() and self._shape_fill_var.get()) else None
        try:
            lw = float(self._shape_lw_var.get())
        except ValueError:
            lw = 1.5
        sh = ShapeRect(
            page_idx=pg_idx,
            x0=(lx0 - ox) / z, y0=(ly0 - oy) / z,
            x1=(lx1 - ox) / z, y1=(ly1 - oy) / z,
            kind=kind,
            stroke_color=stroke,
            fill_color=fill,
            line_width=lw,
        )
        self._shapes.append(sh)
        self._undo_stack.append(sh)
        # Draw permanent preview on canvas
        self._draw_shape_preview(sh, ox, oy)

    def _draw_shape_preview(self, sh: ShapeRect, ox: int, oy: int) -> None:
        """Draw the canvas preview overlay for a committed ShapeRect."""
        z = self._zoom
        sx0 = sh.x0 * z + ox
        sy0 = sh.y0 * z + oy
        sx1 = sh.x1 * z + ox
        sy1 = sh.y1 * z + oy

        def _rgb_to_hex(rgb):
            return "#{:02x}{:02x}{:02x}".format(
                int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255))

        stroke_hex = _rgb_to_hex(sh.stroke_color)
        fill_hex   = _rgb_to_hex(sh.fill_color) if sh.fill_color else ""
        lw_px      = max(1, int(sh.line_width * z / 2))

        if sh.kind == "ellipse":
            cid = self._canvas.create_oval(
                sx0, sy0, sx1, sy1,
                outline=stroke_hex, fill=fill_hex or "", width=lw_px,
                tags="shape_preview")
        elif sh.kind == "line":
            cid = self._canvas.create_line(
                sx0, sy0, sx1, sy1,
                fill=stroke_hex, width=lw_px, tags="shape_preview")
        elif sh.kind == "arrow":
            cid = self._canvas.create_line(
                sx0, sy0, sx1, sy1,
                fill=stroke_hex, width=lw_px,
                arrow=tk.LAST, arrowshape=(lw_px * 5, lw_px * 6, lw_px * 2),
                tags="shape_preview")
        else:
            # rect and rounded_rect both shown as rectangle on canvas
            cid = self._canvas.create_rectangle(
                sx0, sy0, sx1, sy1,
                outline=stroke_hex, fill=fill_hex or "", width=lw_px,
                tags="shape_preview")
        sh.canvas_id = cid

    # ------------------------------------------------------------------
    # Copy-region finalise
    # ------------------------------------------------------------------

    def _finish_copy_region(self, cx: float, cy: float) -> None:
        """Rasterise the selected region from the current page into the clipboard."""
        x0, y0 = self._copy_start
        self._copy_start = None
        if self._copy_preview_id:
            self._canvas.delete(self._copy_preview_id)
            self._copy_preview_id = 0
        z = self._zoom
        rx0, rx1 = sorted([x0, cx])
        ry0, ry1 = sorted([y0, cy])
        if rx1 - rx0 < 8 or ry1 - ry0 < 8:
            return
        # Determine page
        mid_x = (rx0 + rx1) / 2
        mid_y = (ry0 + ry1) / 2
        pg_idx = self._current_page
        for pidx, (ox, oy) in self._page_offsets.items():
            page = self._doc[pidx]
            pw = page.rect.width * z
            ph = page.rect.height * z
            if ox <= mid_x <= ox + pw and oy <= mid_y <= oy + ph:
                pg_idx = pidx
                break
        ox, oy = self._page_offsets.get(pg_idx, (self._PAD, self._PAD))
        # PDF-space rect of the selection
        pdf_x0 = (rx0 - ox) / z
        pdf_y0 = (ry0 - oy) / z
        pdf_x1 = (rx1 - ox) / z
        pdf_y1 = (ry1 - oy) / z
        # Render that clip at current zoom resolution
        page = self._doc[pg_idx]
        clip = fitz.Rect(pdf_x0, pdf_y0, pdf_x1, pdf_y1)
        mat = fitz.Matrix(z, z)
        pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
        from PIL import Image as _Image
        img = _Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self._clipboard_image = img
        self._clipboard_size_pdf = (pdf_x1 - pdf_x0, pdf_y1 - pdf_y0)
        n = len([c for c in self._clips if c.page_idx == pg_idx]) + 1
        self._status_var.set(
            f"Region copied ({pix.width}×{pix.height} px)  —  switch to Paste mode and click to place")

    # ------------------------------------------------------------------
    # Paste
    # ------------------------------------------------------------------

    def _paste_at(self, cx: float, cy: float) -> None:
        """Place the clipboard image at the given canvas position."""
        if self._clipboard_image is None:
            self._status_var.set("Nothing in clipboard — use 'Copy region' first.")
            return
        hit = self._page_at_canvas(cx, cy)
        if hit is None:
            return
        pg_idx, x_pdf, y_pdf = hit
        w_pdf, h_pdf = self._clipboard_size_pdf
        clip = RegionClip(
            page_idx=pg_idx,
            x_pdf=x_pdf, y_pdf=y_pdf,
            w_pdf=w_pdf, h_pdf=h_pdf,
            image=self._clipboard_image.copy(),
        )
        self._clips.append(clip)
        self._undo_stack.append(clip)
        self._current_page = pg_idx
        self._refresh_controls()
        ox, oy = self._page_offsets[pg_idx]
        self._create_clip_widget(clip, ox, oy)

    def _create_clip_widget(self, clip: RegionClip, ox: int, oy: int) -> None:
        """Place a draggable image overlay on the canvas for a pasted RegionClip."""
        z = self._zoom
        cx = clip.x_pdf * z + ox
        cy = clip.y_pdf * z + oy
        disp_w = max(1, int(clip.w_pdf * z))
        disp_h = max(1, int(clip.h_pdf * z))

        # Resize the PIL image to current zoom for display
        thumb = clip.image.resize((disp_w, disp_h), resample=1)
        photo = ImageTk.PhotoImage(thumb)

        frame = tk.Frame(self._canvas, bd=1, relief=tk.RIDGE, cursor="fleur")
        lbl   = tk.Label(frame, image=photo, bd=0, cursor="fleur")
        lbl.image = photo   # prevent GC
        lbl.pack()

        del_btn = tk.Label(frame, text="×", fg="#cc0000", bg="#f0f0f0",
                           cursor="hand2", font=("Arial", 9, "bold"), padx=2)
        del_btn.place(relx=1.0, rely=0.0, anchor=tk.NE)

        win_id = self._canvas.create_window(cx, cy, anchor=tk.NW, window=frame,
                                             width=disp_w + 2, height=disp_h + 2)
        clip.canvas_win_id = win_id
        clip.frame = frame
        # Register the clip frame so dropping a PDF on top of it works.
        self._register_drop_recursive(frame)

        del_btn.bind("<Button-1>", lambda _e, _c=clip: self._delete_clip(_c))
        for w in (frame, lbl):
            w.bind("<ButtonPress-1>",  lambda e, _c=clip: self._clip_drag_start(e, _c))
            w.bind("<B1-Motion>",      lambda e, _c=clip: self._clip_drag_move(e,  _c))

    def _delete_clip(self, clip: RegionClip) -> None:
        """Remove a pasted clip from the canvas and from the internal list."""
        if clip.canvas_win_id:
            self._canvas.delete(clip.canvas_win_id)
            clip.canvas_win_id = 0
        if clip.frame:
            clip.frame.destroy()
            clip.frame = None
        try:
            self._clips.remove(clip)
        except ValueError:
            pass
        try:
            self._undo_stack.remove(clip)
        except ValueError:
            pass

    def _clip_drag_start(self, event: tk.Event, clip: RegionClip) -> None:
        coords = self._canvas.coords(clip.canvas_win_id) if clip.canvas_win_id else []
        self._clip_drag_data = {
            "clip": clip,
            "mouse_x": event.x_root,
            "mouse_y": event.y_root,
            "win_x": coords[0] if coords else clip.x_pdf * self._zoom + self._PAD,
            "win_y": coords[1] if coords else clip.y_pdf * self._zoom + self._PAD,
        }

    def _clip_drag_move(self, event: tk.Event, clip: RegionClip) -> None:
        dd = self._clip_drag_data
        if not dd or dd.get("clip") is not clip:
            return
        new_x = dd["win_x"] + (event.x_root - dd["mouse_x"])
        new_y = dd["win_y"] + (event.y_root - dd["mouse_y"])
        self._canvas.coords(clip.canvas_win_id, new_x, new_y)
        ox, oy = self._page_offsets.get(clip.page_idx, (self._PAD, self._PAD))
        clip.x_pdf = (new_x - ox) / self._zoom
        clip.y_pdf = (new_y - oy) / self._zoom

    # ------------------------------------------------------------------
    # Page operations
    # ------------------------------------------------------------------

    def _rotate_page(self, degrees: int) -> None:
        """Rotate the current page by +90 or -90 degrees (modifies self._doc)."""
        if not self._doc:
            return
        self._flush_page()
        page = self._doc[self._current_page]
        page.set_rotation((page.rotation + degrees) % 360)
        self._render_page()

    def _mirror_page(self, horizontal: bool) -> None:
        """Mirror the current page horizontally or vertically using a content stream."""
        if not self._doc:
            return
        self._flush_page()
        page = self._doc[self._current_page]
        w, h = page.rect.width, page.rect.height
        if horizontal:
            mat = fitz.Matrix(-1, 0, 0, 1, w, 0)  # flip x
        else:
            mat = fitz.Matrix(1, 0, 0, -1, 0, h)  # flip y
        page.transform(mat)
        self._render_page()

    def _delete_page(self) -> None:
        """Delete the current page after user confirmation; re-index remaining data."""
        if not self._doc:
            return
        n = len(self._doc)
        if n == 1:
            messagebox.showwarning("Cannot delete", "The document has only one page.")
            return
        if not messagebox.askyesno("Delete page",
                                   f"Delete page {self._current_page + 1} of {n}?"):
            return
        self._flush_page()
        # Remove placements and erasers for this page; adjust indices for later pages
        idx = self._current_page
        self._placements = [
            dataclasses.replace(pl, page_idx=pl.page_idx - 1)
            if pl.page_idx > idx else pl
            for pl in self._placements if pl.page_idx != idx
        ]
        self._erasers = [
            dataclasses.replace(e, page_idx=e.page_idx - 1)
            if e.page_idx > idx else e
            for e in self._erasers if e.page_idx != idx
        ]
        self._doc.delete_page(idx)
        self._current_page = min(idx, len(self._doc) - 1)
        self._render_page()
        self._refresh_controls()

    def _insert_blank_page(self) -> None:
        """Insert a blank page of the same size after the current page."""
        if not self._doc:
            return
        self._flush_page()
        page = self._doc[self._current_page]
        r = page.rect
        # Shift placements/erasers for pages after the new page
        insert_after = self._current_page
        self._placements = [
            dataclasses.replace(pl, page_idx=pl.page_idx + 1)
            if pl.page_idx > insert_after else pl
            for pl in self._placements
        ]
        self._erasers = [
            dataclasses.replace(e, page_idx=e.page_idx + 1)
            if e.page_idx > insert_after else e
            for e in self._erasers
        ]
        self._doc.insert_page(insert_after + 1, width=r.width, height=r.height)
        self._current_page = insert_after + 1
        self._render_page()
        self._refresh_controls()

    def _duplicate_page(self) -> None:
        """Insert a copy of the current page immediately after it."""
        if not self._doc:
            return
        self._flush_page()
        idx = self._current_page
        # Shift placements/erasers for pages after the copy target
        self._placements = [
            dataclasses.replace(pl, page_idx=pl.page_idx + 1)
            if pl.page_idx > idx else pl
            for pl in self._placements
        ]
        self._erasers = [
            dataclasses.replace(e, page_idx=e.page_idx + 1)
            if e.page_idx > idx else e
            for e in self._erasers
        ]
        self._doc.copy_page(idx, idx + 1)
        self._current_page = idx + 1
        self._render_page()
        self._refresh_controls()

    def _move_page_up(self) -> None:
        """Swap current page with the previous one."""
        if not self._doc or self._current_page == 0:
            return
        self._flush_page()
        idx = self._current_page
        # Swap placements / eraser page indices
        for pl in self._placements:
            if pl.page_idx == idx:        pl.page_idx = idx - 1
            elif pl.page_idx == idx - 1:  pl.page_idx = idx
        for e in self._erasers:
            if e.page_idx == idx:        e.page_idx = idx - 1
            elif e.page_idx == idx - 1:  e.page_idx = idx
        self._doc.move_page(idx, idx - 1)
        self._current_page = idx - 1
        self._render_page()
        self._refresh_controls()

    def _move_page_down(self) -> None:
        """Swap current page with the next one."""
        if not self._doc or self._current_page >= len(self._doc) - 1:
            return
        self._flush_page()
        idx = self._current_page
        for pl in self._placements:
            if pl.page_idx == idx:        pl.page_idx = idx + 1
            elif pl.page_idx == idx + 1:  pl.page_idx = idx
        for e in self._erasers:
            if e.page_idx == idx:        e.page_idx = idx + 1
            elif e.page_idx == idx + 1:  e.page_idx = idx
        self._doc.move_page(idx, idx + 1)
        self._current_page = idx + 1
        self._render_page()
        self._refresh_controls()

    # ------------------------------------------------------------------
    # Navigation & zoom
    # ------------------------------------------------------------------

    def _on_continuous_toggle(self) -> None:
        """Called when the continuous-scroll checkbox is toggled."""
        if self._doc:
            self._flush_page()
            self._render_page()

    def _on_mousewheel(self, event: tk.Event) -> None:
        """Handle vertical scroll; in continuous mode, update _current_page tracker."""
        self._canvas.yview_scroll(int(-event.delta / 120), "units")
        if self._continuous_var.get() and self._doc:
            self._update_current_page_from_scroll()

    def _update_current_page_from_scroll(self) -> None:
        """Set _current_page to whichever page is most visible in the viewport."""
        if not self._page_offsets:
            return
        # Get the visible y-range of the canvas in canvas coordinates
        y_top = self._canvas.canvasy(0)
        y_bot = self._canvas.canvasy(self._canvas.winfo_height())
        best_pg = self._current_page
        best_overlap = -1
        for pg_idx, (ox, oy) in self._page_offsets.items():
            page = self._doc[pg_idx]
            pg_top = oy
            pg_bot = oy + page.rect.height * self._zoom
            overlap = max(0.0, min(pg_bot, y_bot) - max(pg_top, y_top))
            if overlap > best_overlap:
                best_overlap = overlap
                best_pg = pg_idx
        if best_pg != self._current_page:
            self._current_page = best_pg
            self._refresh_controls()

    def _prev_page(self) -> None:
        """Navigate to the previous page (scrolls in continuous mode, re-renders otherwise)."""
        if self._doc and self._current_page > 0:
            self._flush_page()
            self._current_page -= 1
            if self._continuous_var.get():
                # Just scroll to the page — don't re-render
                self._scroll_to_current_page()
            else:
                self._render_page()
            self._refresh_controls()

    def _next_page(self) -> None:
        """Navigate to the next page (scrolls in continuous mode, re-renders otherwise)."""
        if self._doc and self._current_page < len(self._doc) - 1:
            self._flush_page()
            self._current_page += 1
            if self._continuous_var.get():
                self._scroll_to_current_page()
            else:
                self._render_page()
            self._refresh_controls()

    def _scroll_to_current_page(self) -> None:
        """Scroll the canvas so the current page is visible (continuous mode)."""
        if self._current_page not in self._page_offsets:
            return
        _, cy = self._page_offsets[self._current_page]
        bbox = self._canvas.bbox("all")
        if bbox:
            total_height = bbox[3]
            if total_height > 0:
                self._canvas.yview_moveto(max(0.0, cy / total_height - 0.02))

    def _on_zoom_change(self, _event=None) -> None:
        """Apply the newly selected zoom level and re-render the canvas."""
        try:
            self._zoom = float(self._zoom_var.get().rstrip("%")) / 100.0
        except ValueError:
            return
        if self._doc:
            self._flush_page()
            self._render_page()

    def _show_about(self) -> None:
        """Open the About dialog."""
        dlg = _AboutDialog(self)
        self.wait_window(dlg)

    def _refresh_controls(self) -> None:
        """Enable or disable navigation/page-op buttons to match the document state."""
        has = self._doc is not None
        n = len(self._doc) if has else 0
        p = self._current_page
        self._btn_prev.config(state=tk.NORMAL if has and p > 0 else tk.DISABLED)
        self._btn_next.config(state=tk.NORMAL if has and p < n - 1 else tk.DISABLED)
        self._lbl_page.config(text=f"Page {p + 1} / {n}" if has else "—")
        pg_state = tk.NORMAL if has else tk.DISABLED
        for btn in (self._btn_rot_ccw, self._btn_rot_cw,
                    self._btn_mir_h, self._btn_mir_v,
                    self._btn_del_pg, self._btn_ins_pg, self._btn_dup_pg):
            btn.config(state=pg_state)
        self._btn_pg_up.config(state=tk.NORMAL if has and p > 0 else tk.DISABLED)
        self._btn_pg_dn.config(state=tk.NORMAL if has and p < n - 1 else tk.DISABLED)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    """Create and run the Helper: PDF Editor application."""
    app = PDFFormFiller()
    # Support drag-and-drop onto the .exe: Windows passes the dropped file as
    # sys.argv[1].  Use after(0) so the window is fully realised first.
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        app.after(0, lambda: app._load_pdf(sys.argv[1]))
    app.mainloop()


if __name__ == "__main__":
    main()
