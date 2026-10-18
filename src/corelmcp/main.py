import logging
import sys
import time
from functools import wraps
from pathlib import Path

from mcp.server.fastmcp import FastMCP

from . import comdesk
from .comdesk import constants, installed_exe

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
log = logging.getLogger("corel")

app = FastMCP("corel")

export_kinds = {"png": "cdrPNG", "jpg": "cdrJPEG", "svg": "cdrSVG", "pdf": None}


def timed(fn):
    @wraps(fn)
    def wrap(*args, **kwargs):
        start = time.monotonic()
        try:
            return fn(*args, **kwargs)
        finally:
            log.info("%s %.3fs", fn.__name__, time.monotonic() - start)
    return wrap


def active_doc(w):
    k = constants()
    corel = w.connect()
    if corel.Documents.Count == 0:
        raise RuntimeError("no open document, call corel_open first")
    doc = corel.ActiveDocument
    doc.Unit = k.cdrMillimeter
    return doc


def rgb(hex_color):
    text = hex_color.lstrip("#")
    if len(text) != 6:
        raise ValueError(f"color must look like #RRGGBB, got {hex_color!r}")
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


@app.tool(description="CorelDRAW version and open documents, starts CorelDRAW if needed")
@timed
def corel_status() -> dict:
    exe = installed_exe()
    if exe is None:
        return {"installed": False}

    def job(w):
        corel = w.connect()
        return {
            "installed": True,
            "exe": exe,
            "version": f"{corel.VersionMajor}.{corel.VersionMinor}",
            "documents": [corel.Documents.Item(i + 1).Title for i in range(corel.Documents.Count)],
        }
    return comdesk.run(job)


@app.tool(description="Open a .cdr file, or create a new A4 document when path is empty. Units are millimeters")
@timed
def corel_open(path: str = "") -> dict:
    if path and not Path(path).is_file():
        raise FileNotFoundError(f"no such file: {path}")

    def job(w):
        k = constants()
        corel = w.connect()
        doc = corel.OpenDocument(str(Path(path).resolve())) if path else corel.CreateDocument()
        doc.Unit = k.cdrMillimeter
        page = doc.ActivePage
        return {"title": doc.Title, "page_mm": [round(page.SizeWidth, 2), round(page.SizeHeight, 2)]}
    return comdesk.run(job)


@app.tool(description="Draw rect or ellipse, x and y of the lower left corner in mm from the page bottom left, fill #RRGGBB")
@timed
def corel_shape(kind: str, x: float, y: float, w: float, h: float, fill: str = "") -> dict:
    if kind not in ("rect", "ellipse"):
        raise ValueError("kind is rect or ellipse")
    if w <= 0 or h <= 0:
        raise ValueError("w and h must be positive")
    color = rgb(fill) if fill else None

    def job(wk):
        layer = active_doc(wk).ActiveLayer
        maker = layer.CreateRectangle if kind == "rect" else layer.CreateEllipse
        shape = maker(x, y + h, x + w, y)
        if color:
            shape.Fill.UniformColor.RGBAssign(*color)
        return {"id": shape.StaticID, "kind": kind}
    return comdesk.run(job)


@app.tool(description="Add artistic text at x, y in mm, font size in points, optional font name and #RRGGBB color")
@timed
def corel_text(text: str, x: float, y: float, size: float = 24, font: str = "", color: str = "") -> dict:
    if not text:
        raise ValueError("text is empty")
    paint = rgb(color) if color else None

    def job(w):
        shape = active_doc(w).ActiveLayer.CreateArtisticText(x, y, text)
        story = shape.Text.Story
        story.Size = size
        if font:
            story.Font = font
        if paint:
            shape.Fill.UniformColor.RGBAssign(*paint)
        return {"id": shape.StaticID}
    return comdesk.run(job)


@app.tool(description="Export the active page to png, jpg, svg or pdf, dpi applies to png and jpg")
@timed
def corel_export(path: str, fmt: str = "png", dpi: int = 150) -> dict:
    fmt = fmt.lower()
    if fmt not in export_kinds:
        raise ValueError(f"fmt is one of {', '.join(export_kinds)}")
    target = Path(path).resolve().with_suffix("." + fmt)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        target.unlink()

    def job(w):
        k = constants()
        doc = active_doc(w)
        if fmt == "pdf":
            doc.PublishToPDF(str(target))
        elif fmt in ("png", "jpg"):
            filt = doc.ExportBitmap(str(target), getattr(k, export_kinds[fmt]), k.cdrCurrentPage, k.cdrRGBColorImage,
                                    0, 0, dpi, dpi, PaletteOptions=None, ExportArea=None)
            filt.Finish()
        else:
            doc.Export(str(target), getattr(k, export_kinds[fmt]), k.cdrCurrentPage, None, None)
    comdesk.run(job, timeout=180)
    if not target.is_file() or target.stat().st_size == 0:
        raise RuntimeError(f"CorelDRAW reported success but {target} was not written")
    return {"path": str(target), "bytes": target.stat().st_size}


@app.tool(description="Save the active document as .cdr")
@timed
def corel_save(path: str) -> dict:
    target = Path(path).resolve().with_suffix(".cdr")
    target.parent.mkdir(parents=True, exist_ok=True)
    comdesk.run(lambda w: active_doc(w).SaveAs(str(target), None), timeout=120)
    if not target.is_file():
        raise RuntimeError(f"{target} was not written")
    return {"path": str(target), "bytes": target.stat().st_size}


def run():
    app.run(transport="stdio")
