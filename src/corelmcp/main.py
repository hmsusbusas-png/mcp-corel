import logging
import sys
import time
from functools import wraps
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from . import comdesk
from .comdesk import constants, installed_exe

logging.basicConfig(stream=sys.stderr, level=logging.INFO, format="%(levelname)s %(name)s %(message)s")
log = logging.getLogger("corel")

app = FastMCP("corel", instructions=(
    "Drives CorelDRAW through COM. All sizes and coordinates are millimeters, x and y are measured from the lower "
    "left corner of the page and mean the lower left corner of a shape. Every drawing tool returns the shape id, "
    "use it with corel_modify, corel_group, corel_align and corel_export. Start with corel_open to get a document."
))

export_kinds = {"png": "cdrPNG", "jpg": "cdrJPEG", "svg": "cdrSVG", "pdf": None, "bmp": "cdrBMP", "tif": "cdrTIFF",
                "eps": "cdrEPS", "ai": "cdrAI", "dxf": "cdrDXF", "webp": "cdrWEBP"}
raster_kinds = ("png", "jpg", "webp", "bmp", "tif")
shape_types = {1: "rectangle", 2: "ellipse", 3: "curve", 4: "polygon", 5: "bitmap", 6: "text", 7: "group",
               8: "selection", 9: "guideline", 10: "blend", 15: "symbol", 20: "connector", 25: "ole", 26: "contour"}


def tool(description, read_only=False, destructive=False):
    hints = ToolAnnotations(readOnlyHint=read_only, destructiveHint=destructive, idempotentHint=read_only, openWorldHint=False)

    def register(fn):
        @wraps(fn)
        def timed(*args, **kwargs):
            start = time.monotonic()
            try:
                return fn(*args, **kwargs)
            finally:
                log.info("%s %.3fs", fn.__name__, time.monotonic() - start)
        app.tool(description=description, annotations=hints)(timed)
        return fn
    return register


def active_doc(w):
    k = constants()
    corel = w.connect()
    if corel.Documents.Count == 0:
        raise RuntimeError("no open document, call corel_open first")
    doc = corel.ActiveDocument
    doc.Unit = k.cdrMillimeter
    doc.ReferencePoint = k.cdrBottomLeft
    return doc


def rgb(hex_color):
    text = hex_color.lstrip("#")
    if len(text) != 6:
        raise ValueError(f"color must look like #RRGGBB, got {hex_color!r}")
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


def paint(shape, fill, outline, outline_mm):
    if fill == "none":
        shape.Fill.ApplyNoFill()
    elif fill:
        shape.Fill.UniformColor.RGBAssign(*rgb(fill))
    if outline == "none":
        shape.Outline.SetNoOutline()
    elif outline:
        shape.Outline.Color.RGBAssign(*rgb(outline))
    if outline_mm is not None and outline != "none":
        shape.Outline.Width = outline_mm


def find(doc, shape_id):
    pending = [doc.ActivePage.Shapes]
    while pending:
        shapes = pending.pop()
        for i in range(1, shapes.Count + 1):
            shape = shapes.Item(i)
            if shape.StaticID == shape_id:
                return shape
            if shape.Type == 7:
                pending.append(shape.Shapes)
    raise LookupError(f"no shape with id {shape_id} on the active page, see corel_list")


def describe(shape):
    info = {
        "id": shape.StaticID,
        "type": shape_types.get(shape.Type, str(shape.Type)),
        "name": shape.Name,
        "x": round(shape.LeftX, 3), "y": round(shape.BottomY, 3),
        "w": round(shape.SizeWidth, 3), "h": round(shape.SizeHeight, 3),
    }
    if shape.Type == 6:
        info["text"] = shape.Text.Story.Text[:200]
    if shape.Type == 7:
        info["children"] = shape.Shapes.Count
    return info


@tool("CorelDRAW install path, version and open documents with page count. Starts CorelDRAW if needed", read_only=True)
def corel_status() -> dict:
    exe = installed_exe()
    if exe is None:
        return {"installed": False}

    def job(w):
        corel = w.connect()
        docs = []
        for i in range(1, corel.Documents.Count + 1):
            d = corel.Documents.Item(i)
            docs.append({"title": d.Title, "pages": d.Pages.Count, "path": d.FullFileName or None})
        return {"installed": True, "exe": exe, "version": f"{corel.VersionMajor}.{corel.VersionMinor}",
                "active": corel.ActiveDocument.Title if docs else None, "documents": docs}
    return comdesk.run(job)


@tool("Open a .cdr (or any format CorelDRAW opens), or create a new document. width and height set the page size in mm, default A4")
def corel_open(path: str = "", width: float = 0, height: float = 0) -> dict:
    if path and not Path(path).is_file():
        raise FileNotFoundError(f"no such file: {path}")

    def job(w):
        corel = w.connect()
        doc = corel.OpenDocument(str(Path(path).resolve())) if path else corel.CreateDocument()
        doc.Activate()
        doc = active_doc(w)
        if width and height:
            doc.ActivePage.SetSize(width, height)
        page = doc.ActivePage
        return {"title": doc.Title, "pages": doc.Pages.Count, "page_mm": [round(page.SizeWidth, 2), round(page.SizeHeight, 2)]}
    return comdesk.run(job)


@tool("Close the active document. save=true saves it first (it must already have a file name)", destructive=True)
def corel_close(save: bool = False) -> str:
    def job(w):
        doc = active_doc(w)
        title = doc.Title
        if save:
            doc.Save()
        else:
            doc.Dirty = False
        doc.Close()
        return f"closed {title}"
    return comdesk.run(job)


@tool("Pages of the active document: action list, add (count), delete (number), goto (number), size (width, height mm), landscape, portrait", destructive=True)
def corel_pages(action: str = "list", number: int = 0, count: int = 1, width: float = 0, height: float = 0) -> dict:
    def job(w):
        doc = active_doc(w)
        if action == "add":
            doc.AddPages(count)
        elif action in ("delete", "goto"):
            if not 1 <= number <= doc.Pages.Count:
                raise ValueError(f"page number is 1..{doc.Pages.Count}")
            target = doc.Pages.Item(number)
            if action == "delete":
                target.Delete()
            else:
                target.Activate()
        elif action == "size":
            if width <= 0 or height <= 0:
                raise ValueError("width and height must be positive")
            doc.ActivePage.SetSize(width, height)
        elif action in ("landscape", "portrait"):
            page = doc.ActivePage
            short, long = sorted((page.SizeWidth, page.SizeHeight))
            page.SetSize(long, short) if action == "landscape" else page.SetSize(short, long)
        elif action != "list":
            raise ValueError("action is list, add, delete, goto, size, landscape or portrait")
        pages = [{"number": i, "w": round(doc.Pages.Item(i).SizeWidth, 2), "h": round(doc.Pages.Item(i).SizeHeight, 2),
                  "shapes": doc.Pages.Item(i).Shapes.Count} for i in range(1, doc.Pages.Count + 1)]
        return {"active": doc.ActivePage.Index, "pages": pages}
    return comdesk.run(job)


@tool("Draw rect, ellipse, polygon or star at x, y with size w, h in mm. sides for polygon and star, corner_mm rounds rect corners. fill and outline are #RRGGBB or none")
def corel_shape(kind: str, x: float, y: float, w: float, h: float, fill: str = "", outline: str = "",
                outline_mm: float | None = None, sides: int = 5, corner_mm: float = 0, name: str = "") -> dict:
    if kind not in ("rect", "ellipse", "polygon", "star"):
        raise ValueError("kind is rect, ellipse, polygon or star")
    if w <= 0 or h <= 0:
        raise ValueError("w and h must be positive")

    def job(wk):
        layer = active_doc(wk).ActiveLayer
        if kind == "rect":
            shape = layer.CreateRectangle2(x, y, w, h, corner_mm, corner_mm, corner_mm, corner_mm)
        elif kind == "ellipse":
            shape = layer.CreateEllipse(x, y + h, x + w, y)
        else:
            shape = layer.CreatePolygon(x, y + h, x + w, y, max(3, sides), 1, 1, kind == "star", 50, 100)
        paint(shape, fill, outline, outline_mm)
        if name:
            shape.Name = name
        return describe(shape)
    return comdesk.run(job)


@tool("Draw a straight line or a polyline through points [[x, y], ...] in mm, closed=true makes a closed outline")
def corel_line(points: list[list[float]], outline: str = "#000000", outline_mm: float = 0.3, closed: bool = False,
               fill: str = "", name: str = "") -> dict:
    if len(points) < 2:
        raise ValueError("need at least two points")

    def job(w):
        layer = active_doc(w).ActiveLayer
        (x0, y0), (x1, y1) = points[0], points[1]
        shape = layer.CreateLineSegment(x0, y0, x1, y1)
        if len(points) > 2 or closed:
            path = shape.Curve.SubPaths.Item(1)
            for px, py in points[2:]:
                path.AppendLineSegment(px, py)
            if closed:
                path.Closed = True
        paint(shape, fill, outline, outline_mm)
        if name:
            shape.Name = name
        return describe(shape)
    return comdesk.run(job)


@tool("Add text at x, y in mm. With width_mm and height_mm it is a paragraph text frame that wraps, otherwise one line of artistic text. size in points")
def corel_text(text: str, x: float, y: float, size: float = 24, font: str = "", color: str = "", bold: bool = False,
               italic: bool = False, align: str = "left", width_mm: float = 0, height_mm: float = 0, name: str = "") -> dict:
    if not text:
        raise ValueError("text is empty")
    aligns = {"left": 1, "center": 2, "right": 3, "justify": 4}
    if align not in aligns:
        raise ValueError("align is left, center, right or justify")

    def job(w):
        layer = active_doc(w).ActiveLayer
        if width_mm and height_mm:
            shape = layer.CreateParagraphText(x, y + height_mm, x + width_mm, y, text)
        else:
            shape = layer.CreateArtisticText(x, y, text)
        story = shape.Text.Story
        story.Size = size
        if font:
            story.Font = font
        story.Bold = bold
        story.Italic = italic
        story.Alignment = aligns[align]
        if color:
            shape.Fill.UniformColor.RGBAssign(*rgb(color))
        if name:
            shape.Name = name
        return describe(shape)
    return comdesk.run(job)


@tool("Import an image or vector file (png, jpg, svg, pdf, ai, eps, dxf...) onto the active page, optionally placed at x, y and scaled to width_mm")
def corel_import(path: str, x: float | None = None, y: float | None = None, width_mm: float = 0) -> dict:
    source = Path(path).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"no such file: {path}")

    def job(w):
        doc = active_doc(w)
        before = doc.ActivePage.Shapes.Count
        doc.ActiveLayer.Import(str(source), 0, None)
        if doc.ActivePage.Shapes.Count == before:
            raise RuntimeError("CorelDRAW imported nothing")
        picked = doc.SelectionRange
        shape = picked.Item(1) if picked.Count else doc.ActivePage.Shapes.Item(1)
        if width_mm:
            shape.SetSize(width_mm, shape.SizeHeight * width_mm / shape.SizeWidth)
        if x is not None and y is not None:
            shape.SetPositionEx(constants().cdrBottomLeft, x, y)
        return describe(shape)
    return comdesk.run(job, timeout=180)


@tool("List shapes on the active page with id, type, name, position, size and text", read_only=True)
def corel_list(limit: int = 200) -> dict:
    def job(w):
        page = active_doc(w).ActivePage
        shapes = [describe(page.Shapes.Item(i)) for i in range(1, min(page.Shapes.Count, limit) + 1)]
        return {"page": page.Index, "page_mm": [round(page.SizeWidth, 2), round(page.SizeHeight, 2)],
                "count": page.Shapes.Count, "shapes": shapes}
    return comdesk.run(job)


@tool("Change a shape by id: x, y position, w, h size, rotate degrees, fill, outline, outline_mm, text, size, font, name, order front or back, duplicate, delete", destructive=True)
def corel_modify(shape_id: int, x: float | None = None, y: float | None = None, w: float | None = None, h: float | None = None,
                 rotate: float = 0, fill: str = "", outline: str = "", outline_mm: float | None = None, text: str | None = None,
                 size: float | None = None, font: str = "", name: str = "", order: str = "", duplicate: bool = False,
                 delete: bool = False) -> dict:
    if order not in ("", "front", "back"):
        raise ValueError("order is front or back")

    def job(wk):
        k = constants()
        doc = active_doc(wk)
        shape = find(doc, shape_id)
        if delete:
            shape.Delete()
            return {"deleted": shape_id}
        if duplicate:
            shape = shape.Duplicate(0, 0)
        if w is not None or h is not None:
            shape.SetSize(w if w is not None else shape.SizeWidth, h if h is not None else shape.SizeHeight)
        if x is not None or y is not None:
            shape.SetPositionEx(k.cdrBottomLeft, x if x is not None else shape.LeftX, y if y is not None else shape.BottomY)
        if rotate:
            shape.Rotate(rotate)
        paint(shape, fill, outline, outline_mm)
        if shape.Type == 6:
            story = shape.Text.Story
            if text is not None:
                story.Text = text
            if size:
                story.Size = size
            if font:
                story.Font = font
        if name:
            shape.Name = name
        if order == "front":
            shape.OrderToFront()
        elif order == "back":
            shape.OrderToBack()
        return describe(shape)
    return comdesk.run(job)


@tool("Group shapes by ids, or ungroup a group id")
def corel_group(shape_ids: list[int], ungroup: bool = False) -> dict:
    if not shape_ids:
        raise ValueError("shape_ids is empty")

    def job(w):
        doc = active_doc(w)
        if ungroup:
            group = find(doc, shape_ids[0])
            if group.Type != 7:
                raise ValueError(f"shape {shape_ids[0]} is not a group")
            parts = group.UngroupEx()
            return {"ungrouped": shape_ids[0], "shapes": [describe(parts.Item(i)) for i in range(1, parts.Count + 1)]}
        if len(shape_ids) < 2:
            raise ValueError("need two or more ids to group")
        bundle = doc.CreateShapeRangeFromArray(*[find(doc, i) for i in shape_ids]).Group()
        return describe(bundle)
    return comdesk.run(job)


@tool("Align shapes to the page: horizontal left, center, right and vertical top, middle, bottom")
def corel_align(shape_ids: list[int], horizontal: str = "", vertical: str = "") -> list[dict]:
    if horizontal and horizontal not in ("left", "center", "right"):
        raise ValueError("horizontal is left, center or right")
    if vertical and vertical not in ("top", "middle", "bottom"):
        raise ValueError("vertical is top, middle or bottom")

    def job(w):
        k = constants()
        doc = active_doc(w)
        page_w, page_h = doc.ActivePage.SizeWidth, doc.ActivePage.SizeHeight
        out = []
        for i in shape_ids:
            shape = find(doc, i)
            free_x, free_y = page_w - shape.SizeWidth, page_h - shape.SizeHeight
            x = {"left": 0, "center": free_x / 2, "right": free_x}.get(horizontal, shape.LeftX)
            y = {"bottom": 0, "middle": free_y / 2, "top": free_y}.get(vertical, shape.BottomY)
            shape.SetPositionEx(k.cdrBottomLeft, x, y)
            out.append(describe(shape))
        return out
    return comdesk.run(job)


@tool("Export the active page, or only the shapes in shape_ids, to png, jpg, webp, bmp, tif, svg, pdf, eps, ai, dxf. dpi applies to raster formats, transparent to png")
def corel_export(path: str, fmt: str = "png", dpi: int = 150, shape_ids: list[int] | None = None,
                 transparent: bool = False) -> dict:
    fmt = fmt.lower()
    if fmt not in export_kinds:
        raise ValueError(f"fmt is one of {', '.join(export_kinds)}")
    target = Path(path).resolve().with_suffix("." + fmt)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.unlink(missing_ok=True)

    def job(w):
        k = constants()
        doc = active_doc(w)
        scope = k.cdrCurrentPage
        page = doc.ActivePage
        area = None if shape_ids else w.connect().CreateRect(0, 0, page.SizeWidth, page.SizeHeight)
        if shape_ids:
            doc.ClearSelection()
            for i in shape_ids:
                find(doc, i).AddToSelection()
            scope = k.cdrSelection
        try:
            if fmt == "pdf":
                doc.PDFSettings.PublishRange = k.pdfSelection if shape_ids else k.pdfCurrentPage
                doc.PublishToPDF(str(target))
            elif fmt in raster_kinds:
                filt = doc.ExportBitmap(str(target), getattr(k, export_kinds[fmt]), scope, k.cdrRGBColorImage, 0, 0,
                                        dpi, dpi, k.cdrNormalAntiAliasing, False, transparent and fmt == "png",
                                        True, False, k.cdrCompressionNone, None, area)
                filt.Finish()
            else:
                doc.Export(str(target), getattr(k, export_kinds[fmt]), scope, None, None)
        finally:
            if shape_ids:
                doc.ClearSelection()
    comdesk.run(job, timeout=180)
    if not target.is_file() or target.stat().st_size == 0:
        raise RuntimeError(f"CorelDRAW reported success but {target} was not written")
    return {"path": str(target), "bytes": target.stat().st_size}


@tool("Save the active document as .cdr")
def corel_save(path: str) -> dict:
    target = Path(path).resolve().with_suffix(".cdr")
    target.parent.mkdir(parents=True, exist_ok=True)
    comdesk.run(lambda w: active_doc(w).SaveAs(str(target), None), timeout=120)
    if not target.is_file():
        raise RuntimeError(f"{target} was not written")
    return {"path": str(target), "bytes": target.stat().st_size}


@tool("Run an installed CorelDRAW VBA or GMS macro: project, module.macro and optional string arguments", destructive=True)
def corel_macro(project: str, macro: str, args: list[str] | None = None) -> str:
    def job(w):
        try:
            outcome = w.connect().GMSManager.RunMacro(project, macro, *(args or []))
        except Exception as exc:
            raise RuntimeError(f"macro {project}.{macro} failed, check the project and macro names in the Macro Manager: {exc}") from exc
        return "" if outcome is None else str(outcome)
    return comdesk.run(job, timeout=180)


def run():
    app.run(transport="stdio")
