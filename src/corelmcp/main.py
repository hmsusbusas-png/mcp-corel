import logging
import sys
import time
from functools import wraps

from mcp.server.fastmcp import FastMCP

from .comdesk import get_bridge

logging.basicConfig(stream=sys.stderr, level=logging.INFO)
log = logging.getLogger("corel")

app = FastMCP("corel")


def quiet(fn):
    @wraps(fn)
    def wrap(*args, **kwargs):
        start = time.monotonic()
        try:
            return fn(*args, **kwargs)
        finally:
            log.info("%s %.3fs", fn.__name__, time.monotonic() - start)
    return wrap


@app.tool(description="Open a CDR file or create a new document")
@quiet
def corel_open(path: str = "") -> str:
    b = get_bridge()
    if path:
        doc = b.app.OpenDocument(path)
    else:
        doc = b.app.CreateDocument()
    return f"document open: {doc.Title}"


@app.tool(description="Draw a rectangle or ellipse at x,y with size w,h")
@quiet
def corel_shape(kind: str, x: float, y: float, w: float, h: float) -> str:
    shape = get_bridge().draw_shape(kind, x, y, w, h)
    return f"{kind} created, id {shape.StaticID}"


@app.tool(description="Add artistic text at x,y with font size")
@quiet
def corel_text(text: str, x: float, y: float, size: float = 24) -> str:
    shape = get_bridge().draw_text(text, x, y, size)
    return f"text placed, id {shape.StaticID}"


@app.tool(description="Export the active document to PNG or PDF")
@quiet
def corel_export(path: str, fmt: str = "png") -> str:
    if fmt not in ("png", "pdf"):
        raise ValueError("fmt is png or pdf")
    return get_bridge().export(path, fmt)


def run():
    app.run(transport="stdio")
