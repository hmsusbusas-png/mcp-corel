import pythoncom


class ComBridge:
    def __init__(self):
        try:
            import win32com.client
        except ImportError as e:
            raise RuntimeError("pywin32 not installed") from e
        try:
            self.app = win32com.client.Dispatch("CorelDRAW.Application")
            self.app.Visible = True
        except Exception as e:
            raise RuntimeError(
                f"CorelDRAW COM unavailable: {e}. Check that CorelDRAW is installed "
                "and registered (try opening it manually once)."
            ) from e

    def active_doc(self):
        try:
            return self.app.ActiveDocument
        except Exception:
            raise RuntimeError("no open document, call corel_open first")

    def draw_shape(self, kind, x, y, w, h):
        layer = self.active_doc().ActiveLayer
        if kind == "rect":
            return layer.CreateRectangle(x, y, x + w, y + h)
        if kind == "ellipse":
            return layer.CreateEllipse(x, y, x + w, y + h)
        raise ValueError(f"unknown shape kind: {kind}")

    def draw_text(self, text, x, y, size):
        shape = self.active_doc().ActiveLayer.CreateArtisticText(x, y, text)
        shape.Text.Story.Size = size
        return shape

    def export(self, path, fmt):
        doc = self.active_doc()
        # cdrPNG=775, cdrPDF=787 per CorelDRAW VBA object model
        fmt_id = 775 if fmt == "png" else 787
        doc.Export(path, fmt_id)
        return path


bridge = None


def get_bridge():
    global bridge
    if bridge is None:
        bridge = ComBridge()
    return bridge
