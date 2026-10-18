import pythoncom
import winreg

COM_TIMEOUT = 15


def com_registered():
    for hive in (winreg.HKEY_CLASSES_ROOT,):
        try:
            with winreg.OpenKey(hive, r"CorelDRAW.Application"):
                return True
        except FileNotFoundError:
            pass
    return False


def run_com(fn, *args):
    # COM calls can hang on a slow or broken CorelDRAW install
    import threading
    outcome = {}

    def worker():
        try:
            pythoncom.CoInitialize()
            outcome["value"] = fn(*args)
        except Exception as e:
            outcome["error"] = e
        finally:
            pythoncom.CoUninitialize()

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    t.join(COM_TIMEOUT)
    if t.is_alive():
        raise RuntimeError(
            f"CorelDRAW COM call timed out after {COM_TIMEOUT}s. "
            "CorelDRAW is probably busy with a dialog or broken. Close it and retry."
        )
    if "error" in outcome:
        raise outcome["error"]
    return outcome.get("value")


class ComBridge:
    def __init__(self):
        if not com_registered():
            raise RuntimeError(
                "CorelDRAW.Application is not registered in the registry. "
                "CorelDRAW is probably not installed. Launch it manually once to register COM."
            )
        try:
            import win32com.client
            self.app = win32com.client.Dispatch("CorelDRAW.Application")
            self.app.Visible = True
        except Exception as e:
            raise RuntimeError(f"CorelDRAW COM dispatch failed: {e}") from e

    def active_doc(self):
        try:
            return self.app.ActiveDocument
        except Exception:
            raise RuntimeError("no open document, call corel_open first")

    def draw_shape(self, kind, x, y, w, h):
        def call():
            layer = self.active_doc().ActiveLayer
            if kind == "rect":
                return layer.CreateRectangle(x, y, x + w, y + h)
            if kind == "ellipse":
                return layer.CreateEllipse(x, y, x + w, y + h)
        if kind not in ("rect", "ellipse"):
            raise ValueError(f"unknown shape kind: {kind}")
        return run_com(call)

    def draw_text(self, text, x, y, size):
        def call():
            shape = self.active_doc().ActiveLayer.CreateArtisticText(x, y, text)
            shape.Text.Story.Size = size
            return shape
        return run_com(call)

    def export(self, path, fmt):
        def call():
            doc = self.active_doc()
            fmt_id = 775 if fmt == "png" else 787
            doc.Export(path, fmt_id)
        run_com(call)
        from pathlib import Path
        if not Path(path).exists():
            raise RuntimeError(f"export produced no file at {path}")
        return path


bridge = None


def get_bridge():
    global bridge
    if bridge is None:
        bridge = run_com(lambda: ComBridge())
    return bridge
