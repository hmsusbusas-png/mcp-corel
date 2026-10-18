import queue
import threading
import winreg
from pathlib import Path

import pythoncom

prog_id = "CorelDRAW.Application"
start_timeout = 120
call_timeout = 60


def installed_exe():
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id + r"\CLSID") as key:
            clsid = winreg.QueryValue(key, None)
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32") as key:
            command = winreg.QueryValue(key, None)
    except OSError:
        return None
    exe = command.split('"')[1] if command.startswith('"') else command.split(" /")[0]
    return exe if Path(exe).is_file() else None


class ComWorker:
    # every COM object lives and dies on this one STA thread
    def __init__(self):
        self.jobs = queue.Queue()
        self.app = None
        self.thread = threading.Thread(target=self.loop, daemon=True, name="corel-com")
        self.thread.start()

    def loop(self):
        pythoncom.CoInitialize()
        try:
            while True:
                job, box, done = self.jobs.get()
                try:
                    box["value"] = job(self)
                except BaseException as exc:
                    box["error"] = exc
                finally:
                    done.set()
        finally:
            pythoncom.CoUninitialize()

    def connect(self):
        if self.app is None:
            from win32com.client import gencache
            self.app = gencache.EnsureDispatch(prog_id)
            self.app.Visible = True
        return self.app


worker = None
worker_lock = threading.Lock()


def run(job, timeout=call_timeout):
    global worker
    with worker_lock:
        if worker is None:
            if installed_exe() is None:
                raise RuntimeError("CorelDRAW is not installed or its COM server is not registered")
            worker = ComWorker()
            timeout = max(timeout, start_timeout)
        current = worker
    box, done = {}, threading.Event()
    current.jobs.put((job, box, done))
    if not done.wait(timeout):
        with worker_lock:
            if worker is current:
                worker = None
        raise TimeoutError(
            f"CorelDRAW did not answer in {timeout}s, it may be showing a dialog. "
            "Close the dialog or restart CorelDRAW, the next call reconnects"
        )
    if "error" in box:
        raise box["error"]
    return box.get("value")


def constants():
    from win32com.client import constants as k
    return k
