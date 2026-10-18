import asyncio

import pytest

from corelmcp import comdesk, main
from corelmcp.main import app


def test_every_tool_is_registered():
    tools = {t.name: t for t in asyncio.run(app.list_tools())}
    assert len(tools) == 15
    assert tools["corel_list"].annotations.readOnlyHint and tools["corel_modify"].annotations.destructiveHint


def test_color_parsing():
    assert main.rgb("#1E64C8") == (30, 100, 200)
    assert main.rgb("ff0000") == (255, 0, 0)
    with pytest.raises(ValueError):
        main.rgb("red")


def test_argument_checks_need_no_corel():
    with pytest.raises(ValueError):
        main.corel_shape("hexagon", 0, 0, 1, 1)
    with pytest.raises(ValueError):
        main.corel_shape("rect", 0, 0, -1, 1)
    with pytest.raises(ValueError):
        main.corel_line([[0, 0]])
    with pytest.raises(ValueError):
        main.corel_export("x.gif", "gif")
    with pytest.raises(ValueError):
        main.corel_align([1], horizontal="middle")
    with pytest.raises(FileNotFoundError):
        main.corel_open("C:/no/such/file.cdr")


def test_install_lookup_returns_a_real_file_or_nothing():
    exe = comdesk.installed_exe()
    assert exe is None or exe.lower().endswith(".exe")


live = pytest.mark.skipif(comdesk.installed_exe() is None, reason="CorelDRAW is not installed")


@pytest.mark.live
@live
def test_drawing_session(tmp_path):
    assert main.corel_status()["installed"]
    assert main.corel_open(width=200, height=120)["page_mm"] == [200.0, 120.0]
    try:
        box = main.corel_shape("rect", 10, 10, 60, 30, fill="#2060C0", corner_mm=4, name="box")
        assert (box["x"], box["y"], box["w"], box["h"]) == (10.0, 10.0, 60.0, 30.0)
        star = main.corel_shape("star", 120, 10, 30, 30, fill="#F0C000")
        hexagon = main.corel_shape("polygon", 160, 10, 30, 30, sides=6)
        main.corel_line([[10, 60], [60, 90], [110, 60]], outline="#800080", outline_mm=1)
        title = main.corel_text("Проверка", 10, 100, size=20, bold=True)
        main.corel_text("Абзац с переносом строк внутри рамки", 120, 40, size=9, width_mm=60, height_mm=15)
        moved = main.corel_modify(box["id"], x=15, y=12, w=55, fill="#1040A0")
        assert (moved["x"], moved["y"], moved["w"]) == (15.0, 12.0, 55.0)
        assert main.corel_modify(title["id"], text="Проверка 2")["text"] == "Проверка 2"
        centered = main.corel_align([title["id"]], horizontal="center")[0]
        assert centered["x"] == pytest.approx((200 - centered["w"]) / 2, abs=0.01)
        group = main.corel_group([star["id"], hexagon["id"]])
        assert group["type"] == "group" and group["children"] == 2
        assert len(main.corel_group([group["id"]], ungroup=True)["shapes"]) == 2
        listed = main.corel_list()
        assert listed["count"] == 6 and {s["name"] for s in listed["shapes"]} >= {"box"}
        pages = main.corel_pages("add", count=1)
        assert len(pages["pages"]) == 2
        main.corel_pages("goto", number=1)
        for fmt in ("png", "jpg", "svg", "pdf", "webp"):
            assert main.corel_export(str(tmp_path / f"page.{fmt}"), fmt, dpi=72)["bytes"] > 500
        assert main.corel_export(str(tmp_path / "title.png"), "png", shape_ids=[title["id"]], transparent=True)["bytes"] > 100
        saved = main.corel_save(str(tmp_path / "test.cdr"))
        assert saved["bytes"] > 1000
        with pytest.raises(LookupError):
            main.corel_modify(999999, x=1)
    finally:
        main.corel_close()
