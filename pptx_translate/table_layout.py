"""Preserve the rendered height of PowerPoint tables with automatic rows.

OOXML permits zero row heights.  PowerPoint then sizes those rows from their
text.  Translation changes that calculation, which can move borders relative
to independent text boxes laid over the table.  On Windows, an installed
PowerPoint can resolve the source row heights before text is changed.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from pptx.util import Emu

log = logging.getLogger("pptx_translate.table_layout")


def _walk_tables(prs):
    def walk(shapes, slide_no):
        for shape in shapes:
            if getattr(shape, "has_group", False):
                yield from walk(shape.shapes, slide_no)
            elif getattr(shape, "has_table", False):
                yield slide_no, shape

    for slide_no, slide in enumerate(prs.slides, 1):
        yield from walk(slide.shapes, slide_no)


def dynamic_tables(prs) -> list[dict]:
    """List tables whose rendered row layout may depend on translated text."""
    return [
        {"slide": slide_no, "shape_id": shape.shape_id, "name": shape.name,
         "auto_rows": [i + 1 for i, row in enumerate(shape.table.rows) if not row.height]}
        for slide_no, shape in _walk_tables(prs)
        if any(not row.height for row in shape.table.rows)
    ]


def capture_powerpoint_rows(path: Path, wanted: list[dict]) -> dict[tuple[int, int], list[float]]:
    """Read actual row heights in points from a separate PowerPoint instance."""
    if not wanted or sys.platform != "win32":
        return {}
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]

    wanted_keys = {(x["slide"], x["shape_id"]) for x in wanted}
    result: dict[tuple[int, int], list[float]] = {}
    pythoncom.CoInitialize()
    app = deck = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        deck = app.Presentations.Open(str(path.resolve()), True, False, False)

        def walk(shapes, slide_no):
            for shape in shapes:
                key = (slide_no, int(shape.Id))
                try:
                    if key in wanted_keys and shape.HasTable:
                        rows = shape.Table.Rows
                        result[key] = [float(rows.Item(i).Height)
                                       for i in range(1, rows.Count + 1)]
                    if int(shape.Type) == 6:  # msoGroup
                        walk(shape.GroupItems, slide_no)
                except Exception as exc:  # one malformed shape must not hide others
                    log.warning("could not measure table %s: %s", key, exc)

        for slide in deck.Slides:
            walk(slide.Shapes, int(slide.SlideIndex))
    finally:
        if deck is not None:
            deck.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()
    return result


def apply_row_heights(prs, captured: dict[tuple[int, int], list[float]]) -> int:
    """Fix automatic rows to the source's visual heights, keeping table bounds."""
    changed = 0
    for slide_no, shape in _walk_tables(prs):
        key = (slide_no, shape.shape_id)
        heights_pt = captured.get(key)
        if not heights_pt or len(heights_pt) != len(shape.table.rows):
            continue
        if not any(not row.height for row in shape.table.rows):
            continue
        measured = [round(pt * 12700) for pt in heights_pt]
        if not all(x > 0 for x in measured):
            continue
        total = sum(measured)
        if total <= 0:
            continue
        # COM reports rendered points, while the OOXML table frame is in EMU.
        # Normalize rounding so the fixed rows exactly fill the original frame.
        fixed = [round(int(shape.height) * x / total) for x in measured]
        fixed[-1] = int(shape.height) - sum(fixed[:-1])
        if fixed[-1] <= 0:
            continue
        for row, height in zip(shape.table.rows, fixed):
            row.height = Emu(height)
        changed += 1
    return changed
