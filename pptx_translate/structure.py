"""Independent shape and run-style comparison for delivery QA."""

from __future__ import annotations

from lxml import etree
from pptx import Presentation


def _runs(frame) -> tuple[bytes, ...]:
    return tuple(
        etree.tostring(run._r.rPr, method="c14n")
        if run._r.rPr is not None and (len(run._r.rPr) or run._r.rPr.attrib) else b""
        for paragraph in frame.paragraphs for run in paragraph.runs
    )


def _records(shapes, slide_no: int, prefix: str = "") -> dict[str, tuple]:
    found: dict[str, tuple] = {}
    for shape in shapes:
        path = f"{slide_no}/{prefix}{shape.shape_id}"
        geometry = (str(shape.shape_type), int(shape.left), int(shape.top),
                    int(shape.width), int(shape.height))
        style = _runs(shape.text_frame) if shape.has_text_frame else ()
        if getattr(shape, "has_table", False):
            style += tuple(item for row in shape.table.rows for cell in row.cells
                           for item in _runs(cell.text_frame))
        found[path] = (geometry, style)
        if getattr(shape, "has_group", False):
            found.update(_records(shape.shapes, slide_no, prefix + f"{shape.shape_id}/"))
    return found


def compare(source, output) -> dict:
    """Compare geometry and explicit run properties, ignoring text values."""
    a, b = Presentation(str(source)), Presentation(str(output))
    left: dict[str, tuple] = {}
    right: dict[str, tuple] = {}
    for slide_no, slide in enumerate(a.slides, 1):
        left.update(_records(slide.shapes, slide_no))
    for slide_no, slide in enumerate(b.slides, 1):
        right.update(_records(slide.shapes, slide_no))
    shared = left.keys() & right.keys()
    geometry = [key for key in shared if left[key][0] != right[key][0]]
    style = [key for key in shared if left[key][1] != right[key][1]]
    return {
        "shape_set_differences": len(left.keys() ^ right.keys()),
        "geometry_differences": len(geometry),
        "run_style_differences": len(style),
        "examples": (sorted(left.keys() ^ right.keys())[:5] +
                     sorted(geometry)[:5] + sorted(style)[:5]),
    }
