#!/usr/bin/env python3
"""Structural QA helper for dissertation defense PPTX files.

This script intentionally uses only the Python standard library. It checks
package structure, widescreen dimensions, slide count, approximate text/picture
object counts, possible full-slide screenshot slides, and pure-white background
risks.
"""

from __future__ import annotations

import argparse
import json
import math
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


PML = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EMU_PER_POINT = 12700
FORBIDDEN_SLIDE_TERMS = [
    "资料来源",
    "素材来源",
    "答辩提纲",
    "生成说明",
    "AI生成",
    "来自论文原文",
    "截图自",
    "output/ppt_plan.md",
]


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None
    except ET.ParseError as exc:
        raise SystemExit(f"Invalid XML in {name}: {exc}") from exc


def get_slide_size(prs: ET.Element | None) -> tuple[int | None, int | None]:
    if prs is None:
        return None, None
    size = prs.find(f"{PML}sldSz")
    if size is None:
        return None, None
    return int(size.attrib.get("cx", "0")), int(size.attrib.get("cy", "0"))


def shape_bounds(sp_tree: ET.Element) -> list[tuple[int, int, int, int, str]]:
    bounds: list[tuple[int, int, int, int, str]] = []
    for elem in sp_tree.iter():
        if not elem.tag.endswith("}sp") and not elem.tag.endswith("}pic") and not elem.tag.endswith("}graphicFrame"):
            continue
        tag = elem.tag.rsplit("}", 1)[-1]
        xfrm = elem.find(f".//{A}xfrm")
        if xfrm is None:
            continue
        off = xfrm.find(f"{A}off")
        ext = xfrm.find(f"{A}ext")
        if off is None or ext is None:
            continue
        try:
            x = int(off.attrib.get("x", "0"))
            y = int(off.attrib.get("y", "0"))
            cx = int(ext.attrib.get("cx", "0"))
            cy = int(ext.attrib.get("cy", "0"))
        except ValueError:
            continue
        bounds.append((x, y, cx, cy, tag))
    return bounds


def solid_rgb(elem: ET.Element) -> str | None:
    srgb = elem.find(f".//{A}solidFill/{A}srgbClr")
    if srgb is not None:
        value = srgb.attrib.get("val")
        return value.upper() if value else None
    return None


def slide_background_rgb(root: ET.Element) -> str | None:
    bg = root.find(f".//{PML}cSld/{PML}bg")
    if bg is None:
        return None
    return solid_rgb(bg)


def full_slide_nonwhite_shapes(root: ET.Element, slide_w: int | None, slide_h: int | None) -> list[str]:
    if not slide_w or not slide_h:
        return []
    slide_area = slide_w * slide_h
    found: list[str] = []
    for elem in root.iter():
        if local_name(elem.tag) not in {"sp", "pic"}:
            continue
        ext = elem.find(f".//{A}xfrm/{A}ext")
        if ext is None:
            continue
        try:
            cx = int(ext.attrib.get("cx", "0"))
            cy = int(ext.attrib.get("cy", "0"))
        except ValueError:
            continue
        if not slide_area or (cx * cy / slide_area) < 0.92:
            continue
        if local_name(elem.tag) == "pic":
            found.append("full-slide picture")
            continue
        rgb = solid_rgb(elem)
        if rgb and rgb != "FFFFFF":
            found.append(f"full-slide non-white shape #{rgb}")
    return found


def text_width_units(text: str, font_pt: float) -> float:
    units = 0.0
    for ch in text:
        if ch.isspace():
            units += 0.35
        elif unicodedata.east_asian_width(ch) in {"F", "W", "A"}:
            units += 1.0
        elif ch in ".,:;!|/\\-()[]{}":
            units += 0.35
        else:
            units += 0.58
    return units * font_pt


def text_box_name(elem: ET.Element) -> str:
    for child in elem.iter():
        if local_name(child.tag) == "cNvPr":
            return child.attrib.get("name") or child.attrib.get("id") or local_name(elem.tag)
    return local_name(elem.tag)


def text_overflow_candidates(root: ET.Element) -> list[dict]:
    candidates: list[dict] = []
    for elem in root.iter():
        if local_name(elem.tag) not in {"sp", "graphicFrame"}:
            continue
        texts = [t.text for t in elem.findall(f".//{A}t") if t.text]
        text = "".join(texts).strip()
        if len(text) < 16:
            continue

        xfrm = elem.find(f".//{A}xfrm")
        if xfrm is None:
            continue
        ext = xfrm.find(f"{A}ext")
        if ext is None:
            continue
        try:
            width_pt = int(ext.attrib.get("cx", "0")) / EMU_PER_POINT
            height_pt = int(ext.attrib.get("cy", "0")) / EMU_PER_POINT
        except ValueError:
            continue
        if width_pt <= 0 or height_pt <= 0:
            continue

        font_sizes: list[float] = []
        for rpr in elem.findall(f".//{A}rPr") + elem.findall(f".//{A}defRPr") + elem.findall(f".//{A}endParaRPr"):
            sz = rpr.attrib.get("sz")
            if sz and sz.isdigit():
                font_sizes.append(int(sz) / 100)
        font_pt = max(font_sizes) if font_sizes else 20.0

        paragraphs = []
        for para in elem.findall(f".//{A}p"):
            para_text = "".join(t.text or "" for t in para.findall(f".//{A}t")).strip()
            if para_text:
                paragraphs.append(para_text)
        if not paragraphs:
            paragraphs = [text]

        usable_width = max(width_pt * 0.88, 1.0)
        estimated_lines = 0
        for para in paragraphs:
            para_width = text_width_units(para, font_pt)
            estimated_lines += max(1, math.ceil(para_width / usable_width))

        estimated_height = estimated_lines * font_pt * 1.22
        if estimated_height > height_pt * 1.05:
            candidates.append(
                {
                    "shape": text_box_name(elem),
                    "text_chars": len(text),
                    "font_pt": round(font_pt, 1),
                    "box_height_pt": round(height_pt, 1),
                    "estimated_height_pt": round(estimated_height, 1),
                }
            )
    return candidates


def analyze_slide(zf: zipfile.ZipFile, slide_name: str, slide_w: int | None, slide_h: int | None) -> dict:
    root = parse_xml(zf, slide_name)
    if root is None:
        return {"slide": slide_name, "error": "missing"}

    texts = [t.text.strip() for t in root.findall(f".//{A}t") if t.text and t.text.strip()]
    pictures = root.findall(f".//{PML}pic")
    graphic_frames = root.findall(f".//{PML}graphicFrame")
    bounds = shape_bounds(root)

    full_slide_pictures = 0
    bottom_blank_ratio = None
    if slide_w and slide_h:
        slide_area = slide_w * slide_h
        meaningful_bottoms: list[int] = []
        for x, y, cx, cy, tag in bounds:
            if tag == "pic" and slide_area and (cx * cy / slide_area) >= 0.88:
                full_slide_pictures += 1
            if slide_area and (cx * cy / slide_area) >= 0.003:
                meaningful_bottoms.append(y + cy)
        if meaningful_bottoms:
            bottom_blank_ratio = max(0.0, (slide_h - max(meaningful_bottoms)) / slide_h)

    text_chars = sum(len(t) for t in texts)
    joined_text = "".join(texts)
    bg_rgb = slide_background_rgb(root)
    nonwhite_full_slide = full_slide_nonwhite_shapes(root, slide_w, slide_h)
    overflow_candidates = text_overflow_candidates(root)
    forbidden_terms = [term for term in FORBIDDEN_SLIDE_TERMS if term in joined_text]
    return {
        "slide": slide_name,
        "text_runs": len(texts),
        "text_chars": text_chars,
        "pictures": len(pictures),
        "graphic_frames": len(graphic_frames),
        "full_slide_picture_candidates": full_slide_pictures,
        "bottom_blank_ratio": bottom_blank_ratio,
        "background_rgb": bg_rgb,
        "full_slide_background_risks": nonwhite_full_slide,
        "text_overflow_candidates": overflow_candidates,
        "forbidden_slide_terms": forbidden_terms,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a dissertation defense PPTX structure.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--min-slides", type=int, default=18)
    parser.add_argument("--max-slides", type=int, default=25)
    parser.add_argument("--json", action="store_true", help="Emit JSON report.")
    args = parser.parse_args()

    if not args.pptx.exists():
        print(f"ERROR: file not found: {args.pptx}", file=sys.stderr)
        return 2

    warnings: list[str] = []
    errors: list[str] = []

    try:
        with zipfile.ZipFile(args.pptx) as zf:
            names = zf.namelist()
            prs = parse_xml(zf, "ppt/presentation.xml")
            slide_w, slide_h = get_slide_size(prs)
            slide_names = sorted(
                [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
                key=lambda n: int(Path(n).stem.replace("slide", "")),
            )
            media_names = [n for n in names if n.startswith("ppt/media/")]
            slide_reports = [analyze_slide(zf, n, slide_w, slide_h) for n in slide_names]
    except zipfile.BadZipFile:
        print(f"ERROR: not a valid PPTX zip package: {args.pptx}", file=sys.stderr)
        return 2

    slide_count = len(slide_names)
    if slide_count < args.min_slides or slide_count > args.max_slides:
        warnings.append(f"Slide count is {slide_count}; expected {args.min_slides}-{args.max_slides}.")

    aspect = None
    if slide_w and slide_h:
        aspect = slide_w / slide_h
        if not math.isclose(aspect, 16 / 9, rel_tol=0.02):
            errors.append(f"Slide aspect ratio is {aspect:.3f}; expected approximately 16:9.")
    else:
        errors.append("Could not read slide size from ppt/presentation.xml.")

    if not media_names:
        warnings.append("No media files found; a defense PPT should usually include thesis figures or logos.")

    empty_text = [r["slide"] for r in slide_reports if r.get("text_chars", 0) == 0]
    if empty_text:
        warnings.append(f"{len(empty_text)} slide(s) contain no native text; check editability.")

    screenshot_like = [
        r["slide"]
        for r in slide_reports
        if r.get("full_slide_picture_candidates", 0) > 0 and r.get("text_chars", 0) < 20
    ]
    if screenshot_like:
        warnings.append(
            f"{len(screenshot_like)} slide(s) look like possible full-slide screenshots; "
            "verify the deck has not been flattened."
        )

    nonwhite_backgrounds = [
        r["slide"]
        for r in slide_reports
        if r.get("background_rgb") not in {None, "FFFFFF"}
    ]
    if nonwhite_backgrounds:
        warnings.append(
            f"{len(nonwhite_backgrounds)} slide(s) declare non-white slide backgrounds; "
            "the dissertation2ppt default requires pure white backgrounds."
        )

    full_slide_bg_risks = [
        r["slide"]
        for r in slide_reports
        if r.get("full_slide_background_risks")
    ]
    if full_slide_bg_risks:
        warnings.append(
            f"{len(full_slide_bg_risks)} slide(s) contain full-slide images or non-white shapes; "
            "check that these are not colored/photo backgrounds."
        )

    lower_half_sparse = [
        r["slide"]
        for r in slide_reports
        if r.get("bottom_blank_ratio") is not None
        and r.get("bottom_blank_ratio", 0) >= 0.28
        and r.get("text_chars", 0) >= 20
    ]
    if lower_half_sparse:
        warnings.append(
            f"{len(lower_half_sparse)} slide(s) leave roughly the lower third unused; "
            "use richer evidence, structured cards, process diagrams, tables, or larger direct figure placement."
        )

    overflow_slides = [
        r["slide"]
        for r in slide_reports
        if r.get("text_overflow_candidates")
    ]
    if overflow_slides:
        warnings.append(
            f"{len(overflow_slides)} slide(s) have possible text overflow; "
            "shorten text, enlarge boxes, reduce hierarchy safely, or split content."
        )

    forbidden_term_slides = [
        r["slide"]
        for r in slide_reports
        if r.get("forbidden_slide_terms")
    ]
    if forbidden_term_slides:
        warnings.append(
            f"{len(forbidden_term_slides)} slide(s) contain internal/audit terms such as source labels "
            "or '答辩提纲'; remove them from expert-facing slides."
        )

    report = {
        "file": str(args.pptx),
        "slide_count": slide_count,
        "slide_size_emu": {"width": slide_w, "height": slide_h},
        "aspect_ratio": aspect,
        "media_files": len(media_names),
        "slides": slide_reports,
        "warnings": warnings,
        "errors": errors,
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"File: {args.pptx}")
        print(f"Slides: {slide_count}")
        if slide_w and slide_h:
            print(f"Size: {slide_w} x {slide_h} EMU, aspect {aspect:.3f}")
        print(f"Media files: {len(media_names)}")
        for warning in warnings:
            print(f"WARNING: {warning}")
        for error in errors:
            print(f"ERROR: {error}")
        if not warnings and not errors:
            print("OK: structural checks passed.")

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
