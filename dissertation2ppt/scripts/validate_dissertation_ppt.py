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
import re
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
    "答辩逻辑",
    "归纳出的答辩主线",
    "答辩主线",
    "答辩人的思考",
    "注意事项",
    "注意的点",
    "内部思考",
    "生成思路",
    "页面规划",
    "可能被问到",
    "备用回答",
    "讲稿",
    "演讲稿",
    "speaker notes",
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


def shape_text_items(root: ET.Element) -> list[dict]:
    items: list[dict] = []
    for elem in root.iter():
        if local_name(elem.tag) not in {"sp", "graphicFrame"}:
            continue
        texts = [t.text for t in elem.findall(f".//{A}t") if t.text]
        text = "".join(texts).strip()
        if not text:
            continue

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

        alignments: list[str] = []
        for para in elem.findall(f".//{A}p"):
            para_text = "".join(t.text or "" for t in para.findall(f".//{A}t")).strip()
            if not para_text:
                continue
            ppr = para.find(f"{A}pPr")
            alignments.append(ppr.attrib.get("algn", "default") if ppr is not None else "default")

        items.append(
            {
                "name": text_box_name(elem),
                "text": text,
                "x": x,
                "y": y,
                "cx": cx,
                "cy": cy,
                "alignments": alignments or ["default"],
            }
        )
    return items


def normalize_heading_number(value: str) -> str:
    stripped = value.strip().rstrip(".、")
    if stripped.isdigit():
        return str(int(stripped))
    return stripped


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


def cover_alignment_issues(root: ET.Element, slide_w: int | None, slide_h: int | None) -> list[dict]:
    if not slide_w or not slide_h:
        return []
    issues: list[dict] = []
    for item in shape_text_items(root):
        text = item["text"]
        if len(text) < 2:
            continue
        if item["y"] > slide_h * 0.9:
            continue
        if item["cx"] < slide_w * 0.12:
            continue
        # Cover titles and metadata may sit in centered boxes while paragraphs
        # silently default to left alignment. That is the regression to catch.
        if any(algn != "ctr" for algn in item["alignments"]):
            issues.append(
                {
                    "shape": item["name"],
                    "text_sample": text[:40],
                    "alignments": item["alignments"],
                }
            )
    return issues


def duplicate_header_number_issues(root: ET.Element, slide_w: int | None, slide_h: int | None) -> list[dict]:
    if not slide_w or not slide_h:
        return []
    top_items = [
        item
        for item in shape_text_items(root)
        if item["y"] < slide_h * 0.22 and item["x"] < slide_w * 0.86 and len(item["text"]) <= 80
    ]
    number_badges: list[tuple[str, dict]] = []
    numbered_titles: list[tuple[str, dict]] = []
    for item in top_items:
        text = item["text"].strip()
        badge_match = re.fullmatch(r"([0-9]{1,2}|[一二三四五六七八九十]{1,3})[.、]?", text)
        if badge_match:
            number_badges.append((normalize_heading_number(badge_match.group(1)), item))
            continue
        title_match = re.match(r"^([0-9]{1,2}|[一二三四五六七八九十]{1,3})[\s.、]+", text)
        if title_match:
            numbered_titles.append((normalize_heading_number(title_match.group(1)), item))

    issues: list[dict] = []
    for badge_number, badge in number_badges:
        for title_number, title in numbered_titles:
            if badge_number == title_number and badge is not title:
                issues.append(
                    {
                        "number": badge_number,
                        "badge_shape": badge["name"],
                        "title_shape": title["name"],
                        "title_sample": title["text"][:60],
                    }
                )
    return issues


def navigation_like_candidates(root: ET.Element, slide_w: int | None, slide_h: int | None) -> list[dict]:
    if not slide_w or not slide_h:
        return []
    candidates: list[dict] = []
    standalone_number_items: list[tuple[str, dict]] = []
    for item in shape_text_items(root):
        in_nav_zone = item["y"] > slide_h * 0.66 or item["x"] < slide_w * 0.16
        if not in_nav_zone:
            continue
        exact_number = re.fullmatch(r"0?[1-9]", item["text"].strip())
        if exact_number:
            standalone_number_items.append((normalize_heading_number(exact_number.group(0)), item))
            continue
        tokens = [normalize_heading_number(t) for t in re.findall(r"(?<!\d)0?[1-9](?!\d)", item["text"])]
        unique_tokens = sorted(set(tokens), key=lambda value: int(value) if value.isdigit() else 99)
        nav_text_only = re.fullmatch(r"[\s0-9./|·-]+", item["text"].strip()) is not None
        if len(unique_tokens) >= 3 and len(item["text"]) <= 40 and nav_text_only:
            candidates.append(
                {
                    "shape": item["name"],
                    "text_sample": item["text"][:80],
                    "tokens": unique_tokens,
                    "zone": "bottom" if item["y"] > slide_h * 0.66 else "left",
                }
            )
    unique_standalone = sorted(set(number for number, _item in standalone_number_items), key=int)
    if len(unique_standalone) >= 3:
        candidates.append(
            {
                "shape": "multiple standalone number shapes",
                "text_sample": " ".join(unique_standalone),
                "tokens": unique_standalone,
                "zone": "bottom/left",
            }
        )
    return candidates


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

    slide_number = int(Path(slide_name).stem.replace("slide", ""))
    texts = [t.text.strip() for t in root.findall(f".//{A}t") if t.text and t.text.strip()]
    pictures = root.findall(f".//{PML}pic")
    graphic_frames = root.findall(f".//{PML}graphicFrame")
    bounds = shape_bounds(root)

    full_slide_pictures = 0
    bottom_blank_ratio = None
    content_picture_area_ratios: list[float] = []
    if slide_w and slide_h:
        slide_area = slide_w * slide_h
        meaningful_bottoms: list[int] = []
        for x, y, cx, cy, tag in bounds:
            if tag == "pic" and slide_area and (cx * cy / slide_area) >= 0.88:
                full_slide_pictures += 1
            if tag == "pic" and slide_area:
                area_ratio = cx * cy / slide_area
                likely_logo = x > slide_w * 0.78 and y < slide_h * 0.18 and area_ratio < 0.04
                if not likely_logo:
                    content_picture_area_ratios.append(area_ratio)
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
        "slide_number": slide_number,
        "text_runs": len(texts),
        "text_chars": text_chars,
        "pictures": len(pictures),
        "graphic_frames": len(graphic_frames),
        "full_slide_picture_candidates": full_slide_pictures,
        "bottom_blank_ratio": bottom_blank_ratio,
        "max_content_picture_area_ratio": max(content_picture_area_ratios) if content_picture_area_ratios else None,
        "background_rgb": bg_rgb,
        "full_slide_background_risks": nonwhite_full_slide,
        "text_overflow_candidates": overflow_candidates,
        "forbidden_slide_terms": forbidden_terms,
        "cover_alignment_issues": cover_alignment_issues(root, slide_w, slide_h) if slide_number == 1 else [],
        "duplicate_header_number_issues": duplicate_header_number_issues(root, slide_w, slide_h) if slide_number > 1 else [],
        "navigation_like_candidates": navigation_like_candidates(root, slide_w, slide_h) if slide_number > 2 else [],
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

    cover_alignment = [
        r["slide"]
        for r in slide_reports
        if r.get("cover_alignment_issues")
    ]
    if cover_alignment:
        warnings.append(
            "Cover slide has text boxes whose paragraph alignment is not centered; "
            "center the text inside title and metadata boxes, not only the boxes themselves."
        )

    duplicate_header_numbers = [
        r["slide"]
        for r in slide_reports
        if r.get("duplicate_header_number_issues")
    ]
    if duplicate_header_numbers:
        warnings.append(
            f"{len(duplicate_header_numbers)} slide(s) appear to repeat the same chapter number in the header; "
            "use either a number badge or a numbered title, not both."
        )

    navigation_like_slides = [
        r["slide"]
        for r in slide_reports
        if r.get("navigation_like_candidates")
    ]
    if navigation_like_slides:
        warnings.append(
            f"{len(navigation_like_slides)} slide(s) contain bottom/side multi-number patterns that look like "
            "chapter navigation; remove persistent chapter navigation from the deck."
        )

    small_evidence_image_slides = [
        r["slide"]
        for r in slide_reports
        if r.get("slide_number", 0) > 1
        and r.get("pictures", 0) > 0
        and r.get("text_chars", 0) >= 20
        and r.get("max_content_picture_area_ratio") is not None
        and r.get("max_content_picture_area_ratio", 1.0) < 0.12
    ]
    if small_evidence_image_slides:
        warnings.append(
            f"{len(small_evidence_image_slides)} slide(s) have only small content images; "
            "key thesis figures and visual comparisons should be enlarged or placed in a stronger layout."
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
