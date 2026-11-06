#!/usr/bin/env python3
# Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
# Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
# Production use requires a commercial agreement: licensing@eigenpal.com
#
# Copied from DOCX Editor `packages/docx-to-pdf/scripts/evaluation/worker.py`, commit
# 398963c7f67207cebf3c38dcc41d5f0bd1339b5f. PPTX changes, each marked "PPTX:":
# - The visual sketch is landscape, 192 by 144, because slides are landscape.
# - Visual regions are named top, body, and bottom. A slide has no running header or footer.
# - Rendering lowers the resolution to stay inside the pixel budget, so a large custom slide
#   size still measures. `evidence` and `page` report the resolution they used.
# - The word movement module is `word_movement.py` beside this file.
# - Color sketches of different sizes are refused instead of compared as grayscale.
# - `compare` pairs pages by number, and `--candidate-report` renumbers candidate pages to
#   reference pages from the export's `slideNumbers`, because a slide that the engine cannot
#   lay out has no candidate page.
# - The screening and page-size limits come from `thresholds.json` beside this file, generated
#   from `scripts/fidelity/thresholds.ts`, where every limit of the fidelity gates lives.
"""Eval protocol v1: cached PDF measurements, comparisons, and page evidence."""

import argparse
import base64
import functools
import gzip
import hashlib
import importlib.util
import json
from pathlib import Path

import pymupdf
from PIL import Image, ImageChops, ImageOps

PROTOCOL = 1
MAX_PAGES = 2000
MAX_WORDS = 1_000_000
MAX_PIXELS = 16_000_000
# PPTX: landscape, with the same number of samples as DOCX Editor's portrait 144 by 192.
RGB_SIZE = (192, 144)
# PPTX: the comparison B limits that the worker reads, from `scripts/fidelity/thresholds.ts`
# through `thresholds.json`.
SCREEN_KEYS = (
    "colorChannelDelta",
    "grayDelta",
    "tileSamples",
    "minChangedSamplesInTile",
    "maxChangedFraction",
    "textShiftPt",
)


@functools.cache
def limits():
    """PPTX: the comparison B limits, read and checked when a command first needs them, so
    `measure` and `page` run without the file and a broken file fails with one clear error."""
    path = Path(__file__).with_name("thresholds.json")
    try:
        value = json.loads(path.read_text())["comparisonB"]
        numbers = [value["maxPageSizeDeltaPt"], *(value["screening"][key] for key in SCREEN_KEYS)]
    except (OSError, ValueError, KeyError, TypeError) as error:
        raise ValueError(f"{path.name} is missing or incomplete; regenerate it from thresholds.ts") from error
    if not all(isinstance(n, (int, float)) and not isinstance(n, bool) and n >= 0 for n in numbers):
        raise ValueError(f"{path.name} has a limit that is not a number")
    return value


def write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode()
    path.write_bytes(gzip.compress(data, mtime=0) if path.suffix == ".gz" else data)


def read(path):
    path = Path(path)
    data = path.read_bytes()
    return json.loads(gzip.decompress(data) if path.suffix == ".gz" else data)


def fitted_dpi(page, dpi):
    """PPTX: the highest resolution up to `dpi` that keeps the page inside MAX_PIXELS."""
    rect = page.rect
    if rect.width <= 0 or rect.height <= 0:
        raise ValueError("Page has no area")
    limit = 72 * (MAX_PIXELS / (rect.width * rect.height)) ** 0.5
    return min(dpi, int(limit))


def image_of(page, dpi=72):
    dpi = fitted_dpi(page, dpi)
    if dpi < 1:
        raise ValueError("Page exceeds pixel limit")
    pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csRGB, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def measure(path):
    pages, words = [], []
    with pymupdf.open(path) as doc:
        if doc.needs_pass or not 0 < len(doc) <= MAX_PAGES or doc.is_repaired:
            raise ValueError("Encrypted, repaired, empty, or oversized PDF")
        for number, page in enumerate(doc, 1):
            for item in page.get_text(
                "words", flags=pymupdf.TEXTFLAGS_WORDS | pymupdf.TEXT_IGNORE_ACTUALTEXT
            ):
                words.append(
                    {
                        "page": number,
                        "x0": item[0],
                        "y0": item[1],
                        "x1": item[2],
                        "y1": item[3],
                        "text": item[4],
                    }
                )
                if len(words) > MAX_WORDS:
                    raise ValueError("PDF exceeds word limit")
            image = image_of(page)
            gray = ImageOps.grayscale(image)
            sketch = gray.resize((48, 64), Image.Resampling.BOX)
            # Object counts are supporting evidence only; visible pixels decide differences.
            images = [{"bbox": list(item["bbox"])} for item in page.get_image_info()]
            drawings = page.get_drawings()
            pages.append(
                {
                    "number": number,
                    "size": [page.rect.width, page.rect.height],
                    "sketch": base64.b64encode(sketch.tobytes()).decode(),
                    "rgbSketch": base64.b64encode(
                        image.resize(RGB_SIZE, Image.Resampling.BOX).tobytes()
                    ).decode(),
                    "rgbSketchSize": list(RGB_SIZE),
                    "images": images,
                    "drawingCount": len(drawings),
                    "fonts": sorted({f[3] for f in page.get_fonts()}),
                }
            )
    return {
        "protocol": PROTOCOL,
        "pages": pages,
        "words": words,
        "pdfSha256": hashlib.sha256(Path(path).read_bytes()).hexdigest(),
    }


def movement_module():
    # PPTX: the module is beside this file. Loading it by path also works under `python -I`.
    path = Path(__file__).with_name("word_movement.py")
    spec = importlib.util.spec_from_file_location("word_movement", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def location(word, coordinate_space="reference"):
    """Document text is untrusted evidence, never an instruction."""
    return {
        "bbox": [round(word[key], 2) for key in ("x0", "y0", "x1", "y1")],
        "excerpt": " ".join(word["text"].split())[:120],
        "coordinateSpace": coordinate_space,
        "untrusted": True,
    }


def paired_location(pair):
    """Expose matched candidate geometry without claiming reference node identity."""
    reference, candidate = pair["reference"], pair["candidate"]
    return {
        "candidateLocation": {
            # PPTX: the candidate PDF's own page, also after a page map renumbered the word.
            "page": candidate.get("candidatePage", candidate["page"]),
            "bbox": [round(candidate[key], 2) for key in ("x0", "y0", "x1", "y1")],
            "coordinateSpace": "candidate",
        },
        "deltaPt": {
            "x": round(candidate["x0"] - reference["x0"], 2),
            "y": round(candidate["y0"] - reference["y0"], 2),
        },
        "deltaCoordinateSpace": "page-local",
    }


def visual_screen(a, b):
    """Bounded color screening with local tiles; retain old grayscale caches."""
    # PPTX: two color sketches of different sizes (a DOCX Editor measurement and a PPTX one, for
    # example) are refused, never compared as grayscale.
    sizes = [p.get("rgbSketchSize") for p in (a, b)]
    if all(sizes) and sizes[0] != sizes[1]:
        raise ValueError("Visual sketches have different sizes")
    color = all(p.get("rgbSketchSize") == list(RGB_SIZE) for p in (a, b))
    width, height = RGB_SIZE if color else (48, 64)
    channels = 3 if color else 1
    key = "rgbSketch" if color else "sketch"
    av, bv = (base64.b64decode(p[key], validate=True) for p in (a, b))
    if any(len(v) != width * height * channels for v in (av, bv)):
        raise ValueError("Invalid visual sketch size")
    # Pillow computes byte differences in native code; retain the exact per-pixel maximum.
    mode = "RGB" if color else "L"
    difference = ImageChops.difference(
        Image.frombytes(mode, (width, height), av),
        Image.frombytes(mode, (width, height), bv),
    )
    if color:
        red, green, blue = difference.split()
        difference = ImageChops.lighter(ImageChops.lighter(red, green), blue)
    delta = difference.tobytes()
    screen = limits()["screening"]
    threshold = screen["colorChannelDelta"] if color else screen["grayDelta"]
    changed = [i for i, d in enumerate(delta) if d > threshold]
    fraction = len(changed) / len(delta)
    tiles = {}
    for i in changed:
        tile = ((i // width) // screen["tileSamples"], (i % width) // screen["tileSamples"])
        tiles.setdefault(tile, []).append(i)
    # Three changed samples in a tile catch small graphics without one-pixel noise.
    significant = [t for t, indices in tiles.items() if len(indices) >= screen["minChangedSamplesInTile"]]
    detected = fraction > screen["maxChangedFraction"] or (color and bool(significant))
    bbox = None
    region = None
    if detected and changed:
        tile = min(significant) if significant else min(tiles)
        indices = tiles[tile]
        xs, ys = [i % width for i in indices], [i // width for i in indices]
        bbox = [
            round(min(xs) * a["size"][0] / width, 2),
            round(min(ys) * a["size"][1] / height, 2),
            round((max(xs) + 1) * a["size"][0] / width, 2),
            round((max(ys) + 1) * a["size"][1] / height, 2),
        ]
        # PPTX: a slide has no running header or footer, so the regions are positional.
        region = "top" if min(ys) < height / 8 else ("bottom" if min(ys) >= height * 7 / 8 else "body")
    return {
        "changedFraction": round(fraction, 5),
        "mode": "local-rgb" if color else "legacy-grayscale",
        "detected": detected,
        "bbox": bbox,
        "region": region,
    }


def visible_slides(report):
    """PPTX: the visible slide numbers of an export report, after its counts are checked."""
    count = report["slideCount"]
    if not isinstance(count, int) or not 0 <= count <= MAX_PAGES:
        raise ValueError("Export report has an invalid slide count")
    hidden = report.get("hiddenSlides", [])
    if not isinstance(hidden, list) or not all(isinstance(n, int) and 1 <= n <= count for n in hidden):
        raise ValueError("Export report hidden slides must be slide numbers within the slide count")
    hidden = set(hidden)
    return [n for n in range(1, count + 1) if n not in hidden]


def page_map(report):
    """PPTX: candidate page -> reference page, from an export report. Reference page i is
    the i-th visible slide; a candidate page whose slide is hidden maps to nothing."""
    visible = visible_slides(report)
    count = report["slideCount"]
    numbers = report["slideNumbers"]
    if not all(isinstance(n, int) and 1 <= n <= count for n in numbers) or any(
        b <= a for a, b in zip(numbers, numbers[1:])
    ):
        raise ValueError("Export report slide numbers must increase within the slide count")
    reference_page = {slide: page for page, slide in enumerate(visible, 1)}
    return {
        page: reference_page[slide]
        for page, slide in enumerate(report["slideNumbers"], 1)
        if slide in reference_page
    }


def remap(measurement, mapping):
    """PPTX: renumber a measurement's pages and words to reference pages; drop pages without a
    reference page. Words keep their own page as `candidatePage`."""
    value = dict(measurement)
    value["pages"] = [
        dict(page, number=mapping[page["number"]])
        for page in measurement["pages"]
        if page["number"] in mapping
    ]
    value["words"] = [
        dict(word, page=mapping[word["page"]], candidatePage=word["page"])
        for word in measurement["words"]
        if word["page"] in mapping
    ]
    return value


def checked_map(report, candidate_pages, reference_pages):
    """PPTX: the page map of a report that belongs to the candidate PDF, for a reference with one
    page per visible slide, or an error. Without that page count every later slide would pair
    with the wrong reference page."""
    if len(report["slideNumbers"]) != candidate_pages:
        raise ValueError("The export report does not match the candidate PDF: page counts differ")
    visible = len(visible_slides(report))
    if reference_pages != visible:
        raise ValueError(
            f"The reference PDF has {reference_pages} pages for {visible} visible slides; "
            "a reference PDF leaves hidden slides out"
        )
    return page_map(report)


PAGE_MARK = "\u241f"


def unmark(value):
    """PPTX: remove the page qualifier from every string in a movement report."""
    if isinstance(value, str):
        return value.split(PAGE_MARK, 1)[1] if PAGE_MARK in value else value
    if isinstance(value, list):
        return [unmark(item) for item in value]
    if isinstance(value, dict):
        return {key: unmark(item) for key, item in value.items()}
    return value


def compare(left, right, report=None):
    if left.get("protocol") != PROTOCOL or right.get("protocol") != PROTOCOL:
        raise ValueError("Unsupported measurement protocol")
    candidate_page_count = len(right["pages"])
    dropped = []
    if report is not None:
        mapping = checked_map(report, candidate_page_count, len(left["pages"]))
        # PPTX: candidate pages of hidden slides have no reference page.
        dropped = [page["number"] for page in right["pages"] if page["number"] not in mapping]
        right = remap(right, mapping)
        # PPTX: with a page map, a page is a slide, and text cannot move between slides. Each word
        # matches only on its own slide: its text carries the page while words are matched.
        left = dict(left, words=[dict(w, text=f"{w['page']}{PAGE_MARK}{w['text']}") for w in left["words"]])
        right = dict(right, words=[dict(w, text=f"{w['page']}{PAGE_MARK}{w['text']}") for w in right["words"]])
    movement, pairs = movement_module().compare_word_movement(left["words"], right["words"])
    if report is not None:
        for word in left["words"] + right["words"]:
            word["text"] = unmark(word["text"])
        movement = unmark(movement)
        # PPTX: every `candidatePage` is the candidate PDF's own page, as `candidateLocation.page`.
        inverse = {ref: cand for cand, ref in mapping.items()}
        for key in ("largestMovements", "earliestMovements"):
            for item in movement.get(key, []):
                item["candidatePage"] = inverse.get(item["candidatePage"], item["candidatePage"])
    findings, screens = [], []

    def add(kind, severity, page, **details):
        findings.append(dict(kind=kind, severity=severity, page=page, **details))

    # PPTX: one finding on each reference page (slide) that has no candidate page, and one for
    # candidate pages beyond the reference.
    paired = {page["number"] for page in right["pages"]}
    for page in left["pages"]:
        if page["number"] not in paired:
            add(
                "pagination",
                4,
                page["number"],
                referencePages=len(left["pages"]),
                candidatePages=candidate_page_count,
                bbox=None,
                excerpt="",
                confidence="high",
                coordinateSpace="reference",
                note="This slide has no candidate page.",
            )
    # Without a page map, candidate pages past the reference; with one, candidate pages of hidden
    # slides. Both are candidate page numbers.
    reference_numbers = {page["number"] for page in left["pages"]}
    extra = sorted(dropped + [page["number"] for page in right["pages"] if page["number"] not in reference_numbers])
    if extra:
        add(
            "pagination",
            4,
            extra[0],
            referencePages=len(left["pages"]),
            candidatePages=candidate_page_count,
            extraCandidatePages=extra[:100],
            bbox=None,
            excerpt="",
            confidence="high",
            coordinateSpace="candidate",
            note="These candidate pages have no reference page.",
        )
    cross = movement["distanceBuckets"]["crossPage"]
    if cross:
        first = min(
            (p for p in pairs if p["reference"]["page"] != p["candidate"]["page"]),
            key=lambda p: (p["reference"]["page"], p["reference"]["y0"], p["reference"]["x0"]),
        )
        add(
            "pagination",
            4,
            first["reference"]["page"],
            movedWords=cross,
            confidence="high",
            **location(first["reference"]),
            **paired_location(first),
        )
    for source, side, kind in (
        (left, "reference", "text-coverage"),
        (right, "candidate", "text-added"),
    ):
        matched = {id(p[side]) for p in pairs}
        unmatched = [w for w in source["words"] if id(w) not in matched]
        if unmatched:
            first = min(unmatched, key=lambda w: (w["page"], w["y0"], w["x0"]))
            add(
                kind,
                3,
                first["page"],
                unmatchedWords=len(unmatched),
                confidence="medium",
                **location(first, side),
                **({"candidatePage": first.get("candidatePage", first["page"])} if side == "candidate" else {}),
                note="Unmatched extraction does not prove missing or added visible text.",
            )
    # PPTX: pair pages by number, so a page missing from the candidate shifts nothing.
    candidate_pages = {page["number"]: page for page in right["pages"]}
    for a in left["pages"]:
        n = a["number"]
        b = candidate_pages.get(n)
        if b is None:
            continue
        if any(abs(x - y) > limits()["maxPageSizeDeltaPt"] for x, y in zip(a["size"], b["size"])):
            add(
                "page-geometry",
                4,
                n,
                referenceSize=a["size"],
                candidateSize=b["size"],
                bbox=[0, 0, *a["size"]],
                excerpt="",
                confidence="high",
                coordinateSpace="reference",
            )
        screen = visual_screen(a, b)
        screens.append(dict(page=n, **screen))
        if screen["detected"]:
            add(
                "visual-region",
                2,
                n,
                region=screen["region"],
                changedFraction=screen["changedFraction"],
                bbox=screen["bbox"],
                excerpt="",
                coordinateSpace="reference",
                confidence="low",
                screeningMode=screen["mode"],
                note="Visual screening; inspect detailed evidence before assigning a cause.",
            )
    shifted = [p for p in pairs if p["distancePt"] is not None and p["distancePt"] > limits()["screening"]["textShiftPt"]]
    if shifted:
        first = min(shifted, key=lambda p: (p["reference"]["page"], p["reference"]["y0"]))
        add(
            "text-position",
            3 if movement["distanceBuckets"]["beyond8Pt"] else 2,
            first["reference"]["page"],
            shiftedWords=len(shifted),
            firstY=round(first["reference"]["y0"], 2),
            confidence="high",
            **location(first["reference"]),
            **paired_location(first),
        )
    # Prefer grounded geometry/text over low-confidence raster screening.
    grounded = [f for f in findings if f.get("confidence") in ("high", "medium")]
    earliest = min(
        grounded or findings,
        key=lambda f: (
            f["page"],
            (f.get("bbox") or [0, 0])[1],
            (f.get("bbox") or [0, 0])[0],
        ),
        default=None,
    )
    divergence = (
        None
        if earliest is None
        else {
            "type": earliest["kind"],
            "page": earliest["page"],
            "bbox": earliest.get("bbox"),
            "excerpt": earliest.get("excerpt", "")[:120],
            "confidence": earliest["confidence"],
            "coordinateSpace": earliest.get("coordinateSpace", "reference"),
            "untrusted": True,
            **{
                key: earliest[key]
                for key in ("candidateLocation", "deltaPt", "deltaCoordinateSpace")
                if key in earliest
            },
        }
    )
    return {
        "protocol": PROTOCOL,
        "findings": findings,
        "movement": movement,
        "pages": screens,
        "referencePages": len(left["pages"]),
        "candidatePages": len(right["pages"]),
        "screeningOnly": True,
        "firstDivergence": divergence,
    }


def evidence(left, right, output, pages, candidate_report=None):
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    report = []
    with pymupdf.open(left) as a, pymupdf.open(right) as b:
        # PPTX: `pages` are reference pages; the export report maps each to its candidate page.
        if candidate_report is not None:
            mapping = checked_map(candidate_report, len(b), len(a))
            inverse = {ref: cand for cand, ref in mapping.items()}
        for number in pages[:3]:
            images = []
            other = inverse.get(number, 0) if candidate_report is not None else number
            for doc, name, index in ((a, "reference", number), (b, "candidate", other)):
                image = (
                    image_of(doc[index - 1], 144)
                    if 0 < index <= len(doc)
                    else Image.new("RGB", (1, 1), "white")
                )
                image.save(output / f"{number}-{name}.png")
                images.append(image)
            width = max(i.width for i in images)
            height = max(i.height for i in images)
            padded = []
            for image in images:
                canvas = Image.new("RGB", (width, height), "white")
                canvas.paste(image, (0, 0))
                padded.append(canvas)
            diff = ImageChops.difference(*padded)
            strong = ImageOps.grayscale(diff).point(lambda v: 255 if v >= 28 else 0)
            box = strong.getbbox()
            Image.blend(*padded, 0.5).save(output / f"{number}-overlay.png")
            if box:
                crop = (
                    max(0, box[0] - 24),
                    max(0, box[1] - 24),
                    min(width, box[2] + 24),
                    min(height, box[3] + 24),
                )
                pair = Image.new("RGB", ((crop[2] - crop[0]) * 2, crop[3] - crop[1]), "white")
                for index, image in enumerate(padded):
                    pair.paste(image.crop(crop), (index * (crop[2] - crop[0]), 0))
                pair.save(output / f"{number}-crop.png")
            report.append(
                {
                    "page": number,
                    "candidatePage": other if 0 < other <= len(b) else None,
                    "dpi": [
                        fitted_dpi(doc[index - 1], 144) if 0 < index <= len(doc) else None
                        for doc, index in ((a, number), (b, other))
                    ],
                    "changedPixels": sum(strong.histogram()[1:]),
                    "pixels": width * height,
                    "bbox": box,
                }
            )
    return {"protocol": PROTOCOL, "pages": report}


def preview_page(source, output, number):
    """Render one page and expose the transform for measured PDF coordinates."""
    output = Path(output)
    with pymupdf.open(source) as document:
        if document.needs_pass or not 1 <= number <= min(len(document), MAX_PAGES):
            raise ValueError("Page is unavailable")
        page = document[number - 1]
        image = image_of(page, 144)
        output.mkdir(parents=True, exist_ok=True)
        image.save(output / "page.png")
        return {
            "protocol": PROTOCOL,
            "page": number,
            "pages": len(document),
            "dpi": fitted_dpi(page, 144),
            "widthPt": page.rect.width,
            "heightPt": page.rect.height,
            "rotationMatrix": list(page.rotation_matrix),
        }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=["measure", "compare", "evidence", "page"])
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--candidate")
    parser.add_argument("--candidate-report")
    parser.add_argument("--pages", default="1")
    args = parser.parse_args()
    if args.operation == "measure":
        value = measure(args.input)
    elif args.operation == "compare":
        report = read(args.candidate_report) if args.candidate_report else None
        value = compare(read(args.input), read(args.candidate), report)
    elif args.operation == "page":
        value = preview_page(args.input, args.output, int(args.pages))
        write(Path(args.output) / "report.json", value)
        return
    else:
        value = evidence(
            args.input,
            args.candidate,
            args.output,
            [int(p) for p in args.pages.split(",")],
            read(args.candidate_report) if args.candidate_report else None,
        )
        write(Path(args.output) / "report.json", value)
        return
    write(args.output, value)


if __name__ == "__main__":
    main()
