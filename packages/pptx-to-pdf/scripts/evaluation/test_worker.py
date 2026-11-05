# Copyright (c) 2026 EigenPal, Inc. All rights reserved.
# Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
# Production use requires a commercial agreement: licensing@eigenpal.com
#
# Copied from DOCX Editor `packages/docx-to-pdf/scripts/evaluation/test_worker.py`, commit
# 398963c7f67207cebf3c38dcc41d5f0bd1339b5f. PPTX changes: slide-sized pages (960 by 540 points), and
# checks for large custom slide sizes and the reported resolution.
"""Synthetic protocol tests. Run with the evaluator's Python environment."""

import json
import tempfile
import unittest
from copy import deepcopy
from pathlib import Path

import pymupdf
from worker import RGB_SIZE, compare, evidence, measure, preview_page, read, write


class ProtocolTests(unittest.TestCase):
    def variant(
        self,
        path,
        text="Synthetic anchor",
        x=72,
        y=95,
        graphic=True,
        color=(1, 0, 0),
        text_page=1,
        metadata=None,
    ):
        with pymupdf.open() as doc:
            for number in range(1, text_page + 1):
                page = doc.new_page(width=960, height=540)
                if number == text_page:
                    page.insert_text((x, y), text)
                    if graphic:
                        page.draw_rect((300, 300, 312, 312), color=color, fill=color)
            if metadata:
                doc.set_metadata(metadata)
            doc.save(path)

    def variants(self, left=None, right=None):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.variant(root / "a.pdf", **(left or {}))
            self.variant(root / "b.pdf", **(right or {}))
            return compare(measure(root / "a.pdf"), measure(root / "b.pdf"))

    def pdf(self, path, pages=1, header=True, table=True):
        with pymupdf.open() as doc:
            for _ in range(pages):
                page = doc.new_page(width=960, height=540)
                if header:
                    page.insert_text((72, 35), "Synthetic recurring header")
                page.insert_text((72, 95), "Synthetic paragraph content")
                if table:
                    for y in (120, 150, 180):
                        page.draw_line((72, y), (300, y))
                    for x in (72, 180, 300):
                        page.draw_line((x, 120), (x, 180))
                    page.insert_text((80, 140), "Cell one")
            doc.save(path)

    def test_identical_pdf_is_equal(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "input.pdf"
            self.pdf(path)
            value = measure(path)
            self.assertEqual(compare(value, value)["findings"], [])
            self.assertIsNone(compare(value, value)["firstDivergence"])
            target = Path(temp) / "measurement.json.gz"
            write(target, value)
            self.assertEqual(read(target), value)

    def slides(self, path, titles):
        with pymupdf.open() as doc:
            for title in titles:
                page = doc.new_page(width=960, height=540)
                page.insert_text((72, 95), title)
                page.draw_rect((300, 300, 340, 340), color=(0, 0, 1), fill=(0, 0, 1))
            doc.save(path)

    def test_a_slide_without_a_candidate_page_shifts_nothing(self):
        # PPTX: slide 2 could not be laid out, so the candidate has pages for slides 1 and 3.
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.slides(root / "ref.pdf", ["Slide one", "Slide two", "Slide three"])
            self.slides(root / "cand.pdf", ["Slide one", "Slide three"])
            reference, candidate = measure(root / "ref.pdf"), measure(root / "cand.pdf")
            report = {"slideCount": 3, "hiddenSlides": [], "slideNumbers": [1, 3]}
            mapped = compare(reference, candidate, report)
            self.assertEqual([p["page"] for p in mapped["pages"]], [1, 3])
            self.assertFalse(any(p["detected"] for p in mapped["pages"]))
            kinds = {f["kind"] for f in mapped["findings"]}
            self.assertNotIn("text-position", kinds)
            # The finding names the slide without a candidate page.
            pagination = [f for f in mapped["findings"] if f["kind"] == "pagination"]
            self.assertEqual([f["page"] for f in pagination], [2])
            self.assertEqual(mapped["firstDivergence"]["page"], 2)
            # A report for another PDF is refused, never read as a pass.
            with self.assertRaises(ValueError):
                compare(reference, candidate, {"slideCount": 3, "hiddenSlides": [], "slideNumbers": [1]})
            # Evidence pairs reference page 3 with candidate page 2.
            shown = evidence(root / "ref.pdf", root / "cand.pdf", root / "ev", [2, 3], report)
            self.assertEqual([(p["page"], p["candidatePage"]) for p in shown["pages"]], [(2, None), (3, 2)])
            self.assertIsNone(shown["pages"][1]["bbox"])
            # Without the report, the candidate's page 2 is screened against slide 2.
            unmapped = compare(reference, candidate)
            self.assertTrue(any(p["detected"] for p in unmapped["pages"] if p["page"] == 2))

    def test_with_a_page_map_words_match_only_on_their_own_slide(self):
        # PPTX: "Revenue" is missing on slide 1 and extra on slide 2; it did not move.
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.slides(root / "ref.pdf", ["Slide one Revenue", "Slide two"])
            self.slides(root / "cand.pdf", ["Slide one", "Slide two Revenue"])
            report = {"slideCount": 2, "hiddenSlides": [], "slideNumbers": [1, 2]}
            result = compare(measure(root / "ref.pdf"), measure(root / "cand.pdf"), report)
            kinds = [f["kind"] for f in result["findings"]]
            self.assertNotIn("pagination", kinds)
            self.assertIn("text-coverage", kinds)
            self.assertIn("text-added", kinds)
            self.assertEqual(result["movement"]["distanceBuckets"]["crossPage"], 0)
            added = next(f for f in result["findings"] if f["kind"] == "text-added")
            self.assertEqual((added["page"], added["candidatePage"]), (2, 2))
            # `ensure_ascii=False`: the default escapes the mark, and the check could never fail.
            self.assertNotIn("\u241f", json.dumps(result, ensure_ascii=False))

    def test_a_reference_with_a_page_for_a_hidden_slide_is_refused(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.slides(root / "ref.pdf", ["Slide one", "Slide two", "Slide three"])
            self.slides(root / "cand.pdf", ["Slide one", "Slide three"])
            # Slide 2 is hidden, but the reference has a page for it: every later slide would
            # pair with the wrong page, so neither command reads it.
            report = {"slideCount": 3, "hiddenSlides": [2], "slideNumbers": [1, 3]}
            with self.assertRaisesRegex(ValueError, "3 pages for 2 visible slides"):
                compare(measure(root / "ref.pdf"), measure(root / "cand.pdf"), report)
            with self.assertRaisesRegex(ValueError, "3 pages for 2 visible slides"):
                evidence(root / "ref.pdf", root / "cand.pdf", root / "ev", [2], report)

    def test_a_candidate_page_of_a_hidden_slide_is_an_extra_page(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.slides(root / "ref.pdf", ["Slide one", "Slide three"])
            self.slides(root / "cand.pdf", ["Slide one", "Slide two", "Slide three"])
            report = {"slideCount": 3, "hiddenSlides": [2], "slideNumbers": [1, 2, 3]}
            result = compare(measure(root / "ref.pdf"), measure(root / "cand.pdf"), report)
            extra = [f for f in result["findings"] if f["kind"] == "pagination"]
            self.assertEqual([(f["page"], f["extraCandidatePages"]) for f in extra], [(2, [2])])
            self.assertEqual(extra[0]["coordinateSpace"], "candidate")

    def test_hidden_slides_map_to_no_reference_page(self):
        from worker import page_map

        report = {"slideCount": 4, "hiddenSlides": [2], "slideNumbers": [1, 2, 3, 4]}
        self.assertEqual(page_map(report), {1: 1, 3: 2, 4: 3})
        with self.assertRaises(ValueError):
            page_map({"slideCount": 10**9, "hiddenSlides": [], "slideNumbers": []})
        for numbers in ([2, 1], [0, 1], [1, 5], ["1"]):
            with self.assertRaises(ValueError):
                page_map({"slideCount": 4, "hiddenSlides": [], "slideNumbers": numbers})
        for hidden in (["2"], [9], [0], "2"):
            with self.assertRaises(ValueError):
                page_map({"slideCount": 4, "hiddenSlides": hidden, "slideNumbers": [1]})

    def test_moved_text_names_its_own_candidate_page(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.slides(root / "ref.pdf", ["Slide one", "Slide two", "Slide three"])
            with pymupdf.open() as doc:
                for title, y in (("Slide one", 95), ("Slide three", 115)):
                    page = doc.new_page(width=960, height=540)
                    page.insert_text((72, y), title)
                    page.draw_rect((300, 300, 340, 340), color=(0, 0, 1), fill=(0, 0, 1))
                doc.save(root / "cand.pdf")
            report = {"slideCount": 3, "hiddenSlides": [], "slideNumbers": [1, 3]}
            result = compare(measure(root / "ref.pdf"), measure(root / "cand.pdf"), report)
            moved = [f for f in result["findings"] if f["kind"] == "text-position"][0]
            self.assertEqual(moved["page"], 3)
            self.assertEqual(moved["candidateLocation"]["page"], 2)
            # Every `candidatePage` is the candidate PDF's page; `referencePage` is the slide.
            sample = next(m for m in result["movement"]["largestMovements"] if m["text"] == "three")
            self.assertEqual((sample["referencePage"], sample["candidatePage"]), (3, 2))

    def test_added_text_has_candidate_location(self):
        result = self.variants(right={"text": "Synthetic anchor added"})
        finding = next(f for f in result["findings"] if f["kind"] == "text-added")
        self.assertEqual(finding["unmatchedWords"], 1)
        self.assertEqual(finding["coordinateSpace"], "candidate")
        self.assertEqual(finding["excerpt"], "added")
        self.assertEqual(result["firstDivergence"]["type"], "text-added")

    def test_missing_text_uses_actual_page_and_bbox(self):
        result = self.variants(
            left={"text": "Synthetic anchor missing", "text_page": 2},
            right={"text_page": 2},
        )
        first = result["firstDivergence"]
        self.assertEqual(first["type"], "text-coverage")
        self.assertEqual(first["page"], 2)
        self.assertEqual(first["excerpt"], "missing")
        self.assertGreater(first["bbox"][0], 72)
        self.assertTrue(first["untrusted"])

    def test_small_graphic_removal_is_detected_locally(self):
        result = self.variants(right={"graphic": False})
        visual = next(f for f in result["findings"] if f["kind"] == "visual-region")
        self.assertLess(visual["changedFraction"], 0.01)
        self.assertEqual(visual["screeningMode"], "local-rgb")
        self.assertEqual(result["firstDivergence"]["confidence"], "low")
        self.assertLess(visual["bbox"][0], 312)
        self.assertGreater(visual["bbox"][2], 300)

    def test_equal_luminance_color_change_is_detected(self):
        # Red and this green have nearly equal grayscale luminance.
        result = self.variants(right={"color": (0, 0.51, 0)})
        self.assertTrue(any(f["kind"] == "visual-region" for f in result["findings"]))
        self.assertEqual(result["movement"]["missingWordCount"], 0)

    def test_text_shift_and_cross_page_have_grounded_divergence(self):
        shifted = self.variants(right={"y": 107})
        self.assertEqual(shifted["firstDivergence"]["type"], "text-position")
        self.assertEqual(shifted["firstDivergence"]["excerpt"], "Synthetic")
        shifted_first = shifted["firstDivergence"]
        self.assertEqual(shifted_first["candidateLocation"]["page"], 1)
        self.assertEqual(shifted_first["candidateLocation"]["coordinateSpace"], "candidate")
        self.assertEqual(shifted_first["deltaPt"], {"x": 0, "y": 12})
        self.assertEqual(shifted_first["deltaCoordinateSpace"], "page-local")
        self.assertAlmostEqual(shifted_first["candidateLocation"]["bbox"][1] - shifted_first["bbox"][1], 12)
        moved = self.variants(right={"text_page": 2})
        self.assertEqual(moved["firstDivergence"]["type"], "pagination")
        self.assertEqual(moved["firstDivergence"]["page"], 1)
        self.assertEqual(moved["firstDivergence"]["excerpt"], "Synthetic")
        self.assertEqual(moved["firstDivergence"]["candidateLocation"]["page"], 2)
        self.assertEqual(moved["firstDivergence"]["deltaPt"], {"x": 0, "y": 0})
        self.assertEqual(moved["firstDivergence"]["deltaCoordinateSpace"], "page-local")

    def test_paired_location_uses_the_first_shifted_word(self):
        result = self.variants(right={"x": 81, "y": 99})
        finding = next(f for f in result["findings"] if f["kind"] == "text-position")
        self.assertEqual(finding["deltaPt"], {"x": 9, "y": 4})
        self.assertEqual(finding["candidateLocation"]["bbox"][0], 81)
        self.assertEqual(result["firstDivergence"]["candidateLocation"], finding["candidateLocation"])

    def test_nonvisual_metadata_and_small_color_noise_are_controls(self):
        for candidate in (
            {"metadata": {"title": "Different metadata"}},
            {"color": (0.99, 0, 0)},
        ):
            with self.subTest(candidate=candidate):
                self.assertEqual(self.variants(right=candidate)["findings"], [])

    def test_old_measurements_still_compare(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "input.pdf"
            self.pdf(path)
            current = measure(path)
            old = deepcopy(current)
            old["pages"][0].pop("rgbSketch")
            old["pages"][0].pop("rgbSketchSize")
            result = compare(old, current)
            self.assertEqual(result["findings"], [])
            self.assertEqual(result["pages"][0]["mode"], "legacy-grayscale")
            self.assertEqual(current["pages"][0]["rgbSketchSize"], list(RGB_SIZE))

    def test_divergence_excerpt_is_bounded_and_marked_untrusted(self):
        result = self.variants(right={"text": "a" * 125, "x": 1})
        excerpts = [f for f in result["findings"] if f.get("excerpt")]
        self.assertTrue(excerpts)
        self.assertTrue(all(len(f["excerpt"]) <= 120 for f in excerpts))
        self.assertTrue(all(f["untrusted"] for f in excerpts))

    def test_missing_header_and_table_are_visible(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.pdf(root / "a.pdf")
            self.pdf(root / "b.pdf", header=False, table=False)
            result = compare(measure(root / "a.pdf"), measure(root / "b.pdf"))
            self.assertGreater(result["movement"]["missingWordCount"], 0)
            self.assertTrue(any(f["kind"] == "visual-region" for f in result["findings"]))
            detail = evidence(root / "a.pdf", root / "b.pdf", root / "evidence", [1])
            self.assertGreater(detail["pages"][0]["changedPixels"], 0)
            self.assertTrue((root / "evidence/1-crop.png").exists())

    def test_page_count_difference_is_separate(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.pdf(root / "a.pdf", pages=2)
            self.pdf(root / "b.pdf")
            result = compare(measure(root / "a.pdf"), measure(root / "b.pdf"))
            self.assertTrue(any(f["kind"] == "pagination" for f in result["findings"]))

    def test_large_custom_slide_measures_at_a_lower_resolution(self):
        # ECMA-376 allows slides up to 56 inches on a side (ST_SlideSizeCoordinate): 4032 points.
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            with pymupdf.open() as doc:
                page = doc.new_page(width=4032, height=4032)
                page.insert_text((72, 95), "Synthetic large slide")
                doc.save(root / "large.pdf")
            value = measure(root / "large.pdf")
            self.assertEqual(value["pages"][0]["size"], [4032, 4032])
            self.assertEqual(compare(value, value)["findings"], [])
            detail = evidence(root / "large.pdf", root / "large.pdf", root / "evidence", [1])
            self.assertLess(detail["pages"][0]["dpi"][0], 144)
            self.assertEqual(detail["pages"][0]["changedPixels"], 0)

    def test_corrupt_pdf_and_unknown_protocol_refuse(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "bad.pdf"
            path.write_bytes(b"broken")
            with self.assertRaises(pymupdf.FileDataError):
                measure(path)
        with self.assertRaises(ValueError):
            compare({"protocol": 99}, {"protocol": 1})


class PagePreviewTests(unittest.TestCase):
    def test_rotated_page_coordinates_and_missing_page(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            with pymupdf.open() as document:
                page = document.new_page(width=200, height=300)
                page.insert_text((20, 30), "Synthetic page")
                page.set_rotation(90)
                document.save(root / "page.pdf")
            result = preview_page(root / "page.pdf", root / "preview", 1)
            self.assertEqual((result["widthPt"], result["heightPt"]), (300, 200))
            self.assertEqual(result["dpi"], 144)
            self.assertEqual(result["rotationMatrix"], [0, 1, -1, 0, 300, 0])
            self.assertTrue((root / "preview/page.png").is_file())
            for number in [0, 2, 1001]:
                with self.assertRaises(ValueError):
                    preview_page(root / "page.pdf", root / "preview", number)


if __name__ == "__main__":
    unittest.main()
