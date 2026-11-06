# Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
# Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
# Production use requires a commercial agreement: licensing@eigenpal.com
"""The whole protocol on synthetic fixtures: export with Bun, then measure and compare.

Run with the evaluator's Python environment. Bun must be on the PATH. Only decks under
`fixtures/` are used, never a corpus deck.
"""

import json
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from worker import compare, evidence, measure, preview_page

ROOT = Path(__file__).resolve().parents[4]
EXPORT = ROOT / "packages/pptx-to-pdf/scripts/evaluation/export.ts"
FIXTURES = ROOT / "fixtures"


def export(deck, output):
    """Run the export entry as the evaluator does: its own process, one JSON line on stdout."""
    child = subprocess.run(
        ["bun", str(EXPORT), str(FIXTURES / f"{deck}.pptx"), str(output)],
        cwd=ROOT,
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    lines = child.stdout.strip().splitlines()
    if child.returncode != 0 or len(lines) != 1:
        raise AssertionError(child.stderr or child.stdout)
    return json.loads(lines[0])


@unittest.skipIf(shutil.which("bun") is None, "Bun is not on the PATH")
class FixtureProtocolTests(unittest.TestCase):
    def test_candidate_pages_are_slides_with_measured_words(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            report = export("tables", root / "tables.pdf")
            self.assertEqual(report["status"], "exported")
            value = measure(root / "tables.pdf")
            self.assertEqual(len(value["pages"]), report["pages"])
            self.assertEqual(report["slideNumbers"], [1, 2, 3])
            for page in value["pages"]:
                self.assertEqual(page["size"], [960, 540])
                self.assertEqual(page["rgbSketchSize"], [192, 144])
            # Tables draw: no `frame.table` diagnostic, and the cell text is measured words.
            codes = {d["code"] for d in report["diagnostics"]}
            self.assertNotIn("frame.table", codes)
            words = {word["text"] for word in value["words"]}
            self.assertTrue({"Region", "Q1", "North"} <= words)
            report = export("text-runs", root / "text-runs.pdf")
            words = {word["text"] for word in measure(root / "text-runs.pdf")["words"]}
            self.assertTrue({"quick", "brown", "fox"} <= words)

    def test_the_contract_commands_pair_the_reference_through_the_report(self):
        # PPTX: the reference PDF leaves out the hidden slides 2 and 4.
        reference = ROOT / "packages/pptx-to-pdf/src/__tests__/reference"
        contract = json.loads((ROOT / "packages/pptx-to-pdf/scripts/evaluation/contract.json").read_text())
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            child = subprocess.run(
                ["bun", str(EXPORT), str(reference / "hidden-slides.pptx"), str(root / "cand.pdf")],
                cwd=ROOT,
                capture_output=True,
                text=True,
                timeout=120,
            )
            self.assertEqual(child.returncode, 0, child.stderr or child.stdout)
            report = json.loads(child.stdout.strip().splitlines()[-1])
            self.assertEqual(report["slideNumbers"], [1, 3])
            (root / "report.json").write_text(json.dumps(report))
            values = {
                "reference": str(root / "ref.json"),
                "candidate": str(root / "cand.json"),
                "candidateReport": str(root / "report.json"),
                "output": str(root / "comparison.json"),
                "outputDir": str(root / "evidence"),
                "pages": "1,2",
            }

            def run(name, **more):
                command = contract["commands"][name]
                argv = command["argv"] + [arg for group in command.get("optionalArgv", []) for arg in group]
                argv = [{**values, **more}.get(a[1:-1], a) if a.startswith("{") else a for a in argv]
                argv[0] = sys.executable
                subprocess.run(argv, cwd=ROOT, check=True, capture_output=True, timeout=120)

            run("measure", input=str(reference / "hidden-slides.pdf"), output=values["reference"])
            run("measure", input=str(root / "cand.pdf"), output=values["candidate"])
            run("compare")
            comparison = json.loads((root / "comparison.json").read_text())
            self.assertEqual([p["page"] for p in comparison["pages"]], [1, 2])
            kinds = {f["kind"] for f in comparison["findings"]}
            self.assertNotIn("pagination", kinds)
            self.assertNotIn("text-coverage", kinds)
            # `evidence` reads the PDFs: reference slide 2 is the candidate's page 2 (slide 3).
            run("evidence", reference=str(reference / "hidden-slides.pdf"), candidate=str(root / "cand.pdf"))
            shown = json.loads((root / "evidence" / "report.json").read_text())
            self.assertEqual([(p["page"], p["candidatePage"]) for p in shown["pages"]], [(1, 1), (2, 2)])

    def test_same_candidate_compares_equal_and_another_deck_does_not(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            export("text-runs", root / "a.pdf")
            export("text-runs", root / "b.pdf")
            self.assertEqual((root / "a.pdf").read_bytes(), (root / "b.pdf").read_bytes())
            a = measure(root / "a.pdf")
            same = compare(a, measure(root / "b.pdf"))
            self.assertEqual(same["findings"], [])
            self.assertIsNone(same["firstDivergence"])
            export("tables", root / "c.pdf")
            other = compare(a, measure(root / "c.pdf"))
            kinds = {finding["kind"] for finding in other["findings"]}
            self.assertIn("pagination", kinds)
            self.assertIn("text-coverage", kinds)
            self.assertTrue(other["firstDivergence"]["untrusted"])

    def test_evidence_and_page_preview_render_slides(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            export("preset-shapes", root / "a.pdf")
            export("pictures", root / "b.pdf")
            detail = evidence(root / "a.pdf", root / "b.pdf", root / "evidence", [1])
            self.assertEqual(detail["pages"][0]["dpi"], [144, 144])
            self.assertGreater(detail["pages"][0]["changedPixels"], 0)
            self.assertTrue((root / "evidence/1-overlay.png").is_file())
            page = preview_page(root / "a.pdf", root / "preview", 1)
            self.assertEqual((page["widthPt"], page["heightPt"]), (960, 540))
            self.assertTrue((root / "preview/page.png").is_file())

    def test_a_refused_deck_reports_a_typed_error(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            bad = root / "bad.pptx"
            bad.write_bytes(b"PK\x03\x04 not a zip")
            child = subprocess.run(
                ["bun", str(EXPORT), str(bad), str(root / "bad.pdf")],
                cwd=ROOT,
                capture_output=True,
                text=True,
                timeout=120,
                check=False,
            )
            self.assertEqual(child.returncode, 1)
            report = json.loads(child.stdout)
            self.assertEqual(report["status"], "error")
            self.assertEqual(report["code"], "documentOpenFailed")
            self.assertFalse((root / "bad.pdf").exists())


if __name__ == "__main__":
    unittest.main()
