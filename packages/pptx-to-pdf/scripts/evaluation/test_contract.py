# Copyright (c) 2026 EigenPal, Inc. All rights reserved.
# Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
# Production use requires a commercial agreement: licensing@eigenpal.com
"""The Python commands of `contract.json`, run as the contract writes them, match its schemas.

`contract.test.ts` checks the Bun commands the same way. Run with the evaluator's Python
environment.
"""

import gzip
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pymupdf
from worker import PROTOCOL

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
CONTRACT = json.loads((HERE / "contract.json").read_text())


def validate(schema, value, at="$"):
    """The JSON Schema subset the contract uses: type, const, enum, required, properties, items."""
    if isinstance(schema, str) or "$ref" in schema:
        ref = schema if isinstance(schema, str) else schema["$ref"]
        return validate(CONTRACT["schemas"][ref.replace("#/schemas/", "")], value, at)
    kind = (
        "null" if value is None
        else "boolean" if isinstance(value, bool)
        else "integer" if isinstance(value, int)
        else "number" if isinstance(value, float)
        else "string" if isinstance(value, str)
        else "array" if isinstance(value, list)
        else "object"
    )
    types = schema.get("type", [])
    types = [types] if isinstance(types, str) else types
    if types and not any(t == kind or (t == "number" and kind == "integer") for t in types):
        return [f"{at}: expected {types}, got {kind}"]
    if value is None:
        return []
    if "const" in schema and schema["const"] != value:
        return [f"{at}: expected {schema['const']!r}"]
    if "enum" in schema and value not in schema["enum"]:
        return [f"{at}: not in enum"]
    errors = []
    if isinstance(value, str):
        if len(value) < schema.get("minLength", 0) or len(value) > schema.get("maxLength", len(value)):
            errors.append(f"{at}: length")
    elif isinstance(value, list):
        if len(value) < schema.get("minItems", 0) or len(value) > schema.get("maxItems", len(value)):
            errors.append(f"{at}: item count")
        for index, item in enumerate(value):
            if "items" in schema:
                errors += validate(schema["items"], item, f"{at}[{index}]")
    elif isinstance(value, dict):
        errors += [f"{at}.{key}: required" for key in schema.get("required", []) if key not in value]
        for key, sub in schema.get("properties", {}).items():
            if key in value:
                errors += validate(sub, value[key], f"{at}.{key}")
    return errors


def run(name, **values):
    """Run a contract command from the checkout root, with this interpreter as `python`."""
    argv = [values.get(arg[1:-1], arg) if arg.startswith("{") else arg for arg in CONTRACT["commands"][name]["argv"]]
    argv[0] = sys.executable
    subprocess.run(argv, cwd=ROOT, check=True, capture_output=True, timeout=120)


def slide_pdf(path, text="Synthetic slide title", y=95):
    with pymupdf.open() as doc:
        page = doc.new_page(width=960, height=540)
        page.insert_text((72, y), text)
        page.draw_rect((300, 300, 340, 340), color=(0, 0, 1), fill=(0, 0, 1))
        doc.save(path)


class ContractTests(unittest.TestCase):
    def test_protocol_matches(self):
        self.assertEqual(CONTRACT["protocol"], PROTOCOL)

    def test_measure_compare_evidence_and_page_match_their_schemas(self):
        commands = CONTRACT["commands"]
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            slide_pdf(root / "reference.pdf")
            slide_pdf(root / "candidate.pdf", y=101)
            run("measure", input=str(root / "reference.pdf"), output=str(root / "reference.json.gz"))
            run("measure", input=str(root / "candidate.pdf"), output=str(root / "candidate.json"))
            measured = json.loads(gzip.decompress((root / "reference.json.gz").read_bytes()))
            self.assertEqual(validate(commands["measure"]["output"], measured), [])
            run(
                "compare",
                reference=str(root / "reference.json.gz"),
                candidate=str(root / "candidate.json"),
                output=str(root / "comparison.json"),
            )
            comparison = json.loads((root / "comparison.json").read_text())
            self.assertEqual(validate(commands["compare"]["output"], comparison), [])
            self.assertEqual(comparison["firstDivergence"]["type"], "text-position")
            run(
                "evidence",
                reference=str(root / "reference.pdf"),
                candidate=str(root / "candidate.pdf"),
                outputDir=str(root / "evidence"),
                pages="1",
            )
            report = json.loads((root / "evidence/report.json").read_text())
            self.assertEqual(validate(commands["evidence"]["output"], report), [])
            for suffix in ("reference", "candidate", "overlay", "crop"):
                self.assertTrue((root / f"evidence/1-{suffix}.png").is_file())
            run("page", input=str(root / "candidate.pdf"), outputDir=str(root / "page"), page="1")
            page = json.loads((root / "page/report.json").read_text())
            self.assertEqual(validate(commands["page"]["output"], page), [])
            self.assertTrue((root / "page/page.png").is_file())

    def test_equal_pdfs_compare_with_no_divergence(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            slide_pdf(root / "a.pdf")
            run("measure", input=str(root / "a.pdf"), output=str(root / "a.json"))
            run("compare", reference=str(root / "a.json"), candidate=str(root / "a.json"), output=str(root / "c.json"))
            comparison = json.loads((root / "c.json").read_text())
            self.assertEqual(validate("#/schemas/comparison", comparison), [])
            self.assertIsNone(comparison["firstDivergence"])

    def test_a_failed_command_exits_nonzero(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "bad.pdf").write_bytes(b"broken")
            with self.assertRaises(subprocess.CalledProcessError):
                run("measure", input=str(root / "bad.pdf"), output=str(root / "bad.json"))
            self.assertFalse((root / "bad.json").exists())


if __name__ == "__main__":
    unittest.main()
