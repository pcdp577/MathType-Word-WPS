#!/usr/bin/env python3
"""Verify that MathType-Word/WPS is distributable without another skill repo."""

from __future__ import annotations

import ast
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
REQUIRED = [
    ROOT / "SKILL.md",
    ROOT / "requirements.txt",
    ROOT / "references" / "formula-set-integrity.md",
    ROOT / "schemas" / "formula_set_manifest.schema.json",
    ROOT / "scripts" / "audit_formula_set.py",
    ROOT / "scripts" / "mathtype_word_wps.py",
    ROOT / "scripts" / "test_formula_set_audit.py",
]
FORBIDDEN_IMPORT_PREFIXES = {
    "research_paper_writing",
    "paper_review_audit",
    "evidence_grounded_manuscript_skills",
}


def imported_modules(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
    modules: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            modules.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            modules.add(node.module)
    return modules


missing = [str(path.relative_to(ROOT)) for path in REQUIRED if not path.is_file()]
assert not missing, f"missing standalone resources: {missing}"

external_links = []
for path in ROOT.rglob("*"):
    if path.is_symlink() and ROOT not in path.resolve().parents:
        external_links.append(str(path.relative_to(ROOT)))
assert not external_links, f"standalone package contains external symlinks: {external_links}"

for path in ROOT.rglob("*.py"):
    if "__pycache__" in path.parts:
        continue
    for module in imported_modules(path):
        normalized = module.lower().replace("-", "_")
        assert not any(
            normalized == prefix or normalized.startswith(prefix + ".")
            for prefix in FORBIDDEN_IMPORT_PREFIXES
        ), f"cross-skill runtime import in {path.relative_to(ROOT)}: {module}"

sys.path.insert(0, str(ROOT / "scripts"))
from audit_formula_set import audit_manifest  # noqa: E402

probe = {
    "schema_version": "1.0",
    "set_id": "standalone-probe",
    "external_symbols": ["x"],
    "final_outputs": ["y"],
    "formulas": [
        {
            "id": "F01",
            "order": 1,
            "purpose": "Verify the bundled graph auditor.",
            "defines": ["y"],
            "uses": ["x"],
        }
    ],
}
report = audit_manifest(probe)
assert report["status"] == "PASS", report

print("STANDALONE INSTALL TEST PASSED: MathType-Word-WPS")
