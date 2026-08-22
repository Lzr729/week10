#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


required = [
    "README.md",
    "reports/week10_stage09_week10_closure_summary_v1.0.xlsx",
    "reports/week10_stage08_eight_company_descriptive_analysis_and_research_profiles_v1.0.xlsx",
    "data/week8_standard_dataset_v1.0.json",
    "data/week9_stage04c_variable_availability_frozen_v1.0.json",
    "inputs/week10_stage06_original_disclosure_ownership_recovery_and_stop_decision_v1.0.xlsx",
    "inputs/week10_stage07_variable_utilization_academic_commercial_value_matrix_v1.0.xlsx",
    "scripts/build_week10_closure.mjs",
    "scripts/verify_release.py",
    "metadata/release_manifest.json",
    "metadata/independent_audit.json",
    "CHECKSUMS.sha256",
]
for relative in required:
    if not (ROOT / relative).is_file():
        fail(f"missing required file: {relative}")

checksum_lines = (ROOT / "CHECKSUMS.sha256").read_text(encoding="utf-8").splitlines()
for line in checksum_lines:
    if not line.strip():
        continue
    expected, relative = line.split("  ", 1)
    target = ROOT / relative
    if not target.is_file():
        fail(f"checksum target missing: {relative}")
    actual = sha256(target)
    if actual != expected:
        fail(f"checksum mismatch: {relative}")

week8 = json.loads((ROOT / "data/week8_standard_dataset_v1.0.json").read_text(encoding="utf-8"))
week9 = json.loads((ROOT / "data/week9_stage04c_variable_availability_frozen_v1.0.json").read_text(encoding="utf-8"))
if len(week8["tables"]["companies"]) != 8:
    fail("Week 8 company count is not 8")
if len(week9["variable_summary"]) != 32:
    fail("Week 9 variable count is not 32")
if len(week9["result_rows"]) != 256:
    fail("Week 9 company-variable row count is not 256")

statuses = {"VALID": 0, "STRUCTURAL_NA": 0, "NOT_COMPUTABLE": 0}
for row in week9["result_rows"]:
    status = row["result_status"]
    if status in statuses:
        statuses[status] += 1
if statuses != {"VALID": 215, "STRUCTURAL_NA": 38, "NOT_COMPUTABLE": 3}:
    fail(f"unexpected Week 9 status counts: {statuses}")

manifest = json.loads((ROOT / "metadata/release_manifest.json").read_text(encoding="utf-8"))
if manifest.get("scope", {}).get("ninth_company_included") is not False:
    fail("manifest must record that the ninth company is not included")

print("PASS")
