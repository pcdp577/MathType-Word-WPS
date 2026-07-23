#!/usr/bin/env python3
from audit_formula_set import audit_manifest


valid = {
    "schema_version": "1.0",
    "set_id": "shared-method",
    "external_symbols": ["I0", "geometry", "state", "residual"],
    "final_outputs": ["Y"],
    "retained_symbols": ["D", "C", "B"],
    "deprecated_symbols": ["Gamma"],
    "formulas": [
        {
            "id": "F01",
            "order": 1,
            "purpose": "Map geometry and state to shared exposure quantities.",
            "defines": ["D", "C", "B"],
            "uses": ["I0", "geometry", "state"],
            "source": "F01.tex"
        },
        {
            "id": "F02",
            "order": 2,
            "purpose": "Compose the base exposure with a bounded residual.",
            "defines": ["Y"],
            "uses": ["B", "residual"],
            "source": "F02.tex"
        }
    ],
    "instance_contracts": [
        {
            "name": "surface mapping",
            "required_outputs": ["D", "C", "B"],
            "instances": [
                {
                    "id": "instance-a",
                    "formula_ids": ["F01"],
                    "outputs": ["D", "C", "B"]
                }
            ]
        }
    ]
}

report = audit_manifest(valid, require_source=True)
assert report["status"] == "PASS", report

invalid = {
    "schema_version": "1.0",
    "set_id": "broken-method",
    "external_symbols": ["I0"],
    "final_outputs": ["Y"],
    "deprecated_symbols": ["Gamma"],
    "formulas": [
        {
            "id": "F01",
            "order": 1,
            "purpose": "Uses an undefined and a late-defined symbol.",
            "defines": ["A"],
            "uses": ["missing", "B", "Gamma"]
        },
        {
            "id": "F02",
            "order": 2,
            "purpose": "Defines the late symbol but leaves an orphan.",
            "defines": ["B", "orphan"],
            "uses": ["I0"]
        }
    ],
    "instance_contracts": [
        {
            "name": "shared interface",
            "required_outputs": ["A", "B"],
            "instances": [
                {
                    "id": "instance-b",
                    "formula_ids": ["F02"],
                    "outputs": ["B"]
                }
            ]
        }
    ]
}

bad_report = audit_manifest(invalid)
codes = {item["code"] for item in bad_report["issues"]}
assert bad_report["status"] == "FAIL", bad_report
assert {"FORM-UNDEF", "FORM-ORDER", "FORM-DEPRECATED", "FORM-INSTANCE"} <= codes, codes

print("FORMULA SET AUDIT TESTS PASSED")
