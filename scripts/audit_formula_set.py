#!/usr/bin/env python3
"""Audit a manuscript formula-set manifest as a directed computation graph."""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path
from typing import Any


def make_issue(
    code: str,
    severity: str,
    message: str,
    *,
    formula_id: str | None = None,
    symbol: str | None = None,
) -> dict[str, Any]:
    issue: dict[str, Any] = {
        "code": code,
        "severity": severity,
        "message": message,
    }
    if formula_id is not None:
        issue["formula_id"] = formula_id
    if symbol is not None:
        issue["symbol"] = symbol
    return issue


def string_set(value: Any, field: str, issues: list[dict[str, Any]]) -> set[str]:
    if value is None:
        return set()
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        issues.append(make_issue("FORM-MANIFEST", "error", f"{field} must be a string array"))
        return set()
    return set(value)


def find_cycles(graph: dict[str, set[str]]) -> list[list[str]]:
    cycles: list[list[str]] = []
    state: dict[str, int] = {}
    stack: list[str] = []

    def visit(node: str) -> None:
        state[node] = 1
        stack.append(node)
        for nxt in sorted(graph.get(node, set())):
            if state.get(nxt, 0) == 0:
                visit(nxt)
            elif state.get(nxt) == 1 and nxt in stack:
                start = stack.index(nxt)
                cycle = stack[start:] + [nxt]
                if cycle not in cycles:
                    cycles.append(cycle)
        stack.pop()
        state[node] = 2

    for node in sorted(graph):
        if state.get(node, 0) == 0:
            visit(node)
    return cycles


def audit_manifest(
    payload: dict[str, Any],
    *,
    require_source: bool = False,
    require_code_anchor: bool = False,
    require_prose_anchor: bool = False,
    require_artifact: bool = False,
) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []
    suggestions: list[dict[str, Any]] = []

    if not isinstance(payload, dict):
        return {
            "status": "FAIL",
            "summary": {"errors": 1, "warnings": 0, "suggestions": 0},
            "issues": [make_issue("FORM-MANIFEST", "error", "manifest root must be an object")],
            "suggestions": [],
        }

    formulas_raw = payload.get("formulas")
    if not isinstance(formulas_raw, list) or not formulas_raw:
        issues.append(make_issue("FORM-MANIFEST", "error", "formulas must be a non-empty array"))
        formulas_raw = []

    external = string_set(payload.get("external_symbols", []), "external_symbols", issues)
    final_outputs = string_set(payload.get("final_outputs", []), "final_outputs", issues)
    retained = string_set(payload.get("retained_symbols", []), "retained_symbols", issues)
    deprecated = string_set(payload.get("deprecated_symbols", []), "deprecated_symbols", issues)
    allowed_redefinitions = string_set(
        payload.get("allowed_redefinitions", []), "allowed_redefinitions", issues
    )

    formulas: list[dict[str, Any]] = []
    formula_ids: set[str] = set()
    orders: set[int] = set()

    for index, item in enumerate(formulas_raw):
        if not isinstance(item, dict):
            issues.append(
                make_issue("FORM-MANIFEST", "error", f"formulas[{index}] must be an object")
            )
            continue
        formula_id = item.get("id")
        order = item.get("order")
        purpose = item.get("purpose")
        defines = item.get("defines")
        uses = item.get("uses")
        if not isinstance(formula_id, str) or not formula_id:
            issues.append(make_issue("FORM-MANIFEST", "error", f"formulas[{index}].id is required"))
            continue
        if formula_id in formula_ids:
            issues.append(
                make_issue(
                    "FORM-COLLIDE",
                    "error",
                    f"duplicate formula id: {formula_id}",
                    formula_id=formula_id,
                )
            )
        formula_ids.add(formula_id)
        if not isinstance(order, int) or order < 0:
            issues.append(
                make_issue(
                    "FORM-MANIFEST",
                    "error",
                    "order must be a non-negative integer",
                    formula_id=formula_id,
                )
            )
            order = index
        elif order in orders:
            issues.append(
                make_issue(
                    "FORM-COLLIDE",
                    "error",
                    f"duplicate formula order: {order}",
                    formula_id=formula_id,
                )
            )
        orders.add(order)
        if not isinstance(purpose, str) or not purpose.strip():
            issues.append(
                make_issue(
                    "FORM-MANIFEST",
                    "error",
                    "purpose must explain the formula's scientific role",
                    formula_id=formula_id,
                )
            )
        if not isinstance(defines, list) or any(not isinstance(x, str) for x in defines):
            issues.append(
                make_issue("FORM-MANIFEST", "error", "defines must be a string array", formula_id=formula_id)
            )
            defines = []
        if not isinstance(uses, list) or any(not isinstance(x, str) for x in uses):
            issues.append(
                make_issue("FORM-MANIFEST", "error", "uses must be a string array", formula_id=formula_id)
            )
            uses = []
        formula = dict(item)
        formula.update({"id": formula_id, "order": order, "defines": defines, "uses": uses})
        formulas.append(formula)

    formulas.sort(key=lambda item: (item["order"], item["id"]))

    definitions: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    all_definitions: dict[str, list[dict[str, Any]]] = defaultdict(list)
    consumers: dict[str, list[dict[str, Any]]] = defaultdict(list)
    formula_by_id = {item["id"]: item for item in formulas}

    for formula in formulas:
        scope = str(formula.get("scope", "global"))
        for symbol in formula["defines"]:
            definitions[(scope, symbol)].append(formula)
            all_definitions[symbol].append(formula)
        for symbol in formula["uses"]:
            consumers[symbol].append(formula)

    for (scope, symbol), producers in sorted(definitions.items()):
        if len(producers) > 1 and symbol not in allowed_redefinitions:
            ids = ", ".join(item["id"] for item in producers)
            issues.append(
                make_issue(
                    "FORM-COLLIDE",
                    "error",
                    f"{symbol} is defined more than once in scope {scope}: {ids}",
                    symbol=symbol,
                )
            )

    for symbol in sorted(final_outputs):
        if symbol not in all_definitions:
            issues.append(
                make_issue(
                    "FORM-UNDEF",
                    "error",
                    f"final output {symbol} is never defined",
                    symbol=symbol,
                )
            )
    for symbol in sorted(retained):
        if symbol not in all_definitions and symbol not in external:
            issues.append(
                make_issue(
                    "FORM-UNDEF",
                    "error",
                    f"retained symbol {symbol} is never defined or declared external",
                    symbol=symbol,
                )
            )

    dependency_graph: dict[str, set[str]] = {item["id"]: set() for item in formulas}
    for formula in formulas:
        for symbol in formula["uses"]:
            if symbol in deprecated:
                issues.append(
                    make_issue(
                        "FORM-DEPRECATED",
                        "error",
                        f"deprecated symbol {symbol} is still used",
                        formula_id=formula["id"],
                        symbol=symbol,
                    )
                )
            producers = all_definitions.get(symbol, [])
            if not producers:
                if symbol not in external:
                    issues.append(
                        make_issue(
                            "FORM-UNDEF",
                            "error",
                            f"{symbol} is used but never defined or declared external",
                            formula_id=formula["id"],
                            symbol=symbol,
                        )
                    )
                continue
            prior = [item for item in producers if item["order"] < formula["order"]]
            if not prior:
                issues.append(
                    make_issue(
                        "FORM-ORDER",
                        "error",
                        f"{symbol} is used before its defining formula",
                        formula_id=formula["id"],
                        symbol=symbol,
                    )
                )
                producer = min(producers, key=lambda item: item["order"])
            else:
                producer = max(prior, key=lambda item: item["order"])
            dependency_graph[producer["id"]].add(formula["id"])

    for formula in formulas:
        for symbol in formula["defines"]:
            if symbol in deprecated:
                issues.append(
                    make_issue(
                        "FORM-DEPRECATED",
                        "error",
                        f"deprecated symbol {symbol} is still defined",
                        formula_id=formula["id"],
                        symbol=symbol,
                    )
                )
            later_consumers = [
                item for item in consumers.get(symbol, []) if item["order"] > formula["order"]
            ]
            if not later_consumers and symbol not in final_outputs and symbol not in retained:
                issues.append(
                    make_issue(
                        "FORM-ORPHAN",
                        "warning",
                        f"{symbol} has no downstream consumer and is not a final or retained output",
                        formula_id=formula["id"],
                        symbol=symbol,
                    )
                )
            if (
                len(later_consumers) == 1
                and symbol not in final_outputs
                and symbol not in retained
            ):
                suggestions.append(
                    make_issue(
                        "FORM-ALIAS",
                        "suggestion",
                        f"{symbol} has one downstream consumer; inspect whether it is a removable alias",
                        formula_id=formula["id"],
                        symbol=symbol,
                    )
                )

    for cycle in find_cycles(dependency_graph):
        issues.append(
            make_issue(
                "FORM-CYCLE",
                "error",
                "formula dependency cycle: " + " -> ".join(cycle),
            )
        )

    for contract in payload.get("instance_contracts", []):
        if not isinstance(contract, dict):
            issues.append(make_issue("FORM-MANIFEST", "error", "instance contract must be an object"))
            continue
        name = str(contract.get("name", "<unnamed>"))
        required = set(contract.get("required_outputs", []))
        for instance in contract.get("instances", []):
            instance_id = str(instance.get("id", "<unnamed>"))
            outputs = set(instance.get("outputs", []))
            missing = sorted(required - outputs)
            if missing:
                issues.append(
                    make_issue(
                        "FORM-INSTANCE",
                        "error",
                        f"instance {instance_id} in {name} misses outputs: {', '.join(missing)}",
                    )
                )
            actual_outputs: set[str] = set()
            for formula_id in instance.get("formula_ids", []):
                if formula_id not in formula_by_id:
                    issues.append(
                        make_issue(
                            "FORM-INSTANCE",
                            "error",
                            f"instance {instance_id} references unknown formula {formula_id}",
                        )
                    )
                else:
                    actual_outputs.update(formula_by_id[formula_id]["defines"])
            undeclared_actual = sorted(outputs - actual_outputs)
            if undeclared_actual:
                issues.append(
                    make_issue(
                        "FORM-INSTANCE",
                        "error",
                        f"instance {instance_id} declares outputs not defined by its formulas: "
                        + ", ".join(undeclared_actual),
                    )
                )

    for surface, symbols in payload.get("symbol_surfaces", {}).items():
        if not isinstance(symbols, list):
            issues.append(
                make_issue("FORM-MANIFEST", "error", f"symbol_surfaces.{surface} must be an array")
            )
            continue
        for symbol in symbols:
            if symbol in deprecated:
                issues.append(
                    make_issue(
                        "FORM-FIGURE" if surface == "figures" else "FORM-DEPRECATED",
                        "error",
                        f"deprecated symbol {symbol} remains in {surface}",
                        symbol=symbol,
                    )
                )

    required_fields = []
    if require_source:
        required_fields.append(("source", "FORM-SOURCE"))
    if require_code_anchor:
        required_fields.append(("code_anchor", "FORM-CODE"))
    if require_prose_anchor:
        required_fields.append(("prose_anchor", "FORM-PROSE"))
    if require_artifact:
        required_fields.append(("artifact", "FORM-OLE"))
    for formula in formulas:
        for field, code in required_fields:
            if not formula.get(field):
                issues.append(
                    make_issue(
                        code,
                        "error",
                        f"required field {field} is missing",
                        formula_id=formula["id"],
                    )
                )

    errors = sum(issue["severity"] == "error" for issue in issues)
    warnings = sum(issue["severity"] == "warning" for issue in issues)
    status = "FAIL" if errors else ("WARN" if warnings or suggestions else "PASS")
    return {
        "schema_version": payload.get("schema_version"),
        "set_id": payload.get("set_id"),
        "status": status,
        "summary": {
            "formulas": len(formulas),
            "symbols_defined": len(all_definitions),
            "errors": errors,
            "warnings": warnings,
            "suggestions": len(suggestions),
        },
        "issues": issues,
        "suggestions": suggestions,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path, help="formula-set manifest JSON")
    parser.add_argument("--output-json", type=Path)
    parser.add_argument("--require-source", action="store_true")
    parser.add_argument("--require-code-anchor", action="store_true")
    parser.add_argument("--require-prose-anchor", action="store_true")
    parser.add_argument("--require-artifact", action="store_true")
    parser.add_argument("--fail-on-warnings", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    payload = json.loads(args.manifest.read_text(encoding="utf-8-sig"))
    report = audit_manifest(
        payload,
        require_source=args.require_source,
        require_code_anchor=args.require_code_anchor,
        require_prose_anchor=args.require_prose_anchor,
        require_artifact=args.require_artifact,
    )
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output_json:
        args.output_json.parent.mkdir(parents=True, exist_ok=True)
        args.output_json.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    if report["status"] == "FAIL":
        return 1
    if args.fail_on_warnings and report["status"] == "WARN":
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
