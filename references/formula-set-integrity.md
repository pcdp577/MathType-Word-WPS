# Formula-Set Integrity Workflow

Use this workflow when a manuscript contains a connected method equation set rather than one isolated equation.

## Responsibility Boundary

Treat three layers separately:

1. Semantic layer: the formula graph has complete definitions, dependencies, interfaces, limits, and physical meaning.
2. Artifact layer: each source formula maps to the intended editable equation object, number, paragraph, and preview.
3. Rendered layer: Word/WPS and the exported PDF show the same symbols, numbering, spacing, and glyphs.

This skill owns the artifact and rendered layers. It can mechanically audit a semantic manifest, but it must not invent scientific meaning or delete a variable merely because a script reports a simplification candidate.

## Formula-Set Manifest

Create one JSON manifest conforming to `schemas/formula_set_manifest.schema.json`.

For every displayed formula record:

- stable formula ID and order;
- one-sentence scientific purpose;
- symbols defined and symbols used;
- source LaTeX or MathML file;
- code, prose, and figure anchors when available;
- DOCX equation number, paragraph marker, OLE package part, and preview part after insertion.

At set level record:

- external input symbols;
- final outputs;
- retained symbols with independent physical or interface roles;
- deprecated symbols that must disappear;
- allowed redefinitions with explicit scope;
- instance contracts when several physical or algorithmic instances implement one shared interface;
- symbol inventories recovered from equations, prose, and figures.

## Closure Gates

Run `scripts/audit_formula_set.py` before insertion and again after assembly.

The audit reports:

- `FORM-UNDEF`: symbol is neither defined nor declared external;
- `FORM-ORDER`: a symbol is used before its defining formula;
- `FORM-ORPHAN`: output has no downstream consumer and is not retained or final;
- `FORM-COLLIDE`: duplicate formula IDs, orders, or symbol definitions;
- `FORM-INSTANCE`: an instance omits a required shared-interface output;
- `FORM-DEPRECATED`: a retired symbol remains in the formula set or prose inventory;
- `FORM-FIGURE`: a retired symbol remains in the figure inventory;
- `FORM-CYCLE`: the declared formula dependency graph is cyclic;
- `FORM-ALIAS`: a one-consumer symbol should be inspected as a possible removable alias.

`FORM-ALIAS` is a suggestion, not an automatic deletion rule. Retain a one-use quantity when it has independent physical meaning, carries a unit or coordinate transformation, participates in a limit or proof, or is required by a shared interface.

## Artifact Assembly Gate

Before promoting the candidate manuscript:

1. Verify every manifest source maps to the intended `Equation.DSMT4` object.
2. Verify equation numbers remain normal document text and match manifest order.
3. Verify native MathType streams and cached previews exist.
4. Scan prose and embedded figure text against current and deprecated symbol lists.
5. Export with the target Word/WPS backend and inspect every affected page.
6. Confirm no visible missing glyph, stale symbol, clipping, or source-object mismatch.
7. Promote only after semantic, package, and rendered checks all pass.

Byte-level preview warnings remain warnings until native rendering is checked. A visible wrong glyph or a valid OLE object attached to the wrong source is a hard failure.
