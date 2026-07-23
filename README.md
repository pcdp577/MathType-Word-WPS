# MathType-Word/WPS

[中文说明](README.zh-CN.md) | English

With multimodal model assistance for image-to-formula transcription, MathType-Word/WPS provides an end-to-end workflow from formula images to editable MathType equations, automatic Word/WPS document insertion, adaptive manuscript layout, and connected formula-set integrity auditing.

This skill is designed for manuscript production where formulas must remain editable, visually consistent, and publication-ready instead of being pasted as flat screenshots. It automates the fussy parts of the workflow: opening Word/WPS in the background, inserting a real MathType equation object, keeping the equation number as normal right-aligned text, setting the MathType internal formula character size and Times New Roman family from the manuscript body font, fitting the OLE frame around that formula without shrinking complex formulas below their natural MathType height, and checking the saved document so you can tell whether the result is a true editable MathType object or only an image/OMML fallback. The primary target is `.docx`; legacy `.doc` files are supported through the same Word/WPS automation path after opening or converting them with Word/WPS when XML-level inspection is needed.

## Standalone By Design

This repository is a complete standalone skill. It includes its own formula-set schema, graph auditor, tests, MathType OLE implementation, package inspector, and acceptance guidance. It does not import or require `evidence-grounded-manuscript-skills`, `research-paper-writing`, or `paper-review-audit`.

Other manuscript-writing or review skills may consume the same manifest format as an optional interoperability layer, but they are not installation or runtime dependencies. A user who installs only MathType-Word/WPS can still perform formula-set auditing, editable equation insertion, DOCX package inspection, and rendered verification, subject only to the Windows/Word-or-WPS/MathType requirements below.

It supports:

- creating real editable MathType OLE equations (`Equation.DSMT4`, the MathType equation object type that Word/WPS can reopen and edit with MathType)
- using WPS or Word COM backends
- keeping equation numbers as normal right-aligned document text
- setting MathType internal formula character size and Times New Roman family against document body font size, then fitting the OLE frame without unintended downscaling
- inspecting DOCX XML, MathType native OLE streams, and cached WMF/EMF previews for OLE/image/OMML correctness
- validating paragraph placement before batch edits and rebuilding formulas through isolated one-paragraph assets when Word/WPS COM indexing is unstable
- auditing formula-set manifests as directed computation graphs, including undefined or late symbols, orphan outputs, duplicate definitions, instance contracts, deprecated symbols, and removable-alias candidates
- maintaining source-to-equation, prose, code, figure, OLE-part, and preview-part parity across a manuscript revision
- exporting/inserting WMF/EMF formula images as an explicit fallback

## Where It Fits In Image-To-Editable-Formula Workflows

This project does not perform formula OCR from screenshots by itself. Its job starts after a formula image has been recognized or transcribed into MathML or LaTeX.

Typical workflow:

1. Start from a formula screenshot, scanned formula, or formula image already embedded in a manuscript.
2. Use an external recognizer, OCR service, LLM-assisted transcription, or manual editing to produce MathML or LaTeX.
3. Use this tool to replace the image-only formula with a real MathType OLE object (`Equation.DSMT4`) in Word/WPS.
4. Let the tool keep the equation number as normal text, match the formula size to the manuscript body text, and inspect the saved DOCX to confirm that the result is editable OLE rather than another image.

In short: `image -> external recognition/transcription -> MathML/LaTeX -> this tool -> editable MathType object in Word/WPS`.

## Requirements

- Windows desktop session
- Python 3.10+
- Microsoft Word or WPS Office
- An activated/licensed MathType installation (MathType 7.0 recommended); a standard activated install should register the editable `Equation.DSMT4` OLE/COM object automatically for real OLE insertion

Install Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

`latex2mathml` is optional and is only needed for `--latex` input.

```powershell
python -m pip install latex2mathml
```

## Quick Check

```powershell
python .\scripts\mathtype_word_wps.py check-env
```

For real COM probing:

```powershell
python .\scripts\mathtype_word_wps.py check-env --probe-com
```

To probe MathType runtime/activation:

```powershell
python .\scripts\mathtype_word_wps.py check-env --probe-mathtype
```

If MathType is not found, set:

```powershell
$env:MATHTYPE_EXE = "C:\Path\To\MathType.exe"
```

or pass `--mathtype-exe`.

If `Equation.DSMT4` is missing, Word/WPS cannot create the editable MathType object. In that case, activate/license MathType first; if MathType is already activated, repair or reinstall MathType so Windows registers its OLE/COM object again.

## Insert An Editable Formula

When editability or preview glyph fidelity is in doubt, first convert the source MathML to native MathType EF/MTEF. This avoids the fragile "open MathType editor and paste MathML" path and gives you a native MathType payload to validate:

```powershell
python .\scripts\mathtype_word_wps.py native-bridge-mtef `
  --mathml-file ".\formula11.mathml" `
  --output-mtef ".\formula11.mtef"
```

The native bridge performs one bounded MathType DataObject conversion and then cleans up the MathType embedding server. Do not run repeated GUI macro probes or brute-force `Application.Run` loops.

```powershell
python .\scripts\mathtype_word_wps.py replace-docx-ole `
  --docx ".\manuscript.docx" `
  --paragraph-index 49 `
  --mathml-file ".\formula4.mathml" `
  --number "(4)" `
  --backend auto `
  --formula-font-scale 0.8 `
  --formula-lines 1 `
  --backup
```

By default, `--backend auto` tries the full insertion workflow with WPS first and then Word if WPS fails. Use `--backend word`, `--backend wps`, or `--com-progid ...` to force one backend and disable fallback.

Word and WPS are always opened hidden through COM automation. The tool should not show the Word/WPS UI during normal processing.

After insertion, the tool cleans up MathType/MathTypeLib plus safe WPS preview/embedding helper processes and waits until the target document is writable again. This matters because a formula can be visually correct and editable while a stale OLE preview helper still blocks normal typing in the foreground editor.

Sizing rule: `--formula-font-scale 0.8` controls the internal MathType main-character size, not just the external OLE box. The default font family is `Times New Roman`. For complex formulas with fractions, large sums, or multi-row alignment, the OLE frame is treated as a minimum container and is not shrunk below MathType's natural inserted height unless `--allow-downscale-ole` is explicitly passed. MathML is normalized for MathType/Word preview stability and serialized as ASCII numeric entities before paste: for example, scalar middle-dot multiplication is converted from U+00B7 to the dot operator U+22C5, and kernel-size or shape multiplication signs can be converted to ASCII `x` so the visible Word/WPS preview does not become `?` even when the embedded MathType object is editable.

## Inspect

```powershell
python .\scripts\mathtype_word_wps.py inspect-docx `
  --docx ".\manuscript.docx" `
  --paragraph-index 49
```

Expected for real MathType OLE:

- `ole_objects=1`
- `ole_progids=['Equation.DSMT4']`
- `blips=0`
- `omath=0`
- `Equation Native` present in `ole_analysis`
- no suspicious `preview_cache_question_marks` in cached WMF/EMF previews

For a manuscript-wide offline audit that does not open Word or MathType:

```powershell
python .\scripts\mathtype_word_wps.py audit-docx-formulas `
  --docx ".\manuscript.docx" `
  --output-json ".\formula_audit.json" `
  --include-details `
  --min-ole-width-pt 18 `
  --min-ole-height-pt 8
```

The size thresholds are heuristic blank-object warnings. Set either threshold to `0` for intentionally tiny one-symbol equations, or raise it when generated assets have a known minimum width.

Before a batch replacement, test one formula in a scratch copy and verify that the OLE lands beside the intended equation number. OpenXML body indices and Word/WPS COM paragraph indices can differ. If the probe lands in the wrong paragraph, generate each equation in an isolated one-paragraph DOCX, validate it, and transplant the complete OLE relationship, native object, and cached preview into exact OpenXML markers. Keep a formula-number-to-source ledger and audit the assembled candidate before replacing the live manuscript.

Important distinction: an editable MathType OLE shell and the document-visible preview are different layers. If the document shows `?` but the MathType editor opens normally, regenerate the cached preview or reinsert from normalized native MTEF. If MathType opens but the formula behaves as one uneditable pasted object, regenerate the formula from native MTEF or rebuild it in MathType; resizing the OLE frame cannot fix that. Treat `preview_cache_question_marks` as a warning, not a final verdict: byte-level WMF/EMF scans can flag harmless bytes, so final acceptance for this issue should include a Word/PDF rendered screenshot of the formula pages.

If the user cannot type in the manuscript after a run, immediately run `cleanup-leftovers`, then confirm the `.docx` is free by opening it with exclusive read/write access. Do not report the formula task as complete while hidden OLE, MathType, Word, or WPS preview helpers are still affecting foreground editing.

If the user cannot type in any text box, including browsers or Codex itself, stop treating it as a document-lock problem. This is usually a Windows text-service or IME hook failure involving `ctfmon`, `TextInputHost`, a third-party IME such as Tencent WeType, or a hidden Office OLE server such as `VISIO.EXE /Automation -Embedding` pulling the IME into its process tree. Run:

```powershell
python .\scripts\mathtype_word_wps.py recover-text-input --stop-wetype
```

This restarts the Windows text-input chain and, when requested, stops Tencent WeType processes.

## Audit A Connected Formula Set

When several equations form one method or derivation, create a JSON manifest conforming to `schemas/formula_set_manifest.schema.json`. Record each formula's stable ID, order, scientific purpose, defined and used symbols, source file, and optional code, prose, figure, and OLE artifact anchors.

Run the audit before insertion:

```powershell
python .\scripts\audit_formula_set.py ".\formula_set.json" `
  --require-source `
  --require-prose-anchor `
  --output-json ".\formula_set_audit.json"
```

After DOCX assembly, add equation numbers, paragraph anchors, OLE package parts, and preview parts, then re-run with `--require-artifact`.

The audit treats equations as a directed computation graph. It reports undefined and late-defined symbols, orphan outputs, duplicate definitions, dependency cycles, incomplete shared-interface instances, deprecated symbols, and one-consumer alias candidates. Alias findings are suggestions only: retain a quantity when it has independent physical meaning, performs a unit or coordinate conversion, participates in a proof or limiting case, or belongs to the shared interface.

See `references/formula-set-integrity.md` for the semantic, artifact, and rendered acceptance gates. A semantic PASS does not prove that the MathType object is correct, and a valid OLE shell does not prove that the equation set is scientifically closed.

## Use With Cursor, Claude Code, Or Other Agents

The repository is published as a Codex skill, but the actual implementation is a normal Windows Python CLI. Other coding agents can use it as long as they can run PowerShell commands on the same Windows desktop where Word/WPS and MathType are installed.

Notes for Cursor, Claude Code, and similar clients:

- Use the CLI in `scripts/mathtype_word_wps.py` directly; Codex-style automatic skill discovery is not available unless that client implements its own skill loader.
- Copy the relevant workflow rules from `SKILL.md` into the client's project rules or prompt instructions, especially the OLE-first workflow, hidden Word/WPS launch, formula sizing, and XML inspection steps.
- Run from a real Windows desktop session, not WSL, a remote Linux shell, or a headless container. MathType OLE insertion depends on Windows COM automation.
- Keep the target document closed in other Word/WPS processes when doing offline replacements, or reuse the same open application session carefully; duplicate opens can make the file read-only.
- Start with `check-env --probe-com --probe-mathtype` before editing a manuscript, because missing MathType activation or missing `Equation.DSMT4` registration will prevent real editable OLE insertion.

Only the Codex workflow has been exercised end-to-end in this repository so far. Cursor and Claude Code should work through the CLI path, but users should treat them as compatible-by-design rather than fully validated clients until they run the environment check and a small formula replacement test.

## Install As A Codex Skill

Clone or copy this folder into your Codex skills directory:

```powershell
git clone https://github.com/pcdp577/MathType-Word-WPS.git "$env:USERPROFILE\.codex\skills\MathType-Word-WPS"
```

Restart Codex after installation.

## Development Validation

```powershell
python -X utf8 .\scripts\test_formula_set_audit.py
python -X utf8 .\scripts\test_standalone_install.py
python -X utf8 -m py_compile .\scripts\audit_formula_set.py .\scripts\mathtype_word_wps.py
python -X utf8 <skill-creator>\scripts\quick_validate.py .
```
