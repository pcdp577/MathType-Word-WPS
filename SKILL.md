---
name: "MathType-Word/WPS"
description: "Use when a manuscript formula must be inserted, replaced, inspected, or resized in a DOCX document as an editable MathType OLE equation through Microsoft Word or WPS, with optional WMF/EMF fallback and body-font-matched formula sizing."
---

# MathType-Word/WPS

Use this skill when a `.docx` manuscript needs editable MathType formulas rather than screenshots or OMML-only equations. The stable path is:

`MathML or simple LaTeX -> MathType OLE object -> centered formula paragraph -> right-aligned text equation number -> XML inspection`

For image-to-editable-formula work, this skill does not do formula OCR. First obtain MathML or LaTeX from an external recognizer, OCR service, LLM-assisted transcription, or manual cleanup. Then use this skill to replace the image-only formula with an editable MathType OLE object in Word/WPS.

## Core Rules

- Prefer real MathType OLE objects (`Equation.DSMT4`) when the user wants editable equations.
- Put the equation number, e.g. `(1)`, in the Word/WPS paragraph as normal text; do not include it inside MathType.
- Default layout is one paragraph: center tab -> formula object -> right tab -> number.
- Always inspect after insertion: require `ole_objects=1`, `ole_progids=['Equation.DSMT4']`, `blips=0`, `omath=0`, expected text number, and sane tab stops.
- Do not treat insertion as complete until size is normalized against manuscript body text. Default: `formula_font_pt = body_font_pt * 0.8`; object height is derived from line count and `ole_line_height_factor=2.1`.
- Word/WPS must be opened in hidden/background COM mode. Do not show Word/WPS UI during normal processing; `--visible-app` is deprecated and ignored by the script.
- If MathType is missing, not registered, or appears unlicensed/not activated, stop and tell the user to install, repair, register, or activate MathType before retrying.
- Keep a backup before modifying a live manuscript.
- Use `cleanup-leftovers` after a failed run if MathType stays in the background.

## Environment

Required for OLE insertion:

- Windows desktop session
- Python 3.10+
- `pywin32`
- `lxml`
- Microsoft Word or WPS Office with COM automation available
- MathType installed and registered as `Equation.DSMT4`

Optional:

- `latex2mathml` for simple `--latex` or `--latex-file` input. Prefer hand-authored MathML for complex aligned formulas.

Install Python dependencies:

```powershell
python -m pip install -r "${SKILL_DIR}\requirements.txt"
```

Run an environment check first:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" check-env
```

If MathType is not found automatically, pass `--mathtype-exe "<path-to-MathType.exe>"` or set `MATHTYPE_EXE`. If `ole_class ... registered=False`, MathType is not correctly registered as an OLE equation server. If runtime or activation is uncertain, run:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" check-env --probe-mathtype
```

If this probe fails, remind the user to activate/license MathType or repair the installation.

## Insert Editable MathType OLE

Use `--backend auto` by default. In auto mode the script tries the full insertion workflow with WPS first and then Word if WPS fails. Use `--backend word`, `--backend wps`, or `--com-progid ...` only when the user explicitly requests one backend; explicit backend selection disables automatic fallback.

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" replace-docx-ole `
  --docx "<manuscript.docx>" `
  --paragraph-index 49 `
  --mathml-file "<formula.mathml>" `
  --number "(4)" `
  --backend auto `
  --formula-font-scale 0.8 `
  --formula-lines 1 `
  --backup
```

For Word only:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" replace-docx-ole `
  --docx "<manuscript.docx>" `
  --paragraph-index 49 `
  --mathml-file "<formula.mathml>" `
  --number "(4)" `
  --backend word `
  --backup
```

Useful options:

- `--one-based`: treat paragraph index as Word/WPS one-based instead of OpenXML zero-based.
- `--body-font-pt 12`: override body font detection.
- `--formula-font-scale 0.75`: make formulas 75% of body text instead of the default 80%.
- `--formula-lines 2`: force two-line height for aligned formulas.
- `--height-pt 20`: fixed object height override.
- `--ole-class-type Equation.DSMT4`: override only if the local MathType registration differs.

## Resize Existing MathType OLE

Use this when a formula is already an OLE object but the visual size is wrong. This edits the VML/OLE frame only and does not reopen MathType:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" resize-docx-ole `
  --docx "<manuscript.docx>" `
  --paragraph-index 49 `
  --formula-lines 1 `
  --formula-font-scale 0.8 `
  --backup
```

## Inspect A Formula Paragraph

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" inspect-docx `
  --docx "<manuscript.docx>" `
  --paragraph-index 49
```

Report at least:

- `text`
- `run_tabs`
- `blips`
- `ole_objects`
- `ole_progids`
- `omath`
- `ole_shape_width_pt`
- `ole_shape_height_pt`
- `body_font_pt`, `target_formula_font_pt`, `resolved_height_pt` when a sizing operation was run

## WMF/EMF Fallback

Only use image fallback when the user accepts vector-image formulas.

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" make-wmf `
  --mathml-file "<formula.mathml>" `
  --output-dir "<formula_exports>" `
  --base-name "formula1_mathtype"
```

Then insert:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" replace-docx `
  --docx "<manuscript.docx>" `
  --paragraph-index 49 `
  --image "<formula1_mathtype.wmf>" `
  --number "(4)" `
  --formula-font-scale 0.8 `
  --backup
```

## Failure Handling

- If LaTeX input fails with a missing dependency, install `latex2mathml` or provide MathML directly.
- In `--backend auto`, WPS is attempted first and Word second. If a backend opens a modal dialog, close the dialog; the script will report that backend failure and continue to the next backend when possible.
- If the document opens read-only, close all Word/WPS instances holding the file and retry.
- If MathType executable is missing, `Equation.DSMT4` is not registered, the MathType editor window never appears, or MathType cannot return content to the host document, explicitly warn the user that MathType may be missing, unregistered, expired, or not activated.
- If MathType remains after a failed insertion:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" cleanup-leftovers
```
