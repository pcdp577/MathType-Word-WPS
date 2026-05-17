---
name: "MathType-Word/WPS"
description: "Use when a manuscript formula must be inserted, replaced, inspected, or resized in a DOCX document as an editable MathType OLE equation through Microsoft Word or WPS, with optional WMF/EMF fallback and body-font-matched formula sizing."
---

# MathType-Word/WPS

Use this skill when a `.docx` manuscript needs editable MathType formulas rather than screenshots or OMML-only equations. The stable path is:

`formula image -> external MathML/LaTeX transcription -> native MathType EF/MTEF -> MathType OLE object -> centered formula paragraph -> right-aligned text equation number -> native-stream + preview-cache inspection`

For image-to-editable-formula work, this skill does not do formula OCR. First obtain MathML or LaTeX from an external recognizer, OCR service, LLM-assisted transcription, or manual cleanup. Then use this skill to replace the image-only formula with an editable MathType OLE object in Word/WPS.

## Core Rules

- Prefer real MathType OLE objects (`Equation.DSMT4`) when the user wants editable equations.
- Treat `Equation.DSMT4` as necessary but not sufficient. A successful formula must also have a real MathType native stream (`Equation Native` / `MathType EF` / MTEF) and a sane cached WMF/EMF preview.
- Do not validate only the OLE shell. This can miss two distinct failures: the document preview shows `?` while the internal MathType content is normal, or MathType opens but the content is one uneditable pasted object.
- Prefer the native bridge (`native-bridge-mtef`) for final formula generation or diagnosis. The older "open embedded MathType editor and paste MathML" path is legacy and should not be used as final proof when editability or preview glyphs are in doubt.
- Put the equation number, e.g. `(1)`, in the Word/WPS paragraph as normal text; do not include it inside MathType.
- Default layout is one paragraph: center tab -> formula object -> right tab -> number.
- Always inspect after insertion: require `ole_objects=1`, `ole_progids=['Equation.DSMT4']`, `blips=0`, `omath=0`, expected text number, sane tab stops, `Equation Native` present, no XML-only native stream, and no suspicious cached preview glyphs.
- Do not treat insertion as complete until the MathType internal main-character size and font family are normalized against manuscript body text. Default: `formula_font_pt = body_font_pt * 0.8` and `formula_font_family = Times New Roman`; the script injects these into MathML as `mstyle mathsize="...pt" mathfamily="Times New Roman" fontfamily="Times New Roman"` before pasting into MathType.
- The OLE frame is only a container fit, not the primary font-size control. Never shrink a complex formula below MathType's natural inserted height unless the user explicitly allows downscaling with `--allow-downscale-ole`; otherwise fractions, large sums, and multi-row formulas can look falsely smaller even when the internal formula font was set correctly.
- Before pasting into MathType, normalize preview-unstable MathML operators and then serialize MathML to ASCII numeric entities for clipboard formats. In practice, MathType/Word preview caches can display U+00B7 middle dot and U+00D7 multiplication sign as `?` even when the embedded MathType formula opens correctly; the script normalizes these to preview-stable forms (`U+22C5` dot operator for scalar multiplication, ASCII `x` for kernel-size or shape text) before MTEF/OLE insertion.
- Treat cached-preview byte scans as risk indicators, not final visual proof. If `preview_cache_question_marks` remains but the rendered Word/PDF screenshot is clean, report it as a byte-level false positive; if the rendered screenshot still shows `?`, regenerate from normalized MathML/MTEF and re-render.
- Word/WPS must be opened in hidden/background COM mode. Do not show Word/WPS UI during normal processing; `--visible-app` is deprecated and ignored by the script.
- After every OLE insertion, restore the user's editing environment before reporting success: close the hidden Word/WPS COM document, clean MathType/MathTypeLib leftovers, clean safe WPS `/Preview` / `-Embedding` helper processes, and verify the target `.docx` can be opened with exclusive read/write access. A visually correct formula is not a complete success if the user cannot type in the manuscript afterward.
- If the user reports that no text box works anywhere, including browsers or Codex itself, stop diagnosing the manuscript. That is a Windows input-method/text-services failure, usually `ctfmon`, `TextInputHost`, a third-party IME hook such as Tencent WeType, or a hidden Office OLE server such as `VISIO.EXE /Automation -Embedding` that has pulled an IME into its process tree. Run `recover-text-input`; add `--stop-wetype` when WeType processes are present or suspected. Keep the default `--stop-office-embedding` unless the user is intentionally running hidden Office automation.
- Never run unbounded MathType GUI probing, macro brute-force loops, or repeated `Application.Run` attempts. Any MathType probe must be one formula / one attempt / bounded timeout / process cleanup / post-run process verification.
- If MathType is missing, not registered, or appears unlicensed/not activated, stop and tell the user to install, repair, register, or activate MathType before retrying.
- Keep a backup before modifying a live manuscript.
- Use `cleanup-leftovers` after any failed or suspicious run. It should stop MathType/MathTypeLib and safe WPS preview/embedding helpers, but it must not kill ordinary foreground document-editing sessions.

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

When editability or preview fidelity is suspect, first generate native MTEF for the formula source. This uses MathType's native DataObject and cleans up the embedding server after the bounded single attempt:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" native-bridge-mtef `
  --mathml-file "<formula.mathml>" `
  --output-mtef "<formula.mtef>"
```

If this fails, do not keep trying GUI paste variants. Fix MathType activation/registration, simplify or correct the source MathML, or fall back to a manually checked MathType editor insertion.

Use `--backend auto` by default. In auto mode the script tries the full insertion workflow with WPS first and then Word if WPS fails. Use `--backend word`, `--backend wps`, or `--com-progid ...` only when the user explicitly requests one backend; explicit backend selection disables automatic fallback.

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" replace-docx-ole `
  --docx "<manuscript.docx>" `
  --paragraph-index 49 `
  --mathml-file "<formula.mathml>" `
  --number "(4)" `
  --backend auto `
  --formula-font-scale 0.8 `
  --formula-font-family "Times New Roman" `
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
- `--formula-font-scale 0.75`: set MathType internal formula characters to 75% of body text instead of the default 80%, then fit the OLE frame around that formula.
- `--formula-font-family "Times New Roman"`: set the MathType formula family before paste; keep this for manuscripts that require New Roman formulas.
- `--formula-lines 2`: force two-line height for aligned formulas.
- `--height-pt 20`: fixed object height override.
- `--allow-downscale-ole`: allow shrinking the OLE frame below MathType's natural inserted height; avoid this for final manuscripts unless you intentionally want smaller displayed formulas.
- `--ole-class-type Equation.DSMT4`: override only if the local MathType registration differs.

## Resize Existing MathType OLE

Use this only when a formula is already an OLE object and the frame fit is wrong. This edits the VML/OLE frame only and does not reopen MathType, so it cannot change MathType internal character size, font family, or mojibake/question-mark content. To change the actual formula character size or New Roman font family, rerun `replace-docx-ole` from the source MathML/LaTeX.

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
- `ole_analysis` with `Equation Native` stream presence and native stream markers
- `preview_analysis` with cached WMF/EMF target and question-mark risk
- `ole_shape_width_pt`
- `ole_shape_height_pt`
- `body_font_pt`, `target_formula_font_pt`, `mathml_font_size_pt`, `resolved_height_pt` when a sizing operation was run

For a manuscript-wide offline audit that does not open Word or MathType:

```powershell
python "${SKILL_DIR}\scripts\mathtype_word_wps.py" audit-docx-formulas `
  --docx "<manuscript.docx>" `
  --output-json "<formula_audit.json>" `
  --include-details
```

Interpretation:

- `preview_cache_question_marks` means the visible document preview may be bad even if the MathType editor opens normally. Regenerate the cached preview by reinserting from native MTEF or by a controlled editor update.
- `missing_equation_native_stream`, `native_stream_contains_mathml_xml`, or "opens as one uneditable object" means the formula must be regenerated from MathML/LaTeX through the native bridge or manually rebuilt in MathType; resizing the OLE frame will not fix it.
- If byte-level preview audit is clean but the user still sees `?`, do a bounded visual preview/render check. WMF font fallback can fail without literal `?` bytes in the file.

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
