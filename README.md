# MathType-Word/WPS

[中文说明](README.zh-CN.md) | English

With multimodal model assistance for image-to-formula transcription, MathType-Word/WPS provides an end-to-end workflow from formula images to editable MathType equations, automatic Word/WPS document insertion, and adaptive manuscript layout.

This skill is designed for manuscript production where formulas must remain editable, visually consistent, and publication-ready instead of being pasted as flat screenshots. It automates the fussy parts of the workflow: opening Word/WPS in the background, inserting a real MathType equation object, keeping the equation number as normal right-aligned text, setting the MathType internal formula character size from the manuscript body font, fitting the OLE frame around that formula, and checking the saved document so you can tell whether the result is a true editable MathType object or only an image/OMML fallback. The primary target is `.docx`; legacy `.doc` files are supported through the same Word/WPS automation path after opening or converting them with Word/WPS when XML-level inspection is needed.

It supports:

- creating real editable MathType OLE equations (`Equation.DSMT4`, the MathType equation object type that Word/WPS can reopen and edit with MathType)
- using WPS or Word COM backends
- keeping equation numbers as normal right-aligned document text
- setting MathType internal formula character size against document body font size, then fitting the OLE frame
- inspecting DOCX XML for OLE/image/OMML correctness
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
git clone https://github.com/liuyifan577/MathType-Word-WPS.git "$env:USERPROFILE\.codex\skills\MathType-Word-WPS"
```

Restart Codex after installation.
