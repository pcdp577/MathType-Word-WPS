# MathType-Word/WPS

Codex skill and CLI helper for inserting editable MathType equations into Word manuscripts through WPS or Microsoft Word.

This skill is designed for manuscript production where formulas must remain editable, visually consistent, and publication-ready instead of being pasted as flat screenshots. It automates the fussy parts of the workflow: opening Word/WPS in the background, inserting a real MathType equation object, keeping the equation number as normal right-aligned text, matching formula size to the surrounding body text, and checking the saved document so you can tell whether the result is a true editable MathType object or only an image/OMML fallback. The primary target is `.docx`; legacy `.doc` files are supported through the same Word/WPS automation path after opening or converting them with Word/WPS when XML-level inspection is needed.

It supports:

- creating real editable MathType OLE equations (`Equation.DSMT4`, the MathType equation object type that Word/WPS can reopen and edit with MathType)
- using WPS or Word COM backends
- keeping equation numbers as normal right-aligned document text
- resizing formulas against document body font size
- inspecting DOCX XML for OLE/image/OMML correctness
- exporting/inserting WMF/EMF formula images as an explicit fallback

## Requirements

- Windows desktop session
- Python 3.10+
- Microsoft Word or WPS Office
- MathType installed normally; a standard activated MathType installation should register the editable `Equation.DSMT4` OLE/COM object automatically
- An activated/licensed MathType installation for real OLE insertion

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

## Install As A Codex Skill

Clone or copy this folder into your Codex skills directory:

```powershell
git clone https://github.com/liuyifan577/MathType-Word-WPS.git "$env:USERPROFILE\.codex\skills\MathType-Word-WPS"
```

Restart Codex after installation.
