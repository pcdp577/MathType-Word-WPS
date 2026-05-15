# MathType-Word/WPS

Codex skill and CLI helper for inserting editable MathType equations into `.docx` manuscripts through WPS or Microsoft Word.

It supports:

- creating real MathType OLE equations (`Equation.DSMT4`)
- using WPS or Word COM backends
- keeping equation numbers as normal right-aligned document text
- resizing formulas against document body font size
- inspecting DOCX XML for OLE/image/OMML correctness
- exporting/inserting WMF/EMF formula images as an explicit fallback

## Requirements

- Windows desktop session
- Python 3.10+
- Microsoft Word or WPS Office
- MathType installed and registered as `Equation.DSMT4`
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

If MathType is not registered as `Equation.DSMT4`, or the runtime probe fails, activate/license MathType or repair the installation before inserting equations.

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
