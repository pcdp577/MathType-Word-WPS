#!/usr/bin/env python3
"""Create MathType formulas and place them in DOCX manuscripts.

The preferred manuscript path creates a real MathType Equation.DSMT4 OLE
object through WPS/Word Insert Object semantics. WMF/EMF insertion remains as
an explicit vector-image fallback. OLE objects can be auto-sized against the
document body font so insertion is not treated as complete until visual scale
has been normalized.
"""

from __future__ import annotations

import argparse
import ctypes
import os
import platform
import re
import shutil
import struct
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
O_NS = "urn:schemas-microsoft-com:office:office"
V_NS = "urn:schemas-microsoft-com:vml"

NS = {
    "w": W_NS,
    "r": R_NS,
    "m": M_NS,
    "wp": WP_NS,
    "a": A_NS,
    "pic": PIC_NS,
    "o": O_NS,
    "v": V_NS,
}
W = f"{{{W_NS}}}"
R = f"{{{R_NS}}}"
WP = f"{{{WP_NS}}}"
A = f"{{{A_NS}}}"
PIC = f"{{{PIC_NS}}}"
REL = f"{{{REL_NS}}}"
CT = f"{{{CT_NS}}}"
O = f"{{{O_NS}}}"
V = f"{{{V_NS}}}"
EMU_PER_PT = 12700
MATHTYPE_EXE_ENV = "MATHTYPE_EXE"
MATHTYPE_EXE_NAME = "MathType.exe"
MATHTYPE_PASTE_COMMAND_ID = 0x1F9
MATHTYPE_SELECT_ALL_COMMAND_ID = 513
MATHTYPE_UPDATE_HOST_COMMAND_ID = 403
MATHTYPE_CLOSE_RETURN_COMMAND_ID = 402
CSS_PT_RE = re.compile(r"(?P<key>width|height)\s*:\s*(?P<value>[-+]?\d+(?:\.\d+)?)pt", re.IGNORECASE)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_backup(path: Path, label: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = path.with_name(f"{path.stem}.backup_before_{label}_{stamp}{path.suffix}")
    shutil.copyfile(path, backup)
    return backup


def latex_to_mathml(text: str) -> str:
    try:
        from latex2mathml.converter import convert
    except ModuleNotFoundError as exc:
        raise RuntimeError("LaTeX input requires the optional package: pip install latex2mathml") from exc

    stripped = text.strip()
    for prefix, suffix in (("$$", "$$"), ("\\[", "\\]"), ("$", "$")):
        if stripped.startswith(prefix) and stripped.endswith(suffix):
            stripped = stripped[len(prefix) : -len(suffix)].strip()
    return convert(stripped, display="block")


def half_points_to_pt(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return float(value) / 2.0
    except ValueError:
        return None


def format_pt(value: float) -> str:
    text = f"{float(value):.2f}".rstrip("0").rstrip(".")
    return text if text else "0"


def parse_css_pt(style: str | None) -> dict[str, float]:
    if not style:
        return {}
    return {match.group("key").lower(): float(match.group("value")) for match in CSS_PT_RE.finditer(style)}


def set_css_pt(style: str | None, key: str, value_pt: float) -> str:
    key = key.lower()
    value = f"{format_pt(value_pt)}pt"
    style = style or ""
    pattern = re.compile(rf"({re.escape(key)}\s*:\s*)[-+]?\d+(?:\.\d+)?pt", re.IGNORECASE)
    if pattern.search(style):
        return pattern.sub(rf"\g<1>{value}", style, count=1)
    style = style.strip()
    if style and not style.endswith(";"):
        style += ";"
    return f"{style}{key}:{value};"


def paragraph_text(p: etree._Element) -> str:
    return "".join(p.xpath(".//w:t/text()", namespaces=NS))


def paragraph_style_id(p: etree._Element) -> str | None:
    style = p.find("w:pPr/w:pStyle", namespaces=NS)
    return style.get(W + "val") if style is not None else None


def run_font_sizes_pt(p: etree._Element) -> list[float]:
    values: list[float] = []
    for node in p.findall(".//w:rPr/w:sz", namespaces=NS):
        pt = half_points_to_pt(node.get(W + "val"))
        if pt is not None:
            values.append(pt)
    return values


def is_body_like_text(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if len(compact) < 8:
        return False
    if re.fullmatch(r"[\(\)\[\]（），,.;:：；\d\s]+", compact):
        return False
    return True


def is_plausible_body_pt(value: float) -> bool:
    return 8.0 <= float(value) <= 14.0


def style_font_size_map(styles_root: etree._Element | None) -> tuple[dict[str, float], float | None]:
    if styles_root is None:
        return {}, None
    direct_sizes: dict[str, float] = {}
    based_on: dict[str, str] = {}
    for style in styles_root.findall(".//w:style", namespaces=NS):
        style_id = style.get(W + "styleId")
        if not style_id:
            continue
        parent = style.find("w:basedOn", namespaces=NS)
        if parent is not None and parent.get(W + "val"):
            based_on[style_id] = parent.get(W + "val")
        size = style.find("w:rPr/w:sz", namespaces=NS)
        pt = half_points_to_pt(size.get(W + "val") if size is not None else None)
        if pt is not None:
            direct_sizes[style_id] = pt

    mapping: dict[str, float] = {}

    def resolve(style_id: str, seen: set[str] | None = None) -> float | None:
        if style_id in mapping:
            return mapping[style_id]
        if style_id in direct_sizes:
            mapping[style_id] = direct_sizes[style_id]
            return mapping[style_id]
        seen = seen or set()
        if style_id in seen:
            return None
        seen.add(style_id)
        parent = based_on.get(style_id)
        if parent:
            value = resolve(parent, seen)
            if value is not None:
                mapping[style_id] = value
                return value
        return None

    for style_id in set(direct_sizes) | set(based_on):
        resolve(style_id)

    default_size = styles_root.find(".//w:docDefaults/w:rPrDefault/w:rPr/w:sz", namespaces=NS)
    default_pt = half_points_to_pt(default_size.get(W + "val") if default_size is not None else None)
    return mapping, default_pt


def detect_body_font_pt(
    docx: Path,
    paragraph_index: int | None = None,
    explicit_pt: float | None = None,
) -> tuple[float, str]:
    if explicit_pt is not None and explicit_pt > 0:
        return float(explicit_pt), "explicit"

    with ZipFile(docx, "r") as zf:
        document_root = etree.fromstring(zf.read("word/document.xml"))
        styles_root = etree.fromstring(zf.read("word/styles.xml")) if "word/styles.xml" in zf.namelist() else None

    style_sizes, default_style_pt = style_font_size_map(styles_root)
    paras = document_root.findall(".//w:body/w:p", namespaces=NS)

    def most_common(counter: Counter[float]) -> float | None:
        if not counter:
            return None
        return counter.most_common(1)[0][0]

    if paragraph_index is not None and 0 <= paragraph_index < len(paras):
        nearby = Counter()
        ordered_indexes = sorted(
            range(len(paras)),
            key=lambda idx: (abs(idx - paragraph_index), idx),
        )
        for idx in ordered_indexes[:120]:
            p = paras[idx]
            if not is_body_like_text(paragraph_text(p)):
                continue
            for pt in run_font_sizes_pt(p):
                if is_plausible_body_pt(pt):
                    nearby[round(pt, 2)] += 3
            style_id = paragraph_style_id(p)
            if style_id and style_id in style_sizes and is_plausible_body_pt(style_sizes[style_id]):
                nearby[round(style_sizes[style_id], 2)] += 1
            elif not style_id:
                normal_pt = style_sizes.get("Normal") or style_sizes.get("1") or default_style_pt
                if normal_pt is not None and is_plausible_body_pt(normal_pt):
                    nearby[round(normal_pt, 2)] += 1
        detected = most_common(nearby)
        if detected is not None:
            return detected, "nearby-body-runs"

    all_doc = Counter()
    for p in paras:
        if not is_body_like_text(paragraph_text(p)):
            continue
        for pt in run_font_sizes_pt(p):
            if is_plausible_body_pt(pt):
                all_doc[round(pt, 2)] += 1
        style_id = paragraph_style_id(p)
        if style_id and style_id in style_sizes and is_plausible_body_pt(style_sizes[style_id]):
            all_doc[round(style_sizes[style_id], 2)] += 1
        elif not style_id:
            normal_pt = style_sizes.get("Normal") or style_sizes.get("1") or default_style_pt
            if normal_pt is not None and is_plausible_body_pt(normal_pt):
                all_doc[round(normal_pt, 2)] += 1
    detected = most_common(all_doc)
    if detected is not None:
        return detected, "document-body-runs"

    for key in ("Normal", "1"):
        pt = style_sizes.get(key)
        if pt is not None and is_plausible_body_pt(pt):
            return round(pt, 2), f"style:{key}"

    if default_style_pt is not None and is_plausible_body_pt(default_style_pt):
        return round(default_style_pt, 2), "docDefaults"

    return 12.0, "fallback-12pt"


def estimate_formula_lines(mathml: str) -> int:
    if not mathml.strip():
        return 1
    try:
        root = etree.fromstring(mathml.encode("utf-8"), parser=etree.XMLParser(recover=True))
        rows = root.xpath("//*[local-name()='mtr']")
        if rows:
            return max(1, len(rows))
    except Exception:
        pass
    return 1


def resolve_formula_object_size(
    docx: Path,
    paragraph_index: int,
    mathml: str,
    *,
    explicit_height_pt: float | None,
    explicit_body_font_pt: float | None,
    formula_font_scale: float,
    ole_line_height_factor: float,
    formula_lines: int,
) -> dict[str, object]:
    body_pt, body_source = detect_body_font_pt(docx, paragraph_index, explicit_body_font_pt)
    scale = float(formula_font_scale)
    line_factor = float(ole_line_height_factor)
    lines = int(formula_lines) if int(formula_lines or 0) > 0 else estimate_formula_lines(mathml)
    target_formula_font_pt = body_pt * scale
    auto_height_pt = target_formula_font_pt * line_factor * max(1, lines)
    resolved_height_pt = float(explicit_height_pt) if explicit_height_pt is not None else auto_height_pt
    return {
        "body_font_pt": round(body_pt, 2),
        "body_font_source": body_source,
        "formula_font_scale": round(scale, 4),
        "target_formula_font_pt": round(target_formula_font_pt, 2),
        "formula_lines": max(1, lines),
        "ole_line_height_factor": round(line_factor, 4),
        "auto_height_pt": round(auto_height_pt, 2),
        "resolved_height_pt": round(resolved_height_pt, 2),
        "height_source": "explicit-height-pt" if explicit_height_pt is not None else "body-font-auto",
    }


def print_size_info(size_info: dict[str, object]) -> None:
    for key in (
        "body_font_pt",
        "body_font_source",
        "formula_font_scale",
        "target_formula_font_pt",
        "formula_lines",
        "ole_line_height_factor",
        "auto_height_pt",
        "resolved_height_pt",
        "height_source",
    ):
        print(f"{key}={size_info[key]}")


def is_windows() -> bool:
    return os.name == "nt"


def common_mathtype_paths() -> list[Path]:
    candidates: list[Path] = []
    for env_name in ("ProgramFiles(x86)", "ProgramFiles", "ProgramW6432"):
        base = os.environ.get(env_name)
        if not base:
            continue
        for relative in (
            r"MathType\MathType.exe",
            r"MathType 7\MathType.exe",
            r"Design Science\MathType\MathType.exe",
        ):
            candidates.append(Path(base) / relative)
    return candidates


def registry_mathtype_paths() -> list[Path]:
    if not is_windows():
        return []
    try:
        import winreg
    except Exception:
        return []
    keys = (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MathType.exe"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MathType.exe"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\MathType.exe"),
    )
    paths: list[Path] = []
    for hive, key in keys:
        try:
            with winreg.OpenKey(hive, key) as handle:
                value, _ = winreg.QueryValueEx(handle, "")
                if value:
                    paths.append(Path(str(value)))
        except OSError:
            pass
    return paths


def find_mathtype_exe(explicit: str | None = None) -> tuple[Path | None, str]:
    probes: list[tuple[str, Path]] = []
    if explicit:
        probes.append(("argument", Path(explicit).expanduser()))
    env_value = os.environ.get(MATHTYPE_EXE_ENV)
    if env_value:
        probes.append((MATHTYPE_EXE_ENV, Path(env_value).expanduser()))
    probes.extend(("registry", path) for path in registry_mathtype_paths())
    probes.extend(("common-path", path) for path in common_mathtype_paths())
    which = shutil.which(MATHTYPE_EXE_NAME)
    if which:
        probes.append(("PATH", Path(which)))
    for source, path in probes:
        if path.exists() and path.is_file():
            return path.resolve(), source
    return None, "not-found"


def mathtype_troubleshooting_hint() -> str:
    return (
        "Check that MathType is installed, activated/licensed, and registered as an OLE class "
        "(usually Equation.DSMT4). Run: mathtype_word_wps.py check-env --probe-mathtype"
    )


def backend_to_progid(backend: str, explicit_progid: str | None = None) -> str:
    if explicit_progid:
        return explicit_progid
    normalized = (backend or "auto").lower()
    if normalized == "word":
        return "Word.Application"
    if normalized == "wps":
        return "Kwps.Application"
    if normalized == "auto":
        return "Kwps.Application"
    raise ValueError(f"Unsupported backend: {backend}")


def backend_attempts(backend: str, explicit_progid: str | None = None) -> list[tuple[str, str]]:
    if explicit_progid:
        return [("explicit", explicit_progid)]
    normalized = (backend or "auto").lower()
    if normalized == "wps":
        return [("wps", "Kwps.Application")]
    if normalized == "word":
        return [("word", "Word.Application")]
    if normalized == "auto":
        return [("wps", "Kwps.Application"), ("word", "Word.Application")]
    raise ValueError(f"Unsupported backend: {backend}")


def cleanup_leftovers(args: argparse.Namespace) -> None:
    if not is_windows():
        print("cleanup=skipped reason=not-windows")
        return
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            "Get-Process MathType,MathTypeLib -ErrorAction SilentlyContinue | Stop-Process -Force",
        ],
        capture_output=True,
        text=True,
    )
    print("cleanup=done processes=MathType,MathTypeLib")


def check_com_available(progid: str, *, probe: bool) -> tuple[bool, str]:
    if not is_windows():
        return False, "not-windows"
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:
        return False, f"pywin32-import-failed: {exc}"
    if not probe:
        return True, "pywin32-import-ok"
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32.DispatchEx(progid)
        configure_document_app(app, progid, visible=False)
        return True, "dispatch-ok"
    except Exception as exc:
        return False, f"dispatch-failed: {exc}"
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def check_ole_class_registered(class_type: str) -> tuple[bool, str]:
    if not is_windows():
        return False, "not-windows"
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, class_type + r"\CLSID") as handle:
            value, _ = winreg.QueryValueEx(handle, "")
            return True, str(value)
    except Exception as exc:
        return False, str(exc)


def probe_mathtype_runtime(explicit_exe: str | None = None) -> tuple[bool, str]:
    if not is_windows():
        return False, "not-windows"
    mathtype_exe, source = find_mathtype_exe(explicit_exe)
    if mathtype_exe is None:
        return False, "MathType executable not found. " + mathtype_troubleshooting_hint()
    try:
        import win32con
        import win32gui

        before = set(list_mathtype_windows())
        proc = subprocess.Popen([str(mathtype_exe)])
        hwnd = wait_for_mathtype_window(before, timeout=15.0)
        title = win32gui.GetWindowText(hwnd)
        lower_title = title.lower()
        activation_terms = (
            "activate",
            "activation",
            "license",
            "licence",
            "expired",
            "trial",
            "evaluation",
            "激活",
            "许可证",
            "授權",
            "授权",
            "试用",
            "過期",
            "过期",
        )
        if any(term in lower_title for term in activation_terms):
            ok = False
            detail = f"possible activation/license dialog title={title!r}"
        else:
            ok = True
            detail = f"editor-window-found title={title!r} source={source}"
        try:
            if win32gui.IsWindow(hwnd):
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass
        try:
            proc.terminate()
        except Exception:
            pass
        return ok, detail
    except Exception as exc:
        return False, f"MathType runtime probe failed: {exc}. " + mathtype_troubleshooting_hint()


def check_env(args: argparse.Namespace) -> None:
    print(f"python={sys.executable}")
    print(f"python_version={platform.python_version()}")
    print(f"platform={platform.platform()}")
    print(f"windows={is_windows()}")
    print(f"lxml=ok version={etree.LXML_VERSION}")
    try:
        import win32com.client  # noqa: F401
        import win32gui  # noqa: F401
        print("pywin32=ok")
    except Exception as exc:
        print(f"pywin32=missing detail={exc}")
    try:
        import latex2mathml  # noqa: F401
        print("latex2mathml=ok")
    except Exception as exc:
        print(f"latex2mathml=missing optional=true detail={exc}")

    mathtype_exe, source = find_mathtype_exe(args.mathtype_exe)
    print(f"mathtype_exe={mathtype_exe if mathtype_exe else ''}")
    print(f"mathtype_exe_source={source}")
    if mathtype_exe is None:
        print("mathtype_warning=not-found; install MathType or set MATHTYPE_EXE / --mathtype-exe")

    class_ok, class_detail = check_ole_class_registered(args.ole_class_type)
    print(f"ole_class={args.ole_class_type} registered={class_ok} detail={class_detail}")
    if not class_ok:
        print("mathtype_warning=ole-class-not-registered; reinstall/repair MathType or check --ole-class-type")

    if args.probe_mathtype:
        probe_ok, probe_detail = probe_mathtype_runtime(args.mathtype_exe)
        print(f"mathtype_runtime_probe={probe_ok} detail={probe_detail}")
        if not probe_ok:
            print("mathtype_warning=runtime-or-activation-problem; verify MathType activation/license manually")
    else:
        print("mathtype_runtime_probe=skipped use --probe-mathtype to detect activation/runtime issues")

    for backend in ("wps", "word"):
        progid = backend_to_progid(backend)
        ok, detail = check_com_available(progid, probe=bool(args.probe_com))
        print(f"backend={backend} progid={progid} available={ok} detail={detail}")


def list_mathtype_windows() -> list[int]:
    import win32gui

    handles: list[int] = []

    def callback(hwnd, _):
        try:
            if win32gui.GetClassName(hwnd) == "EQNWINCLASS":
                handles.append(int(hwnd))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(callback, None)
    return handles


def wait_for_mathtype_window(before: set[int], timeout: float = 25.0) -> int:
    deadline = time.time() + timeout
    while time.time() < deadline:
        current = list_mathtype_windows()
        new_handles = [hwnd for hwnd in current if hwnd not in before]
        if new_handles:
            return int(new_handles[-1])
        if current and not before:
            return int(current[-1])
        time.sleep(0.2)
    raise TimeoutError("MathType window was not found.")


def set_mathml_clipboard(mathml: str) -> None:
    import win32clipboard

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(mathml)
        custom = {
            name: win32clipboard.RegisterClipboardFormat(name)
            for name in ("MathML", "MathML Presentation", "application/mathml+xml")
        }
        win32clipboard.SetClipboardData(custom["MathML"], mathml.encode("utf-16le"))
        win32clipboard.SetClipboardData(custom["MathML Presentation"], mathml.encode("utf-8"))
        win32clipboard.SetClipboardData(custom["application/mathml+xml"], mathml.encode("utf-16le"))
    finally:
        win32clipboard.CloseClipboard()


def capture_window(hwnd: int, output: Path) -> None:
    import win32con
    import win32gui
    import win32ui

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = max(1, right - left)
    height = max(1, bottom - top)
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    src_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    mem_dc = src_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(src_dc, width, height)
    old_obj = mem_dc.SelectObject(bitmap)
    try:
        ok = ctypes.windll.user32.PrintWindow(hwnd, mem_dc.GetSafeHdc(), 2)
        if not ok:
            mem_dc.BitBlt((0, 0), (width, height), src_dc, (0, 0), win32con.SRCCOPY)
        bitmap.SaveBitmapFile(mem_dc, str(output))
    finally:
        try:
            mem_dc.SelectObject(old_obj)
        except Exception:
            pass
        try:
            win32gui.DeleteObject(bitmap.GetHandle())
        except Exception:
            pass
        try:
            mem_dc.DeleteDC()
        except Exception:
            pass
        try:
            src_dc.DeleteDC()
        except Exception:
            pass
        win32gui.ReleaseDC(hwnd, hwnd_dc)


def save_clipboard_metafiles(wmf_path: Path, emf_path: Path) -> dict[str, object]:
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32
    kernel32 = ctypes.windll.kernel32
    user32.GetClipboardData.restype = ctypes.c_void_p
    kernel32.GlobalLock.restype = ctypes.c_void_p
    CF_METAFILEPICT = 3
    CF_ENHMETAFILE = 14

    result: dict[str, object] = {}
    if not user32.OpenClipboard(None):
        raise ctypes.WinError()
    try:
        hemf = user32.GetClipboardData(CF_ENHMETAFILE)
        if hemf:
            copied = gdi32.CopyEnhMetaFileW(ctypes.c_void_p(hemf), str(emf_path))
            if copied:
                gdi32.DeleteEnhMetaFile(copied)
            result["emf"] = str(emf_path)
            result["emf_size"] = emf_path.stat().st_size if emf_path.exists() else 0

        hglobal = user32.GetClipboardData(CF_METAFILEPICT)
        if hglobal:
            ptr = kernel32.GlobalLock(ctypes.c_void_p(hglobal))
            if ptr:
                class METAFILEPICT(ctypes.Structure):
                    _fields_ = [
                        ("mm", ctypes.c_long),
                        ("xExt", ctypes.c_long),
                        ("yExt", ctypes.c_long),
                        ("hMF", ctypes.c_void_p),
                    ]

                meta = METAFILEPICT.from_address(ptr)
                copied = gdi32.CopyMetaFileW(ctypes.c_void_p(meta.hMF), str(wmf_path))
                if copied:
                    gdi32.DeleteMetaFile(copied)
                result.update(
                    {
                        "wmf": str(wmf_path),
                        "wmf_size": wmf_path.stat().st_size if wmf_path.exists() else 0,
                        "metafile_map_mode": int(meta.mm),
                        "metafile_xExt": int(meta.xExt),
                        "metafile_yExt": int(meta.yExt),
                    }
                )
                kernel32.GlobalUnlock(ctypes.c_void_p(hglobal))
    finally:
        user32.CloseClipboard()
    if not wmf_path.exists():
        raise RuntimeError("MathType copy did not expose a WMF metafile.")
    return result


def make_wmf(args: argparse.Namespace) -> None:
    import win32con
    import win32gui

    if args.mathml_file:
        mathml = read_text(Path(args.mathml_file))
    elif args.latex_file:
        mathml = latex_to_mathml(read_text(Path(args.latex_file)))
    elif args.mathml:
        mathml = args.mathml
    elif args.latex:
        mathml = latex_to_mathml(args.latex)
    else:
        raise SystemExit("Provide --mathml-file, --mathml, --latex-file, or --latex.")

    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    base = args.base_name
    wmf_path = output_dir / f"{base}.wmf"
    emf_path = output_dir / f"{base}.emf"
    shot_path = output_dir / f"{base}_mathtype_window.bmp"

    mathtype_exe, source = find_mathtype_exe(args.mathtype_exe)
    if mathtype_exe is None:
        raise RuntimeError(
            "MathType executable was not found. Set --mathtype-exe or the MATHTYPE_EXE environment variable. "
            + mathtype_troubleshooting_hint()
        )
    print(f"mathtype_exe={mathtype_exe} source={source}")

    before = set(list_mathtype_windows())
    subprocess.Popen([str(mathtype_exe)])
    try:
        hwnd = wait_for_mathtype_window(before)
    except Exception as exc:
        raise RuntimeError(str(exc) + ". " + mathtype_troubleshooting_hint()) from exc
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(args.wait_seconds)

    set_mathml_clipboard(mathml)
    win32gui.SendMessage(hwnd, win32con.WM_COMMAND, 505, 0)  # Edit > Paste
    time.sleep(args.wait_seconds)
    capture_window(hwnd, shot_path)
    win32gui.SendMessage(hwnd, win32con.WM_COMMAND, 513, 0)  # Edit > Select All
    time.sleep(0.2)
    win32gui.SendMessage(hwnd, win32con.WM_COMMAND, 504, 0)  # Edit > Copy
    time.sleep(0.8)
    result = save_clipboard_metafiles(wmf_path, emf_path)

    if not args.keep_window:
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass

    print(f"wmf={result.get('wmf')} size={result.get('wmf_size')}")
    print(f"emf={result.get('emf')} size={result.get('emf_size')}")
    print(f"screenshot={shot_path}")


def parse_emf_aspect(emf_path: Path) -> float | None:
    if not emf_path.exists() or emf_path.stat().st_size < 48:
        return None
    data = emf_path.read_bytes()
    # ENHMETAHEADER: iType, nSize, rclBounds, rclFrame.
    left, top, right, bottom = struct.unpack_from("<iiii", data, 24)
    width = right - left
    height = bottom - top
    if width > 0 and height > 0:
        return width / height
    return None


def infer_aspect(image_path: Path) -> float:
    candidates = []
    if image_path.suffix.lower() == ".emf":
        candidates.append(image_path)
    candidates.append(image_path.with_suffix(".emf"))
    for candidate in candidates:
        aspect = parse_emf_aspect(candidate)
        if aspect:
            return aspect
    raise RuntimeError(
        "Unable to infer formula aspect. Provide --width-pt or keep a same-stem .emf next to the WMF."
    )


def relationship_id(existing: list[str]) -> str:
    used = set(existing)
    index = 1
    while f"rId{index}" in used:
        index += 1
    return f"rId{index}"


def next_media_name(zip_file: ZipFile, image_path: Path) -> str:
    ext = image_path.suffix.lower()
    stem = image_path.stem
    existing = set(zip_file.namelist())
    index = 1
    while True:
        name = f"word/media/{stem}{'' if index == 1 else '_' + str(index)}{ext}"
        if name not in existing:
            return name
        index += 1


def ensure_content_type(root: etree._Element, extension: str) -> None:
    normalized = extension.lower().lstrip(".")
    exists = any(
        child.tag == CT + "Default" and child.get("Extension", "").lower() == normalized
        for child in root
    )
    if exists:
        return
    content_type = {
        "wmf": "image/x-wmf",
        "emf": "image/x-emf",
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
    }.get(normalized)
    if not content_type:
        raise RuntimeError(f"Unsupported image extension for content type: {extension}")
    node = etree.Element(CT + "Default")
    node.set("Extension", normalized)
    node.set("ContentType", content_type)
    root.insert(0, node)


def max_docpr_id(root: etree._Element) -> int:
    values = []
    for node in root.findall(".//wp:docPr", namespaces=NS):
        try:
            values.append(int(node.get("id", "0")))
        except ValueError:
            pass
    return max(values or [0])


def make_picture_drawing(rid: str, width_emu: int, height_emu: int, docpr_id: int, name: str) -> etree._Element:
    drawing = etree.Element(W + "drawing")
    inline = etree.SubElement(drawing, WP + "inline")
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "114300")
    inline.set("distR", "114300")
    extent = etree.SubElement(inline, WP + "extent")
    extent.set("cx", str(width_emu))
    extent.set("cy", str(height_emu))
    effect = etree.SubElement(inline, WP + "effectExtent")
    for key in ("l", "t", "r", "b"):
        effect.set(key, "0")
    docpr = etree.SubElement(inline, WP + "docPr")
    docpr.set("id", str(docpr_id))
    docpr.set("name", name)
    docpr.set("descr", Path(name).stem)
    c_nv = etree.SubElement(inline, WP + "cNvGraphicFramePr")
    locks = etree.SubElement(c_nv, A + "graphicFrameLocks")
    locks.set("noChangeAspect", "1")
    graphic = etree.SubElement(inline, A + "graphic")
    graphic_data = etree.SubElement(graphic, A + "graphicData")
    graphic_data.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")
    pic = etree.SubElement(graphic_data, PIC + "pic")
    nv = etree.SubElement(pic, PIC + "nvPicPr")
    cnvpr = etree.SubElement(nv, PIC + "cNvPr")
    cnvpr.set("id", str(docpr_id))
    cnvpr.set("name", name)
    cnvpr.set("descr", Path(name).stem)
    cnvpic = etree.SubElement(nv, PIC + "cNvPicPr")
    etree.SubElement(cnvpic, A + "picLocks").set("noChangeAspect", "1")
    blip_fill = etree.SubElement(pic, PIC + "blipFill")
    blip = etree.SubElement(blip_fill, A + "blip")
    blip.set(R + "embed", rid)
    stretch = etree.SubElement(blip_fill, A + "stretch")
    etree.SubElement(stretch, A + "fillRect")
    sppr = etree.SubElement(pic, PIC + "spPr")
    xfrm = etree.SubElement(sppr, A + "xfrm")
    off = etree.SubElement(xfrm, A + "off")
    off.set("x", "0")
    off.set("y", "0")
    ext = etree.SubElement(xfrm, A + "ext")
    ext.set("cx", str(width_emu))
    ext.set("cy", str(height_emu))
    geom = etree.SubElement(sppr, A + "prstGeom")
    geom.set("prst", "rect")
    etree.SubElement(geom, A + "avLst")
    return drawing


def set_extent(drawing: etree._Element, width_emu: int, height_emu: int) -> None:
    for ext in drawing.findall(".//wp:extent", namespaces=NS):
        ext.set("cx", str(width_emu))
        ext.set("cy", str(height_emu))
    for ext in drawing.findall(".//a:ext", namespaces=NS):
        ext.set("cx", str(width_emu))
        ext.set("cy", str(height_emu))


def paragraph_text_width_twips(root: etree._Element) -> int:
    sect = root.find(".//w:sectPr", namespaces=NS)
    if sect is None:
        raise RuntimeError("No sectPr found; cannot calculate equation tab stops.")
    pg_sz = sect.find("w:pgSz", namespaces=NS)
    mar = sect.find("w:pgMar", namespaces=NS)
    if pg_sz is None or mar is None:
        raise RuntimeError("Missing page size or margins in sectPr.")
    page_w = int(pg_sz.get(W + "w"))
    left = int(mar.get(W + "left", "0"))
    right = int(mar.get(W + "right", "0"))
    return page_w - left - right


def normalize_formula_paragraph(
    p: etree._Element,
    drawing: etree._Element,
    text_width_twips: int,
    number: str,
    font_size_half_points: int,
) -> None:
    ppr = p.find("w:pPr", namespaces=NS)
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        p.insert(0, ppr)

    for child in list(p):
        if child is not ppr:
            p.remove(child)

    for node in ppr.findall("w:tabs", namespaces=NS):
        ppr.remove(node)
    tabs = etree.SubElement(ppr, W + "tabs")
    tab_center = etree.SubElement(tabs, W + "tab")
    tab_center.set(W + "val", "center")
    tab_center.set(W + "pos", str(text_width_twips // 2))
    tab_right = etree.SubElement(tabs, W + "tab")
    tab_right.set(W + "val", "right")
    tab_right.set(W + "pos", str(text_width_twips))

    ind = ppr.find("w:ind", namespaces=NS)
    if ind is None:
        ind = etree.SubElement(ppr, W + "ind")
    for attr in ("left", "leftChars", "right", "rightChars", "firstLine", "firstLineChars"):
        ind.set(W + attr, "0")

    jc = ppr.find("w:jc", namespaces=NS)
    if jc is None:
        jc = etree.SubElement(ppr, W + "jc")
    jc.set(W + "val", "left")

    def run_tab() -> etree._Element:
        run = etree.Element(W + "r")
        etree.SubElement(run, W + "tab")
        return run

    def run_text(value: str) -> etree._Element:
        run = etree.Element(W + "r")
        rpr = etree.SubElement(run, W + "rPr")
        fonts = etree.SubElement(rpr, W + "rFonts")
        fonts.set(W + "ascii", "Times New Roman")
        fonts.set(W + "hAnsi", "Times New Roman")
        fonts.set(W + "eastAsia", "宋体")
        sz = etree.SubElement(rpr, W + "sz")
        sz.set(W + "val", str(font_size_half_points))
        szcs = etree.SubElement(rpr, W + "szCs")
        szcs.set(W + "val", str(font_size_half_points))
        text = etree.SubElement(run, W + "t")
        text.text = value
        return run

    img_run = etree.Element(W + "r")
    img_run.append(drawing)
    p.append(run_tab())
    p.append(img_run)
    p.append(run_tab())
    p.append(run_text(number))


def copy_docx_with_updates(
    docx: Path,
    output: Path,
    document_xml: bytes,
    rels_xml: bytes | None = None,
    content_types_xml: bytes | None = None,
    new_media_name: str | None = None,
    new_media_data: bytes | None = None,
) -> None:
    tmp = output.with_suffix(output.suffix + ".tmp")
    with ZipFile(docx, "r") as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        written = set()
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = document_xml
            elif rels_xml is not None and item.filename == "word/_rels/document.xml.rels":
                data = rels_xml
            elif content_types_xml is not None and item.filename == "[Content_Types].xml":
                data = content_types_xml
            zout.writestr(item, data)
            written.add(item.filename)
        if new_media_name and new_media_data is not None and new_media_name not in written:
            zout.writestr(new_media_name, new_media_data)
    shutil.move(tmp, output)


def replace_docx(args: argparse.Namespace) -> None:
    docx = Path(args.docx).resolve()
    if not docx.exists():
        raise FileNotFoundError(docx)
    output = Path(args.output).resolve() if args.output else docx
    if output == docx and args.backup:
        backup = write_backup(docx, "formula")
        print(f"backup={backup}")

    paragraph_index = int(args.paragraph_index)
    if args.one_based:
        paragraph_index -= 1
    if paragraph_index < 0:
        raise ValueError("paragraph index must be positive after index-base conversion")

    size_info: dict[str, object] | None = None
    with ZipFile(docx, "r") as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        rels_root = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
        ct_root = etree.fromstring(zin.read("[Content_Types].xml"))
        paras = root.findall(".//w:body/w:p", namespaces=NS)
        if paragraph_index >= len(paras):
            raise IndexError(f"paragraph index out of range: {paragraph_index} / {len(paras)}")
        p = paras[paragraph_index]
        existing_drawings = p.findall(".//w:drawing", namespaces=NS)

        if args.scale_existing:
            if len(existing_drawings) != 1:
                raise RuntimeError("scale-existing requires exactly one existing drawing in the paragraph.")
            drawing = existing_drawings[0]
            drawing.getparent().remove(drawing)
            ext = drawing.find(".//wp:extent", namespaces=NS)
            if ext is None:
                raise RuntimeError("existing drawing has no wp:extent")
            width_emu = round(int(ext.get("cx")) * float(args.scale_existing))
            height_emu = round(int(ext.get("cy")) * float(args.scale_existing))
            set_extent(drawing, width_emu, height_emu)
            new_media_name = None
            new_media_data = None
            rels_xml = None
            ct_xml = None
        else:
            if not args.image:
                raise RuntimeError("Provide --image, or use --scale-existing to resize the current paragraph image.")
            image = Path(args.image).resolve()
            if not image.exists():
                raise FileNotFoundError(image)
            if args.width_pt:
                width_pt = float(args.width_pt)
            elif args.height_pt:
                width_pt = float(args.height_pt) * infer_aspect(image)
            else:
                size_info = resolve_formula_object_size(
                    docx,
                    paragraph_index,
                    "",
                    explicit_height_pt=None,
                    explicit_body_font_pt=args.body_font_pt,
                    formula_font_scale=args.formula_font_scale,
                    ole_line_height_factor=args.image_line_height_factor,
                    formula_lines=args.formula_lines,
                )
                width_pt = float(size_info["resolved_height_pt"]) * infer_aspect(image)
            height_pt = float(args.height_pt) if args.height_pt else width_pt / infer_aspect(image)
            if size_info is None and args.height_pt:
                size_info = resolve_formula_object_size(
                    docx,
                    paragraph_index,
                    "",
                    explicit_height_pt=args.height_pt,
                    explicit_body_font_pt=args.body_font_pt,
                    formula_font_scale=args.formula_font_scale,
                    ole_line_height_factor=args.image_line_height_factor,
                    formula_lines=args.formula_lines,
                )
            width_emu = round(width_pt * EMU_PER_PT)
            height_emu = round(height_pt * EMU_PER_PT)

            existing_rids = [rel.get("Id") for rel in rels_root.findall(f"{REL}Relationship")]
            rid = relationship_id([rid for rid in existing_rids if rid])
            new_media_name = next_media_name(zin, image)
            new_media_data = image.read_bytes()
            rel = etree.SubElement(rels_root, REL + "Relationship")
            rel.set("Id", rid)
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            rel.set("Target", new_media_name.replace("word/", ""))
            ensure_content_type(ct_root, image.suffix)
            drawing = make_picture_drawing(rid, width_emu, height_emu, max_docpr_id(root) + 1, image.name)
            rels_xml = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
            ct_xml = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

        normalize_formula_paragraph(
            p,
            drawing,
            paragraph_text_width_twips(root),
            args.number,
            int(args.font_size_half_points),
        )

        document_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        copy_docx_with_updates(
            docx,
            output,
            document_xml,
            rels_xml=rels_xml,
            content_types_xml=ct_xml,
            new_media_name=new_media_name,
            new_media_data=new_media_data,
        )

    info = inspect_docx_path(output, paragraph_index)
    print_inspection(info)
    if size_info is not None:
        print_size_info(size_info)
    if args.validate_wps:
        print(f"wps_open={validate_wps_open(output)}")


def inspect_docx_path(docx: Path, paragraph_index: int) -> dict[str, object]:
    with ZipFile(docx, "r") as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
        rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
        relmap = {
            rel.get("Id"): (rel.get("Type"), rel.get("Target"))
            for rel in rels_root.findall(f"{REL}Relationship")
        }
        paras = root.findall(".//w:body/w:p", namespaces=NS)
        p = paras[paragraph_index]
        ppr = p.find("w:pPr", namespaces=NS)
        ext = p.find(".//wp:extent", namespaces=NS)
        text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
        ole_objects = p.findall(".//o:OLEObject", namespaces=NS)
        ole_rids = [node.get(R + "id") for node in ole_objects]
        v_shapes = p.findall(".//v:shape", namespaces=NS)
        info: dict[str, object] = {
            "docx": str(docx),
            "paragraph_index": paragraph_index,
            "paragraph_count": len(paras),
            "text": text,
            "run_tabs": len(p.findall(".//w:tab", namespaces=NS)),
            "blips": len(p.findall(".//a:blip", namespaces=NS)),
            "ole_objects": len(ole_objects),
            "ole_progids": [node.get("ProgID") for node in ole_objects],
            "ole_targets": [relmap.get(rid) for rid in ole_rids],
            "omath": len(p.findall(".//m:oMath", namespaces=NS)),
            "total_omath": len(root.findall(".//m:oMath", namespaces=NS)),
        }
        if v_shapes:
            style = v_shapes[0].get("style", "")
            parsed_style = parse_css_pt(style)
            info["ole_shape_style"] = style
            if "width" in parsed_style:
                info["ole_shape_width_pt"] = round(parsed_style["width"], 2)
            if "height" in parsed_style:
                info["ole_shape_height_pt"] = round(parsed_style["height"], 2)
        if ext is not None:
            info["image_width_pt"] = round(int(ext.get("cx")) / EMU_PER_PT, 2)
            info["image_height_pt"] = round(int(ext.get("cy")) / EMU_PER_PT, 2)
        if ppr is not None:
            jc = ppr.find("w:jc", namespaces=NS)
            if jc is not None:
                info["jc"] = jc.get(W + "val")
            ind = ppr.find("w:ind", namespaces=NS)
            if ind is not None:
                info["ind"] = {key.split("}", 1)[-1]: value for key, value in ind.attrib.items()}
            info["tab_stops"] = [
                (tab.get(W + "val"), tab.get(W + "pos"))
                for tab in ppr.findall(".//w:tabs/w:tab", namespaces=NS)
            ]
        return info


def print_inspection(info: dict[str, object]) -> None:
    for key in (
        "docx",
        "paragraph_index",
        "text",
        "run_tabs",
        "blips",
        "ole_objects",
        "ole_progids",
        "ole_targets",
        "omath",
        "total_omath",
        "ole_shape_width_pt",
        "ole_shape_height_pt",
        "ole_shape_style",
        "image_width_pt",
        "image_height_pt",
        "jc",
        "tab_stops",
        "ind",
    ):
        if key in info:
            print(f"{key}={info[key]}")


def validate_wps_open(docx: Path) -> str:
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32.DispatchEx("Kwps.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(str(docx), ReadOnly=True, AddToRecentFiles=False)
        return f"ok name={doc.Name} paragraphs={doc.Paragraphs.Count} inline_shapes={doc.InlineShapes.Count}"
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


def call_with_retry(action, attempts: int = 90, delay: float = 0.5):
    last_error: Exception | None = None
    for _ in range(max(1, int(attempts))):
        try:
            return action()
        except Exception as exc:  # COM calls can transiently reject automation.
            last_error = exc
            try:
                import pythoncom

                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(delay)
    raise last_error  # type: ignore[misc]


def is_word_progid(progid: str) -> bool:
    return "word.application" in (progid or "").lower()


def set_optional_com_attr(obj, name: str, value) -> None:
    try:
        setattr(obj, name, value)
    except Exception:
        pass


def configure_document_app(app, progid: str, *, visible: bool) -> None:
    set_optional_com_attr(app, "Visible", bool(visible))
    set_optional_com_attr(app, "DisplayAlerts", 0)
    set_optional_com_attr(app, "ScreenUpdating", False)

    # Word-specific quiet automation settings. WPS ignores or lacks most of
    # these, so keep them optional and non-fatal.
    if is_word_progid(progid):
        set_optional_com_attr(app, "AutomationSecurity", 3)  # Force-disable macros.
        set_optional_com_attr(app, "FeatureInstall", 0)
        options = getattr(app, "Options", None)
        if options is not None:
            for name, value in (
                ("ConfirmConversions", False),
                ("SaveNormalPrompt", False),
                ("UpdateLinksAtOpen", False),
                ("CheckSpellingAsYouType", False),
                ("CheckGrammarAsYouType", False),
                ("SuggestSpellingCorrections", False),
                ("BackgroundSave", False),
                ("AllowReadingMode", False),
            ):
                set_optional_com_attr(options, name, value)


def open_word_app(progid: str, *, visible: bool, allow_fallback: bool = True):
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    candidates: list[str] = []
    preferred = (progid or "Kwps.Application").strip()
    if is_word_progid(preferred):
        candidates.append(preferred)
    elif preferred:
        candidates.append(preferred)
        if allow_fallback:
            for fallback in ("Kwps.Application", "Word.Application"):
                if fallback not in candidates:
                    candidates.append(fallback)

    last_error: Exception | None = None
    app = None
    for candidate in candidates:
        try:
            app = win32.DispatchEx(candidate)
            break
        except Exception as exc:
            last_error = exc
    if app is None:
        pythoncom.CoUninitialize()
        raise RuntimeError(f"Cannot start WPS/Word COM: {last_error}") from last_error

    configure_document_app(app, candidate, visible=visible)
    return pythoncom, app


def open_document(app, docx_path: Path, *, read_only: bool, com_progid: str = ""):
    if is_word_progid(com_progid):
        doc = call_with_retry(
            lambda: app.Documents.Open(
                FileName=str(docx_path.resolve()),
                ConfirmConversions=False,
                ReadOnly=read_only,
                AddToRecentFiles=False,
                PasswordDocument="",
                PasswordTemplate="",
                Revert=False,
                WritePasswordDocument="",
                WritePasswordTemplate="",
                Format=0,
                Encoding=0,
                Visible=False,
                OpenAndRepair=False,
                DocumentDirection=0,
                NoEncodingDialog=True,
            )
        )
    else:
        doc = call_with_retry(
            lambda: app.Documents.Open(str(docx_path.resolve()), ReadOnly=read_only, AddToRecentFiles=False)
        )
    time.sleep(2.0)
    if not read_only:
        try:
            if bool(doc.ReadOnly):
                raise RuntimeError(f"Document opened read-only through {com_progid or 'COM'}: {docx_path}")
        except AttributeError:
            pass
    return doc


def try_exclusive_open(path: Path) -> tuple[bool, str]:
    try:
        import pywintypes
        import win32con
        import win32file

        handle = win32file.CreateFile(
            str(path.resolve()),
            win32con.GENERIC_READ | win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )
        win32file.CloseHandle(handle)
        return True, ""
    except Exception as exc:
        return False, str(exc)


def wait_for_exclusive_file(path: Path, timeout: float = 20.0, interval: float = 0.5) -> None:
    deadline = time.time() + float(timeout)
    last_error = ""
    while time.time() < deadline:
        ok, detail = try_exclusive_open(path)
        if ok:
            return
        last_error = detail
        time.sleep(float(interval))
    raise RuntimeError(f"File is still locked and cannot be opened exclusively: {path}; last_error={last_error}")


def insert_blank_mathtype_ole(doc, target_range, class_type: str = "Equation.DSMT4"):
    return call_with_retry(
        lambda rng=target_range: doc.InlineShapes.AddOLEObject(
            ClassType=class_type,
            FileName="",
            LinkToFile=False,
            DisplayAsIcon=False,
            Range=rng,
        )
    )


def find_new_mathtype_window(before: set[int], timeout: float = 20.0) -> int:
    deadline = time.time() + float(timeout)
    last_seen: list[int] = []
    while time.time() < deadline:
        current = [int(hwnd) for hwnd in list_mathtype_windows()]
        last_seen = current
        new_handles = [hwnd for hwnd in current if hwnd not in before]
        if new_handles:
            return int(new_handles[-1])
        if current and not before:
            return int(current[-1])
        time.sleep(0.1)
    if len(last_seen) == 1:
        return int(last_seen[0])
    raise TimeoutError(f"MathType editor window was not found. Current windows: {last_seen}")


def capture_saved_ole_screenshot(
    docx: Path,
    paragraph_index: int,
    screenshot: Path,
    *,
    com_progid: str,
    window_timeout: float,
) -> None:
    import win32con
    import win32gui

    pythoncom = None
    app = None
    doc = None
    hwnd = 0
    try:
        pythoncom, app = open_word_app(com_progid, visible=False, allow_fallback=False)
        doc = open_document(app, docx, read_only=False, com_progid=com_progid)
        paragraph = call_with_retry(lambda: doc.Paragraphs(paragraph_index + 1))
        inline_shapes = call_with_retry(lambda p=paragraph: p.Range.InlineShapes)
        if int(call_with_retry(lambda: inline_shapes.Count)) < 1:
            raise RuntimeError("No inline shape found for saved OLE screenshot.")
        shape = call_with_retry(lambda: inline_shapes(1))
        before = set(list_mathtype_windows())
        call_with_retry(lambda: shape.OLEFormat.DoVerb(0))
        hwnd = find_new_mathtype_window(before, timeout=window_timeout)
        time.sleep(2.0)
        capture_window(hwnd, screenshot)
    finally:
        if hwnd:
            try:
                import win32gui
                import win32con

                if win32gui.IsWindow(hwnd):
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
            time.sleep(1.0)
        try:
            if doc is not None:
                call_with_retry(lambda: doc.Close(False))
        finally:
            try:
                if app is not None:
                    call_with_retry(lambda: app.Quit())
            finally:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()
    wait_for_exclusive_file(docx, timeout=window_timeout)


def replace_docx_ole_once(
    args: argparse.Namespace,
    *,
    output: Path,
    paragraph_index: int,
    mathml: str,
    size_info: dict[str, object],
    backend_label: str,
    com_progid: str,
) -> None:
    import win32con
    import win32gui

    pythoncom = None
    word = None
    doc = None
    hwnd = 0
    try:
        print(f"backend_attempt={backend_label} progid={com_progid}")
        pythoncom, word = open_word_app(com_progid, visible=False, allow_fallback=False)
        doc = open_document(word, output, read_only=False, com_progid=com_progid)
        paragraph = call_with_retry(lambda: doc.Paragraphs(paragraph_index + 1))
        rng = call_with_retry(lambda p=paragraph: p.Range.Duplicate)
        rng.End = rng.End - 1
        call_with_retry(lambda r=rng: setattr(r, "Text", f"\t\t{args.number}"))

        paragraph = call_with_retry(lambda: doc.Paragraphs(paragraph_index + 1))
        pf = call_with_retry(lambda p=paragraph: p.Range.ParagraphFormat)
        call_with_retry(lambda: pf.TabStops.ClearAll())
        pf.LeftIndent = 0
        pf.RightIndent = 0
        pf.FirstLineIndent = 0
        pf.Alignment = 0
        center_pt = float(args.center_tab_twips) / 20.0
        right_pt = float(args.right_tab_twips) / 20.0
        call_with_retry(lambda: pf.TabStops.Add(Position=center_pt, Alignment=1))
        call_with_retry(lambda: pf.TabStops.Add(Position=right_pt, Alignment=2))

        start = int(call_with_retry(lambda p=paragraph: p.Range.Start))
        insert_range = call_with_retry(lambda: doc.Range(start + 1, start + 1))
        before = set(list_mathtype_windows())
        shape = insert_blank_mathtype_ole(doc, insert_range, class_type=args.ole_class_type)
        try:
            hwnd = find_new_mathtype_window(before, timeout=float(args.window_timeout))
        except TimeoutError:
            call_with_retry(lambda target=shape: target.OLEFormat.DoVerb(0))
            hwnd = find_new_mathtype_window(before, timeout=float(args.window_timeout))
        if args.foreground_editor:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        time.sleep(1.0)

        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, MATHTYPE_SELECT_ALL_COMMAND_ID, 0)
        time.sleep(0.3)

        set_mathml_clipboard(mathml)
        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, MATHTYPE_PASTE_COMMAND_ID, 0)
        time.sleep(float(args.after_paste_wait))
        if args.editor_screenshot:
            capture_window(hwnd, Path(args.editor_screenshot).resolve())

        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, MATHTYPE_UPDATE_HOST_COMMAND_ID, 0)
        time.sleep(float(args.after_update_wait))
        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, MATHTYPE_CLOSE_RETURN_COMMAND_ID, 0)
        deadline = time.time() + float(args.window_timeout)
        while time.time() < deadline and win32gui.IsWindow(hwnd):
            time.sleep(0.5)
        if win32gui.IsWindow(hwnd):
            raise RuntimeError("MathType window did not close after Close and Return command.")

        paragraph = call_with_retry(lambda: doc.Paragraphs(paragraph_index + 1))
        inline_shapes = call_with_retry(lambda p=paragraph: p.Range.InlineShapes)
        shape_count = int(call_with_retry(lambda: inline_shapes.Count))
        if shape_count:
            inserted = call_with_retry(lambda: inline_shapes(1))
            try:
                inserted.LockAspectRatio = True
            except Exception:
                pass
            inserted.Height = float(size_info["resolved_height_pt"])

        call_with_retry(lambda: doc.Save())
        print(f"backend_used={backend_label} progid={com_progid}")
    finally:
        if hwnd:
            try:
                if win32gui.IsWindow(hwnd):
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
            time.sleep(1.0)
        try:
            if doc is not None:
                call_with_retry(lambda: doc.Close(False))
        finally:
            try:
                if word is not None:
                    call_with_retry(lambda: word.Quit())
            finally:
                if pythoncom is not None:
                    pythoncom.CoUninitialize()


def replace_docx_ole(args: argparse.Namespace) -> None:
    """Replace a formula paragraph with a real MathType Equation.DSMT4 object.

    This follows the WPS insert-object workflow: create a blank MathType OLE
    object, paste MathML into the embedded MathType editor, then trigger
    MathType's "Update host document" and "Close and Return" commands.
    """

    docx = Path(args.docx).resolve()
    if not docx.exists():
        raise FileNotFoundError(docx)
    output = Path(args.output).resolve() if args.output else docx
    if output != docx:
        shutil.copyfile(docx, output)
    elif args.backup:
        backup = write_backup(docx, "formula_ole")
        print(f"backup={backup}")
    wait_for_exclusive_file(output, timeout=float(args.window_timeout))

    paragraph_index = int(args.paragraph_index)
    if args.one_based:
        paragraph_index -= 1
    if paragraph_index < 0:
        raise ValueError("paragraph index must be positive after index-base conversion")

    if args.mathml_file:
        mathml = read_text(Path(args.mathml_file).resolve())
    elif args.mathml:
        mathml = str(args.mathml)
    elif args.latex_file:
        mathml = latex_to_mathml(read_text(Path(args.latex_file).resolve()))
    elif args.latex:
        mathml = latex_to_mathml(str(args.latex))
    else:
        raise RuntimeError("Provide --mathml-file, --mathml, --latex-file, or --latex.")

    size_info = resolve_formula_object_size(
        output,
        paragraph_index,
        mathml,
        explicit_height_pt=args.height_pt,
        explicit_body_font_pt=args.body_font_pt,
        formula_font_scale=args.formula_font_scale,
        ole_line_height_factor=args.ole_line_height_factor,
        formula_lines=args.formula_lines,
    )

    if args.cleanup_mathtype:
        cleanup_leftovers(args)

    used_backend: tuple[str, str] | None = None
    errors: list[str] = []
    attempts = backend_attempts(args.backend, args.com_progid)
    for backend_label, com_progid in attempts:
        try:
            replace_docx_ole_once(
                args,
                output=output,
                paragraph_index=paragraph_index,
                mathml=mathml,
                size_info=size_info,
                backend_label=backend_label,
                com_progid=com_progid,
            )
            used_backend = (backend_label, com_progid)
            break
        except Exception as exc:
            message = f"{backend_label}({com_progid}) failed: {exc}. {mathtype_troubleshooting_hint()}"
            errors.append(message)
            print(f"backend_failed={message}")
            cleanup_leftovers(args)
            wait_for_exclusive_file(output, timeout=float(args.window_timeout))
    if used_backend is None:
        raise RuntimeError("All requested Word/WPS backends failed: " + " | ".join(errors))
    wait_for_exclusive_file(output, timeout=float(args.window_timeout))

    info = inspect_docx_path(output, paragraph_index)
    print_inspection(info)
    print_size_info(size_info)
    if args.validate_wps:
        print(f"wps_open={validate_wps_open(output)}")
    if args.reopen_screenshot:
        capture_saved_ole_screenshot(
            output,
            paragraph_index,
            Path(args.reopen_screenshot).resolve(),
            com_progid=used_backend[1],
            window_timeout=float(args.window_timeout),
        )
        print(f"reopen_screenshot={Path(args.reopen_screenshot).resolve()}")


def resize_docx_ole(args: argparse.Namespace) -> None:
    """Resize an existing MathType OLE object by matching manuscript body font scale."""

    docx = Path(args.docx).resolve()
    if not docx.exists():
        raise FileNotFoundError(docx)
    output = Path(args.output).resolve() if args.output else docx
    if output == docx and args.backup:
        backup = write_backup(docx, "formula_ole_resize")
        print(f"backup={backup}")
    if output == docx:
        wait_for_exclusive_file(output, timeout=float(args.window_timeout))
    else:
        output.parent.mkdir(parents=True, exist_ok=True)

    paragraph_index = int(args.paragraph_index)
    if args.one_based:
        paragraph_index -= 1
    if paragraph_index < 0:
        raise ValueError("paragraph index must be positive after index-base conversion")

    size_info = resolve_formula_object_size(
        docx,
        paragraph_index,
        "",
        explicit_height_pt=args.height_pt,
        explicit_body_font_pt=args.body_font_pt,
        formula_font_scale=args.formula_font_scale,
        ole_line_height_factor=args.ole_line_height_factor,
        formula_lines=args.formula_lines,
    )

    with ZipFile(docx, "r") as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        paras = root.findall(".//w:body/w:p", namespaces=NS)
        if paragraph_index >= len(paras):
            raise IndexError(f"paragraph index out of range: {paragraph_index} / {len(paras)}")
        p = paras[paragraph_index]
        shapes = p.findall(".//v:shape", namespaces=NS)
        if not shapes:
            raise RuntimeError("No VML shape found in the target paragraph; is this a MathType OLE paragraph?")

        shape = shapes[0]
        style = shape.get("style", "")
        dims = parse_css_pt(style)
        old_height = dims.get("height")
        old_width = dims.get("width")
        new_height = float(size_info["resolved_height_pt"])
        if args.width_pt is not None:
            new_width = float(args.width_pt)
        elif old_width is not None and old_height and old_height > 0:
            new_width = float(old_width) * new_height / float(old_height)
        else:
            new_width = old_width

        style = set_css_pt(style, "height", new_height)
        if new_width is not None:
            style = set_css_pt(style, "width", float(new_width))
        shape.set("style", style)

        document_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        copy_docx_with_updates(docx, output, document_xml)

    info = inspect_docx_path(output, paragraph_index)
    print_inspection(info)
    print_size_info(size_info)
    if args.validate_wps:
        print(f"wps_open={validate_wps_open(output)}")


def inspect_docx(args: argparse.Namespace) -> None:
    paragraph_index = int(args.paragraph_index)
    if args.one_based:
        paragraph_index -= 1
    info = inspect_docx_path(Path(args.docx).resolve(), paragraph_index)
    print_inspection(info)
    if args.validate_wps:
        print(f"wps_open={validate_wps_open(Path(args.docx).resolve())}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    env = sub.add_parser("check-env", help="Check local MathType, WPS/Word COM, and Python dependencies.")
    env.add_argument("--mathtype-exe", default="", help="Optional explicit MathType.exe path.")
    env.add_argument("--ole-class-type", default="Equation.DSMT4", help="OLE class used for MathType equations.")
    env.add_argument("--probe-com", action="store_true", help="Actually start and quit WPS/Word COM apps in hidden/background mode.")
    env.add_argument("--probe-mathtype", action="store_true", help="Launch MathType briefly to detect runtime or activation/license problems.")
    env.set_defaults(func=check_env)

    clean = sub.add_parser("cleanup-leftovers", help="Stop leftover MathType/MathTypeLib background processes.")
    clean.set_defaults(func=cleanup_leftovers)

    make = sub.add_parser("make-wmf", help="Create WMF/EMF from MathType-rendered MathML or LaTeX.")
    make.add_argument("--mathml-file")
    make.add_argument("--mathml")
    make.add_argument("--latex-file")
    make.add_argument("--latex")
    make.add_argument("--output-dir", required=True)
    make.add_argument("--base-name", required=True)
    make.add_argument("--mathtype-exe", default="", help="Optional MathType.exe path. Falls back to MATHTYPE_EXE, registry, common paths, then PATH.")
    make.add_argument("--wait-seconds", type=float, default=1.2)
    make.add_argument("--keep-window", action="store_true")
    make.set_defaults(func=make_wmf)

    rep = sub.add_parser("replace-docx", help="Insert or resize a centered formula image with right equation number.")
    rep.add_argument("--docx", required=True)
    rep.add_argument("--output", default="")
    rep.add_argument("--paragraph-index", required=True, type=int, help="OpenXML zero-based paragraph index by default.")
    rep.add_argument("--one-based", action="store_true", help="Treat paragraph-index as Word/WPS one-based.")
    rep.add_argument("--image", help="WMF/EMF image to insert. Omit when using --scale-existing.")
    rep.add_argument("--number", default="(1)")
    rep.add_argument("--height-pt", type=float)
    rep.add_argument("--width-pt", type=float)
    rep.add_argument("--scale-existing", type=float, help="Scale the current paragraph image, e.g. 0.8 for 20 percent smaller.")
    rep.add_argument("--body-font-pt", type=float, help="Override detected manuscript body font size in points.")
    rep.add_argument("--formula-font-scale", type=float, default=0.8, help="Formula font target as body-font multiple. Default: 0.8.")
    rep.add_argument("--image-line-height-factor", type=float, default=2.1, help="Image formula height per formula font point per line. Default: 2.1.")
    rep.add_argument("--formula-lines", type=int, default=1, help="Formula line count for image fallback auto-size. Default: 1.")
    rep.add_argument("--font-size-half-points", type=int, default=21)
    rep.add_argument("--backup", action="store_true")
    rep.add_argument("--validate-wps", action="store_true")
    rep.set_defaults(func=replace_docx)

    ole = sub.add_parser("replace-docx-ole", help="Replace a formula paragraph with a real MathType Equation.DSMT4 OLE object.")
    ole.add_argument("--docx", required=True)
    ole.add_argument("--output", default="")
    ole.add_argument("--paragraph-index", required=True, type=int, help="OpenXML zero-based paragraph index by default.")
    ole.add_argument("--one-based", action="store_true", help="Treat paragraph-index as Word/WPS one-based.")
    ole.add_argument("--mathml-file")
    ole.add_argument("--mathml")
    ole.add_argument("--latex-file")
    ole.add_argument("--latex")
    ole.add_argument("--number", default="(1)")
    ole.add_argument("--height-pt", type=float, help="Override OLE object height in points. Default: auto from body font.")
    ole.add_argument("--body-font-pt", type=float, help="Override detected manuscript body font size in points.")
    ole.add_argument("--formula-font-scale", type=float, default=0.8, help="Formula font target as body-font multiple. Default: 0.8.")
    ole.add_argument("--ole-line-height-factor", type=float, default=2.1, help="OLE visual frame height per formula font point per line. Default: 2.1.")
    ole.add_argument("--formula-lines", type=int, default=0, help="Override formula line count. Default: infer from MathML mtable rows.")
    ole.add_argument("--center-tab-twips", type=int, default=4649)
    ole.add_argument("--right-tab-twips", type=int, default=9298)
    ole.add_argument("--editor-screenshot", default="")
    ole.add_argument("--backend", choices=("auto", "wps", "word"), default="auto", help="Document COM backend. Default: auto (WPS first, then Word).")
    ole.add_argument("--com-progid", default="", help="Override COM ProgID, e.g. Kwps.Application or Word.Application.")
    ole.add_argument("--ole-class-type", default="Equation.DSMT4", help="OLE class type for MathType equations.")
    ole.add_argument("--window-timeout", type=float, default=20.0)
    ole.add_argument("--after-paste-wait", type=float, default=2.0)
    ole.add_argument("--after-update-wait", type=float, default=3.0)
    ole.add_argument("--reopen-screenshot", default="")
    ole.add_argument("--visible-app", action="store_true", help="Deprecated/no-op: Word/WPS COM is always opened hidden.")
    ole.add_argument("--foreground-editor", action="store_true", help="Bring MathType editor to foreground. Word/WPS stays hidden.")
    ole.add_argument("--cleanup-mathtype", action="store_true")
    ole.add_argument("--backup", action="store_true")
    ole.add_argument("--validate-wps", action="store_true")
    ole.set_defaults(func=replace_docx_ole)

    resize_ole = sub.add_parser("resize-docx-ole", help="Resize an existing MathType OLE object against manuscript body font scale.")
    resize_ole.add_argument("--docx", required=True)
    resize_ole.add_argument("--output", default="")
    resize_ole.add_argument("--paragraph-index", required=True, type=int, help="OpenXML zero-based paragraph index by default.")
    resize_ole.add_argument("--one-based", action="store_true", help="Treat paragraph-index as Word/WPS one-based.")
    resize_ole.add_argument("--height-pt", type=float, help="Override OLE object height in points. Default: auto from body font.")
    resize_ole.add_argument("--width-pt", type=float, help="Override OLE object width in points. Default: preserve current aspect ratio.")
    resize_ole.add_argument("--body-font-pt", type=float, help="Override detected manuscript body font size in points.")
    resize_ole.add_argument("--formula-font-scale", type=float, default=0.8, help="Formula font target as body-font multiple. Default: 0.8.")
    resize_ole.add_argument("--ole-line-height-factor", type=float, default=2.1, help="OLE visual frame height per formula font point per line. Default: 2.1.")
    resize_ole.add_argument("--formula-lines", type=int, default=1, help="Formula line count for existing OLE objects. Default: 1; use 2 for two-line formulas.")
    resize_ole.add_argument("--window-timeout", type=float, default=20.0)
    resize_ole.add_argument("--backup", action="store_true")
    resize_ole.add_argument("--validate-wps", action="store_true")
    resize_ole.set_defaults(func=resize_docx_ole)

    ins = sub.add_parser("inspect-docx", help="Inspect formula paragraph structure and optional WPS openability.")
    ins.add_argument("--docx", required=True)
    ins.add_argument("--paragraph-index", required=True, type=int)
    ins.add_argument("--one-based", action="store_true")
    ins.add_argument("--validate-wps", action="store_true")
    ins.set_defaults(func=inspect_docx)
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
