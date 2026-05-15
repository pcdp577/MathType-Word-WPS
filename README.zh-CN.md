# MathType-Word/WPS

中文说明 | [English](README.md)

通过 WPS 或 Microsoft Word，把可编辑的 MathType 公式插入 Word 手稿。

这个工具面向论文、学位论文、报告等正式文档场景：公式不能只是截图，插入后应该能继续双击用 MathType 编辑，字号和正文协调，编号仍然是文档里的普通右对齐文本，后续排版、修改、审稿都更稳。它会自动处理很多容易踩坑的细节，包括后台打开 Word/WPS、创建真实 MathType 公式对象、按正文字号缩放公式、保留公式编号、检查保存后的文档结构，并区分“真正可编辑的 MathType 对象”和“只是图片或 OMML 公式”的情况。

主要目标格式是 `.docx`。旧版 `.doc` 也可以通过 Word/WPS 自动化流程处理；如果需要做 XML 级检查，建议先用 Word/WPS 打开或转换为 `.docx`。

## 支持功能

- 创建真实可编辑的 MathType OLE 公式对象：`Equation.DSMT4`。这是 Word/WPS 能重新调用 MathType 打开并编辑的公式对象类型。
- 自动选择 WPS 或 Microsoft Word COM 后端。
- 公式编号保持为普通文档文本，并放在右侧对齐。
- 根据正文默认字号调整公式整体大小，避免公式过大或过小。
- 检查 DOCX XML，确认公式是否为 OLE 对象、图片、OMML 或其它形式。
- 在明确需要时，可使用 WMF/EMF 矢量图片作为后备方案。

## 环境要求

- Windows 桌面会话。
- Python 3.10 或更新版本。
- Microsoft Word 或 WPS Office。
- 已安装并激活 MathType，推荐 MathType 7.0；标准激活安装通常会自动注册可编辑的 `Equation.DSMT4` OLE/COM 对象，以支持真实 OLE 公式插入。

安装 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

`latex2mathml` 是可选依赖，只在使用 `--latex` 输入时需要：

```powershell
python -m pip install latex2mathml
```

## 快速检查

```powershell
python .\scripts\mathtype_word_wps.py check-env
```

检查 Word/WPS COM 后端：

```powershell
python .\scripts\mathtype_word_wps.py check-env --probe-com
```

检查 MathType 运行与激活状态：

```powershell
python .\scripts\mathtype_word_wps.py check-env --probe-mathtype
```

如果找不到 MathType，可以设置：

```powershell
$env:MATHTYPE_EXE = "C:\Path\To\MathType.exe"
```

也可以在命令中传入 `--mathtype-exe`。

如果 `Equation.DSMT4` 缺失，Word/WPS 就无法创建真正可编辑的 MathType 对象。此时应先确认 MathType 已激活；如果已经激活，尝试修复安装或重新安装 MathType，让 Windows 重新注册对应的 OLE/COM 对象。

## 插入可编辑公式

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

默认情况下，`--backend auto` 会先尝试 WPS，失败后再尝试 Word。也可以使用 `--backend word`、`--backend wps` 或 `--com-progid ...` 强制指定后端并关闭自动回退。

正常处理时，Word 和 WPS 都会通过 COM 后台隐藏打开，不应该弹出前台窗口。

## 检查插入结果

```powershell
python .\scripts\mathtype_word_wps.py inspect-docx `
  --docx ".\manuscript.docx" `
  --paragraph-index 49
```

真正的 MathType OLE 公式通常应显示：

- `ole_objects=1`
- `ole_progids=['Equation.DSMT4']`
- `blips=0`
- `omath=0`

## 在 Cursor、Claude Code 或其它 Agent 中使用

这个仓库以 Codex skill 的形式发布，但核心实现是普通 Windows Python CLI。只要其它客户端能在同一台安装了 Word/WPS 和 MathType 的 Windows 桌面上运行 PowerShell 命令，就可以直接使用。

对 Cursor、Claude Code 等客户端的注意事项：

- 直接调用 `scripts/mathtype_word_wps.py` 即可；除非客户端自己实现了 skill 加载机制，否则不会自动识别 Codex skill。
- 建议把 `SKILL.md` 里的关键规则复制到对应客户端的项目规则或提示词中，尤其是 OLE 优先、后台打开 Word/WPS、公式尺寸匹配和 XML 检查这些规则。
- 必须在真实 Windows 桌面会话中运行，不能在 WSL、远程 Linux shell 或无界面容器中运行。MathType OLE 插入依赖 Windows COM 自动化。
- 离线替换文档时，目标文档最好不要被其它 Word/WPS 进程打开；如果必须复用已打开的应用会话，要避免重复打开导致文档变成只读。
- 编辑正式手稿前，先运行 `check-env --probe-com --probe-mathtype`，再做一个小公式替换测试。

目前本仓库只完整验证了 Codex 工作流。Cursor 和 Claude Code 可以通过 CLI 路径使用，但在完成环境检查和小样本替换前，应视为“设计上兼容”，而不是已经逐客户端完整验证。

## 作为 Codex Skill 安装

把仓库克隆或复制到 Codex skills 目录：

```powershell
git clone https://github.com/liuyifan577/MathType-Word-WPS.git "$env:USERPROFILE\.codex\skills\MathType-Word-WPS"
```

安装后重启 Codex。
