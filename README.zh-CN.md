# MathType-Word/WPS

中文说明 | [English](README.md)

在多模态大模型辅助识别/转写公式图片的前提下，MathType-Word/WPS 提供从公式图片到可编辑 MathType 公式、自动插入 Word/WPS 文档、自适应手稿布局、结果校验与关联公式集完整性审查的全流程能力。

这个工具面向论文、学位论文、报告等正式文档场景：公式不能只是截图，插入后应该能继续双击用 MathType 编辑，字号和正文协调，编号仍然是文档里的普通右对齐文本，后续排版、修改、审稿都更稳。它会自动处理很多容易踩坑的细节，包括后台打开 Word/WPS、创建真实 MathType 公式对象、按正文字号设置 MathType 内部公式主字符大小和 Times New Roman 字体族、再匹配 OLE 外框尺寸但不把复杂公式压缩到低于 MathType 自然高度、保留公式编号、检查保存后的文档结构，并区分“真正可编辑的 MathType 对象”和“只是图片或 OMML 公式”的情况。

主要目标格式是 `.docx`。旧版 `.doc` 也可以通过 Word/WPS 自动化流程处理；如果需要做 XML 级检查，建议先用 Word/WPS 打开或转换为 `.docx`。

## 独立分发与使用

本仓库是一套完整、可独立安装的 skill，随包提供公式集 schema、计算图审计器、测试、MathType OLE 实现、DOCX 包检查和验收规则。它不会导入或依赖 `evidence-grounded-manuscript-skills`、`research-paper-writing` 或 `paper-review-audit`。

其它论文写作或审查 skill 可以把同一份清单格式作为可选互操作接口，但不是安装或运行前提。用户只安装 MathType-Word/WPS，也可以完成公式集审查、可编辑公式插入、DOCX 包检查和渲染验证；所需条件仅为下文明确列出的 Windows、Word/WPS、MathType 和 Python 环境。

## 支持功能

- 创建真实可编辑的 MathType OLE 公式对象：`Equation.DSMT4`。这是 Word/WPS 能重新调用 MathType 打开并编辑的公式对象类型。
- 自动选择 WPS 或 Microsoft Word COM 后端。
- 公式编号保持为普通文档文本，并放在右侧对齐。
- 根据正文默认字号设置 MathType 内部公式主字符大小和 Times New Roman 字体族，再匹配 OLE 外框尺寸，避免只缩放外框造成公式视觉不统一。
- 检查 DOCX XML、MathType 原生 OLE 流和缓存 WMF/EMF 预览，确认公式是否为 OLE 对象、图片、OMML 或其它形式。
- 批量修改前验证段落定位；Word/WPS COM 段落索引不稳定时，支持以单段公式资产完成受控重建与包级移植。
- 将公式集清单视为有向计算图，检查未定义或先用后定义符号、孤立输出、重复定义、实例接口缺项、废弃符号和可简化别名候选。
- 维护公式源、正文、代码、图片、OLE 部件和预览部件之间的对应关系。
- 在明确需要时，可使用 WMF/EMF 矢量图片作为后备方案。

## 在“图片转可编辑公式”流程中的位置

这个项目本身不负责从截图里 OCR 识别公式。它介入的位置是：公式图片已经通过外部识别工具、OCR 服务、LLM 辅助转写或人工整理，变成 MathML 或 LaTeX 之后。

典型流程是：

1. 输入是一张公式截图、扫描公式，或手稿里原本嵌入的公式图片。
2. 先用外部识别/转写步骤得到 MathML 或 LaTeX。
3. 再用本工具把原来的图片公式替换成 Word/WPS 中真正可编辑的 MathType OLE 对象：`Equation.DSMT4`。
4. 工具负责保留右侧公式编号、匹配正文字号、调整公式大小，并检查保存后的 DOCX，确认结果是可编辑 OLE，而不是另一张图片。

简单说就是：`图片 -> 外部识别/人工转写 -> MathML/LaTeX -> 本工具 -> Word/WPS 里的可编辑 MathType 公式`。

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

如果怀疑公式不可编辑、双击后像一个整体对象，或者正文预览中出现问号，先把源 MathML 转成 MathType 原生 `MathType EF/MTEF`。这一步不走“打开 MathType 编辑器再粘贴 MathML”的脆弱路径，而是直接调用 MathType 原生 DataObject：

```powershell
python .\scripts\mathtype_word_wps.py native-bridge-mtef `
  --mathml-file ".\formula11.mathml" `
  --output-mtef ".\formula11.mtef"
```

原生桥每次只做一次有超时的转换，结束后会清理 MathType embedding 后台进程。不要再用批量 GUI 探测、宏循环或反复 `Application.Run` 去试公式。

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

插入完成后，工具会清理 MathType/MathTypeLib 以及安全范围内的 WPS 预览/嵌入 helper 进程，并等待目标文档重新变成可写状态。这个检查很重要：公式本身可能已经可编辑、显示也正确，但残留的 OLE 预览 helper 仍可能让用户在前台编辑器里无法正常输入。

尺寸规则：`--formula-font-scale 0.8` 控制的是 MathType 内部主字符字号，不只是外层 OLE 框大小；默认公式字体族为 `Times New Roman`。遇到分式、大型求和、多行对齐等复杂公式时，OLE 外框只作为最小容器，默认不会把对象压缩到低于 MathType 自然插入高度，除非显式传入 `--allow-downscale-ole`。粘贴前 MathML 会先做 MathType/Word 预览稳定化，再转换为 ASCII 数字实体：例如标量乘法中点从 U+00B7 归一为点乘 U+22C5，卷积核尺寸或形状中的乘号可归一为 ASCII `x`，避免嵌入对象可编辑但 Word/WPS 正文预览显示 `?`。

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
- `ole_analysis` 中存在 `Equation Native` 原生流
- 缓存 WMF/EMF 预览中没有明显 `preview_cache_question_marks` 风险

整篇手稿可以用离线审计，不会打开 Word 或 MathType：

```powershell
python .\scripts\mathtype_word_wps.py audit-docx-formulas `
  --docx ".\manuscript.docx" `
  --output-json ".\formula_audit.json" `
  --include-details `
  --min-ole-width-pt 18 `
  --min-ole-height-pt 8
```

尺寸阈值只用于提示公式对象可能接近空白。合法的单字符小公式可把相应阈值设为 `0`；批量生成资产时，也可以根据已知的最小合理尺寸提高阈值。

批量替换前应先在手稿副本中测试一个公式，确认 OLE 落在预期编号旁。OpenXML 正文索引与 Word/WPS COM 段落索引可能并不一致。探针落点错误时，应改为在单段 DOCX 中逐式生成并验证公式，再把完整 OLE 关系、原生对象和缓存预览移植到精确的 OpenXML 标记位置。正式替换前必须维护公式编号与源文件台账，并审计组装后的候选手稿。

注意区分两层问题：MathType OLE 原生内容和文档里显示出来的 WMF/EMF 预览不是一回事。如果正文里有问号，但双击进入 MathType 后内容正常，优先用归一化后的 MathML/MTEF 重新生成缓存预览；如果双击进入 MathType 后公式本身就是一个整体、无法逐个选中字母数字，那必须从 MathML/LaTeX 重新生成原生 MTEF 或在 MathType 里重建，单纯调整 OLE 外框大小无效。`preview_cache_question_marks` 只能作为风险提示，不是最终视觉结论；WMF/EMF 字节扫描可能误报，最终验收必须导出 Word/PDF 页面截图确认正文层是否还有可见问号。

如果运行后用户无法在手稿里输入文本，第一步就是运行 `cleanup-leftovers`，然后确认 `.docx` 能被独占读写打开。只要隐藏的 OLE、MathType、Word 或 WPS 预览 helper 仍在影响前台编辑，就不能把公式任务判定为完成。

如果不是某一个文档不能输入，而是浏览器、Codex、WPS 等任何输入框都不能输入，就不要继续按文档锁排查。这通常是 Windows 输入法/文本服务卡住，尤其是 `ctfmon`、`TextInputHost`、第三方输入法 hook，或隐藏的 Office OLE 自动化进程例如 `VISIO.EXE /Automation -Embedding` 把输入法拉进了自己的进程树。此时运行：

```powershell
python .\scripts\mathtype_word_wps.py recover-text-input --stop-wetype
```

该命令会重启 Windows 文本输入链，并在需要时停止腾讯 WeType 相关进程。

## 审查关联公式集

当多条公式共同构成一个方法、推导链或共享接口时，先建立符合 `schemas/formula_set_manifest.schema.json` 的 JSON 清单。每个公式记录稳定 ID、顺序、科学用途、定义量、依赖量、源文件，以及可选的代码、正文、图片和 OLE 锚点。

插入前运行：

```powershell
python .\scripts\audit_formula_set.py ".\formula_set.json" `
  --require-source `
  --require-prose-anchor `
  --output-json ".\formula_set_audit.json"
```

DOCX 组装后，补充公式编号、段落锚点、OLE 包部件和预览部件，再使用 `--require-artifact` 复查。

审查器会把公式集作为有向计算图，报告未定义量、先用后定义、孤立输出、重复定义、循环依赖、共享实例输出缺项、废弃符号和单消费者别名候选。别名结果只是一项建议：变量若具有独立物理意义、承担量纲或坐标变换、参与证明或极限关系，或属于共享接口，就应保留。

完整的语义层、公式对象层和原生渲染层验收规则见 `references/formula-set-integrity.md`。语义 PASS 不代表 MathType 对象正确；合法 OLE 外壳也不代表公式集已经在科学上闭环。

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
git clone https://github.com/pcdp577/MathType-Word-WPS.git "$env:USERPROFILE\.codex\skills\MathType-Word-WPS"
```

安装后重启 Codex。

## 开发验证

```powershell
python -X utf8 .\scripts\test_formula_set_audit.py
python -X utf8 .\scripts\test_standalone_install.py
python -X utf8 -m py_compile .\scripts\audit_formula_set.py .\scripts\mathtype_word_wps.py
python -X utf8 <skill-creator>\scripts\quick_validate.py .
```
