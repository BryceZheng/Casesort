# Casesort — 通用案例库整理工具

五阶段法律案例库整理工具，适用于在威科先行、企查查、北大法宝等法律数据库下载 PDF 后的全流程自动化处理。可作为 AI 助手的指令文件（Claude Code、Codex 等支持自定义 Skill/Agent 的系统），也可直接运行 Python 脚本独立使用。

## 功能概览

| 阶段 | 名称 | 输入 | 输出 |
|------|------|------|------|
| 第一阶段 | 检索关键词生成 | 用户描述案例类型 | 结构化检索式（AND/OR/NOT） |
| 第二阶段 | PDF 批量识别 | PDF 文件夹 | Markdown 文件夹 |
| 第三阶段 | AI 全量核查 | Markdown 文件夹 | 3色核查 Excel + 分类文件夹 |
| 第四阶段 | 结构化提取 | 确认纳入案例文件夹 | 案例库 Excel + MD 文件重命名 |
| 第五阶段 | 报告产出 | 案例库 Excel | 法律分析报告 Word 文档 |

每个阶段结束后询问是否继续，支持从任意阶段单独启动。

## 适用场景

- 行政处罚案例库整理（商业贿赂、虚假宣传、食品安全等）
- 刑事判决案例整理
- 任意需要从大量 PDF 文件中批量提取、筛选、分析案例的场景

## 安装方法

### 系统兼容性

| 系统 | 支持情况 |
|------|---------|
| macOS | ✅ 完整支持 |
| Linux | ✅ 完整支持 |
| Windows | ⚠️ 需安装 [WSL](https://learn.microsoft.com/zh-cn/windows/wsl/install)（Windows Subsystem for Linux），安装后按 Linux 步骤操作 |

### 前置要求

- Python 3.7+
- 以下 Python 依赖：

```bash
pip install openpyxl python-docx requests pdfplumber
```

- PDF 转换依赖（根据使用的 OCR 方式选一种）：

**macOS：**
```bash
# 方式A：pdftotext
brew install poppler

# 方式B：Tesseract
brew install tesseract tesseract-lang
```

**Linux（Ubuntu/Debian）：**
```bash
# 方式A：pdftotext
sudo apt-get install poppler-utils

# 方式B：Tesseract
sudo apt-get install tesseract-ocr tesseract-ocr-chi-sim poppler-utils
```

方式C（PaddleOCR）无需本地安装，但需要 [AI Studio](https://aistudio.baidu.com/) API Token。

### 方式一：作为 AI Agent 的指令文件（推荐）

`SKILL.md` 是本工具的核心指令文件，可加载到支持自定义指令的 AI 助手中使用。

**Claude Code（Skills 系统）：**
```bash
cp -r Casesort ~/.claude/skills/
```
重启后输入触发词即可启动。

**其他 AI Agent / Codex 等：**
将 `SKILL.md` 的内容作为系统提示词（System Prompt）或自定义指令加载到对应平台，按各平台文档配置即可。

### 方式二：直接运行 Python 脚本

无需 AI 环境，直接调用脚本处理商业贿赂案例（旧版流程）：

```bash
# 交互式启动
python3 run.py

# 直接指定路径
python3 extract_cases.py "输入目录" "输出.xlsx"
```

## 使用方法

### 通过 AI Agent

加载 `SKILL.md` 后，向 AI 输入以下任意触发词：

> `整理案例` / `筛选案例` / `案例库` / `检索关键词` / `全量核查` / `案例分析` / `Casesort`

AI 会询问从哪个阶段开始，然后逐步引导完成。

### 直接运行脚本

参见上方「方式二」，或查阅 `QUICKSTART.txt`。

## 各阶段说明

### 第一阶段：检索关键词生成

AI 根据你描述的案例类型（法律依据、时间范围、特殊要求），生成结构化的数据库检索式，包含 AND / OR / 排除词，并按法律或场景分组，可直接复制到威科先行等数据库使用。

### 第二阶段：PDF 批量识别

支持三种 OCR 方式（pdftotext / Tesseract / PaddleOCR），将 PDF 文件夹下的所有文件批量转为 Markdown，输出到 `Markdown_转换结果/` 子文件夹。

### 第三阶段：AI 全量核查

- 读取前 30 个文件做预分析，识别混淆类型，确认纳入标准后再批量处理
- 每 20 个文件一组，经 3 轮验证（初判 → 审查边界 → 深度核验疑似）
- 输出 3 色 Excel（蓝=确认纳入 / 橙=疑似 / 红=确认排除）
- 自动将文件按结论分类复制到 `全量核查筛选案例/` 文件夹

### 第四阶段：结构化字段提取

两种输出格式：

**格式 A（商业贿赂专用，17 字段）**：含穿透式定性规则、2019修正/2025修订法条版本自动识别、行业分类（21类）、地域推断（120+ 城市映射）

**格式 B（通用，10 基础字段 + AI 推荐额外字段）**：适用于其他类型行政处罚、刑事案件等

每 10 个文件一组，经 2 轮复核后写入 Excel，最后对 MD 文件按序号加前缀重命名。

### 第五阶段：法律分析报告

基于案例库 Excel，AI 自动分类分析，生成包含数据概览、分类分析（附查询表格）、监管执法关注点、企业合规建议的 Word 文档。

## 文件结构

```
Casesort/
├── SKILL.md              # AI Agent 核心指令文件（五阶段流程说明）
├── CHANGELOG.md          # 版本更新日志
├── QUICKSTART.txt        # 快速上手指引
├── phase3_excel.py       # 第三阶段：全量核查 Excel 生成
├── phase4_excel.py       # 第四阶段：成果产出 Excel 生成
├── phase5_report.py      # 第五阶段：法律分析报告 Word 生成
├── paddle_ocr.py         # 第二阶段方式C：PaddleOCR 批量识别
├── extract_bribery_cases.py  # 旧版：商业贿赂 PDF 提取
├── extract_cases.py          # 旧版：docx 提取
├── deep_verify.py            # 旧版：深度核验（构成要件分析）
├── verify_cases.py           # 旧版：三层核查
├── phase2_manual.py          # 旧版：手动 PDF 处理
└── run.py                    # 旧版：交互式启动入口
```

> 旧版文件（`extract_*`、`verify_*`、`run.py`）保留向后兼容，商业贿赂历史流程仍可使用。

## 配合使用的工具

以下为 Claude Code Skills 系统中的可选配套工具，其他平台用户可用等效方式替代：

- **PaddleOCR**（第二阶段可选）：用于将扫描件 PDF 转为文字，需自行申请 API Token
- **docx 工具**（第五阶段可选）：用于生成格式化 Word 报告；非 Claude Code 用户可直接使用 `python-docx` 手动生成

## 版本历史

| 版本 | 日期 | 主要变更 |
|------|------|----------|
| v3.0 | 2026-06-27 | 通用化升级：两阶段 → 五阶段，适用任意案例类型 |
| v2.1 | 2026-06-17 | 假阳性过滤优化 |
| v2.0 | 2026-06-17 | 新增 PDF 直接处理流程 |
| v1.4 | 2026-06-11 | 企查查数据处理，法条版本 Bug 修复 |
| v1.3 | 2026-06-11 | 字段质量全面提升（行业/地域/法律依据/主要事实） |
| v1.2 | 2026-06-11 | 新增深度核验（构成要件分析，置信度评分） |
| v1.0 | 2026-06-11 | 初始版本（CBcaseSum） |

## 注意事项

1. **安装路径**：脚本通过 `find` 动态定位，安装到任意目录均可，无需固定路径
2. **第二阶段方式C（PaddleOCR）**：需要用户自备 PaddleOCR API Token，未申请时请选择方式A（pdftotext）或方式B（Tesseract）
3. **第五阶段 Word 报告**：由内置的 `phase5_report.py` 生成，无需额外依赖
4. **PaddleOCR API Key**：如使用 PaddleOCR 方式，需在调用命令中传入自己的 API Token，请勿将 Token 写入代码后上传到公开仓库
5. **文件格式**：第四阶段旧版脚本仅支持 `.docx` 格式；第三/四阶段新流程基于 Markdown 文件
6. **Excel 输出**：需要 `openpyxl` 库，生成的 `.xlsx` 文件不建议上传到版本控制

## 许可证

MIT License
