---
name: Casesort
description: 通用案例库整理工具（五阶段）。第一阶段生成检索关键词；第二阶段将PDF批量转为Markdown；第三阶段AI驱动全量核查；第四阶段结构化字段提取生成案例库Excel；第五阶段撰写法律分析报告。触发词：整理案例、筛选案例、案例库、检索关键词、全量核查、案例分析、生成案例、处理案例、案例整理、Casesort。
---

# Casesort — 通用案例库整理工具（v3.0）

> **安装说明**：第三、四、五阶段需要调用 Python 脚本，脚本路径通过 `find` 动态定位，安装到任意目录均可正常使用。
>
> **功能限制说明**：
> - 第二阶段方式C（PaddleOCR）需要用户自备 PaddleOCR API Token，未申请时请选择方式A或方式B
> - 第五阶段 Word 报告由内置的 `phase5_report.py` 生成，无需额外依赖

## 工具概述

本工具分为**五个独立阶段**，每个阶段结束后询问用户是否继续下一阶段。

| 阶段 | 名称 | 输入 | 输出 |
|------|------|------|------|
| **第一阶段** | 检索关键词生成 | 用户描述案例类型 | 结构化检索式（对话框展示） |
| **第二阶段** | PDF 识别 | PDF 文件夹 | Markdown 文件夹 |
| **第三阶段** | 全量核查 | Markdown 文件夹 | 3色核查 Excel + 分类文件夹 |
| **第四阶段** | 成果产出 | 确认纳入案例文件夹 | 案例库 Excel + MD 文件重命名 |
| **第五阶段** | 报告产出 | 案例库 Excel | 法律分析报告 Word 文档 |

收到用户指令后，**先询问**要从哪个阶段开始执行。

---

## 第一阶段：检索关键词生成

### 目标

帮助用户构建专业、结构化的数据库检索式，用于在威科先行、企查查、北大法宝等法律数据库中检索目标案例。

### 执行步骤

**步骤 1：询问需求**

向用户提问，收集以下信息：
- 想要整理的**案例类型**（如商业贿赂行政处罚、非法吸收公众存款刑事案件、虚假宣传行政处罚等）
- 涉及的**具体法条**（如《反不正当竞争法》第8条、《刑法》第176条等）
- **时间范围**（如2026年、2024-2025年等）
- **其他特殊要求**（如特定行业、地域、金额门槛等）

**步骤 2：生成检索式**

根据用户描述，在对话框中直接输出结构化检索式：

**输出格式规范：**

- 将检索式拆分为若干**独立检索组**，每组对应一部法律或一类场景，用户可分组分次检索
- 每组内分 AND（必须同时出现）和 OR（任一命中即可）两层
- **AND 关键词须覆盖同义/近义表达**，避免因用词差异遗漏案例（如"商业贿赂"可能在文书中写作"贿赂"或"好处费"）
- **OR 行须同时列出汉字版和阿拉伯数字版条款**（如"第七条 第 7 条"），适配不同数据库的录入规范
- 不建议将"行政处罚"作为 AND 词，因部分文书正文不出现该词
- 排除词标注**（可选）**，用户自行判断是否输入

```
【建议检索式】

适用数据库：威科先行 / 企查查 / 北大法宝（根据案例类型推荐）
检索时间范围：[用户指定范围]
文书类型筛选：行政处罚决定书（在数据库筛选项中选择，无需写入检索词）
检索范围建议：同段（若结果过少改为全文）

第一组：[法律名称或场景描述]
  AND：关键词A   关键词B   关键词C
  OR ：第X条   第 N 条   第Y条   第 M 条

第二组：[法律名称或场景描述]
  AND：关键词D   关键词E   关键词F
  OR ：第X条   第 N 条

排除词（可选）：排除词X   排除词Y

【说明】
- 每组独立检索后合并去重
- 先用"同段"测试；结果过少时改"全文"
- OR 行同时列汉字版和数字版条款，适配不同数据库录入格式
```

**步骤 3：确认并调整**

询问用户：检索式是否满足需求？是否需要调整关键词或范围？用户确认后结束第一阶段，提示可前往数据库下载 PDF 后进入第二阶段。

---

## 第二阶段：PDF 识别

### 目标

将用户下载的 PDF 文件批量转换为 Markdown 格式，为后续 AI 分析做准备。

### 执行步骤

**步骤 1：询问路径**

> "请提供 PDF 文件所在的文件夹路径。"

**步骤 2：询问 OCR 方式**

> "请选择 OCR 识别方式：
> - **pdftotext**（极快，秒级完成，仅适合含文字层的 PDF，扫描件无效）
> - **Tesseract**（本地 OCR，适合扫描件，分钟级，需安装 tesseract + poppler）
> - **PaddleOCR**（云端 API，布局最准确，适合复杂版式/表格，5-30秒/文件，需联网）"

**步骤 3：检查依赖**

根据选择检查依赖：

```bash
# pdftotext
which pdftotext || echo "未安装，请运行: brew install poppler"

# Tesseract
tesseract --version && tesseract --list-langs 2>&1 | grep chi || echo "未安装，请运行: brew install tesseract tesseract-lang"
which pdftoppm || echo "未安装，请运行: brew install poppler"

# PaddleOCR（无需本地依赖，调用云端 API）
python3 -c "import requests" || pip3 install requests
```

**步骤 4：创建输出文件夹**

在 PDF 所在目录下创建子文件夹：
```bash
mkdir -p "<PDF路径>/Markdown_转换结果"
```

**步骤 5：批量转换**

根据用户选择执行对应命令：

**方式 A：pdftotext**
```bash
INPUT_DIR="<用户提供的PDF路径>"
OUTPUT_DIR="${INPUT_DIR}/Markdown_转换结果"
mkdir -p "$OUTPUT_DIR"
for PDF in "$INPUT_DIR"/*.pdf; do
    BASENAME=$(basename "$PDF" .pdf)
    { echo "# ${BASENAME}"; echo ""; pdftotext -layout "$PDF" -; } > "$OUTPUT_DIR/${BASENAME}.md"
    echo "✅ $BASENAME"
done
echo "完成"
```

**方式 B：Tesseract**
```bash
INPUT_DIR="<用户提供的PDF路径>"
OUTPUT_DIR="${INPUT_DIR}/Markdown_转换结果"
TMPDIR="$HOME/Desktop/.ocr_tmp"
mkdir -p "$OUTPUT_DIR" "$TMPDIR"
for PDF in "$INPUT_DIR"/*.pdf; do
    BASENAME=$(basename "$PDF" .pdf)
    pdftoppm -r 300 -png "$PDF" "$TMPDIR/page" 2>/dev/null
    { echo "# ${BASENAME}"; echo "";
      for IMG in "$TMPDIR"/page-*.png; do
          tesseract "$IMG" stdout -l chi_sim+eng --psm 3 2>/dev/null; echo ""; done
    } > "$OUTPUT_DIR/${BASENAME}.md"
    rm -f "$TMPDIR"/page-*.png
    echo "✅ $BASENAME"
done
rm -rf "$TMPDIR"
echo "完成"
```

**方式 C：PaddleOCR**

> 选择此方式前，先询问用户：
> "请提供您的 PaddleOCR API Token。如尚未申请，可前往 https://aistudio.baidu.com/paddleocr 获取。"

收到 Token 后，定位脚本并执行：

```bash
CASESORT_SCRIPT=$(python3 -c "
import os, pathlib; h = pathlib.Path.home()
for r,d,f in os.walk(h):
    if len(pathlib.Path(r).relative_to(h).parts)>=8: d.clear()
    if 'paddle_ocr.py' in f: print(os.path.join(r,'paddle_ocr.py')); break
")
python3 "$CASESORT_SCRIPT" \
  "<用户提供的PDF路径>" \
  "<用户提供的PDF路径>/Markdown_转换结果" \
  --token "<用户提供的TOKEN>"
```

**步骤 6：报告结果**

转换完成后统计并报告：
- 成功转换：X 个
- 失败/跳过：X 个（列出文件名）
- 输出位置：`<PDF路径>/Markdown_转换结果/`

询问用户是否继续进入第三阶段。

---

## 第三阶段：全量核查

### 目标

对第二阶段转出的全部 Markdown 文件逐一进行 AI 驱动的核查，判断每个案例是否属于目标类型，输出 3 色全量核查 Excel，并按结论将 MD 文件分类归档。

### 执行步骤

**步骤 1：确认 Markdown 文件夹路径**

询问用户 `Markdown_转换结果/` 文件夹的路径（默认为第二阶段的输出路径）。统计文件总数并告知用户。

检查是否存在上次未完成的临时文件：
```bash
python3 -c "
import os, tempfile
f = os.path.join(tempfile.gettempdir(), 'casesort_phase3_tmp.json')
print('⚠️ 发现上次的临时文件，是否清除后重新开始？' if os.path.exists(f) else '✅ 无残留临时文件')
"
```
若用户确认重新开始，执行后继续：
```bash
python3 -c "import os, tempfile; f=os.path.join(tempfile.gettempdir(),'casesort_phase3_tmp.json'); os.path.exists(f) and os.remove(f)"
```

**步骤 2：预分析（读取前 30 个文件）**

自动读取前 30 个 MD 文件，在对话框生成分析报告：

```
【预分析报告】

文件总数：XXX 个
已读取样本：30 个

样本观察：
- 主要文案类型：[行政处罚决定书 / 刑事判决书 / 其他]
- 涉及法律：[列举主要引用法律]
- 可能存在的混淆情形：
  1. [例：第七条虚假宣传案混入商业贿赂]
  2. [例：受贿方被处罚案件（应排除）]
  3. [其他识别到的混淆类型]

建议的纳入标准：
- 确认纳入：[列出判断条件]
- 疑似（需人工核查）：[列出边界条件]
- 确认排除：[列出排除条件]
```

> 询问用户："请确认是否按以上方案进行核查？如有补充或修改，请说明。"

收到用户确认（及补充）后，将判断标准固定，进入批量核查。

**步骤 3：批量核查（20 条/组，3 轮验证）**

将所有 MD 文件按文件名排序，每 20 个为一组逐批处理：

每批执行 3 轮：
- **第 1 轮**：逐一阅读 20 个 MD 文件，根据确认的判断标准作出初步结论（确认纳入 / 疑似 / 确认排除）及简要理由
- **第 2 轮**：回顾第 1 轮全部结论，重点审查边界案例和「疑似」案例，修正明显偏差
- **第 3 轮**：对「疑似」案例再次深度核验，锁定最终结论

3 轮结束后，整理本批次结果（JSON 格式），**必须先将 JSON 写入临时文件，再通过 `$(cat file)` 传入脚本**（直接嵌入 shell 字符串会因法律文书中的引号导致 JSON 解析错误）：

```bash
# 第一步：将 JSON 写入临时文件
BATCH_FILE=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase3_batchN.json'))")
# （将 JSON 内容写入 $BATCH_FILE）

# 第二步：定位脚本并调用
CASESORT_SCRIPT=$(python3 -c "
import os, pathlib; h = pathlib.Path.home()
for r,d,f in os.walk(h):
    if len(pathlib.Path(r).relative_to(h).parts)>=8: d.clear()
    if 'phase3_excel.py' in f: print(os.path.join(r,'phase3_excel.py')); break
")
TMP3=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase3_tmp.json'))")
python3 "$CASESORT_SCRIPT" \
  --append "$(cat "$BATCH_FILE")" \
  --tmp "$TMP3"
```

> ⚠️ **禁止**在 `--append` 后直接写 JSON 字符串，法律文书中的中文引号（如"以...名义"）会破坏 JSON 格式，导致 `JSONDecodeError`。

JSON 数据结构：
```json
[
  {
    "案号": "...",
    "案件事实": "违法目的：...\n违法事实：...",
    "判断结论": "确认纳入",
    "判断依据": "..."
  }
]
```

`案件事实` 写作规范：
- 仅包含**违法目的**和**违法事实**两项，不写处罚时间、行为主体、处罚结果
- 违法目的：违法行为动机，一句话概括
- 违法事实：具体违法行为描述，可适当展开，但全字段（含"违法目的："和"违法事实："标签）**不超过 150 字**
- 确认排除的案件：案件事实仍按此格式简要填写（便于人工复核时理解排除原因）

每批完成后在对话框输出简短进度提示：
> "第 X 组（第 X-X 个）核查完成，确认纳入：X 个，疑似：X 个，排除：X 个。"

**步骤 4：合并输出核查 Excel**

所有批次完成后，调用合并命令：

```bash
CASESORT_SCRIPT=$(python3 -c "
import os, pathlib; h = pathlib.Path.home()
for r,d,f in os.walk(h):
    if len(pathlib.Path(r).relative_to(h).parts)>=8: d.clear()
    if 'phase3_excel.py' in f: print(os.path.join(r,'phase3_excel.py')); break
")
TMP3=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase3_tmp.json'))")
python3 "$CASESORT_SCRIPT" \
  --merge "<Markdown文件夹路径>/../全量核查报告.xlsx" \
  --tmp "$TMP3"
```

输出 Excel 格式：
- 列：案号 | 案件事实 | 判断结论 | 判断依据
- 颜色：🔵 蓝底（#BDD7EE）= 确认纳入，🟠 橙底（#F4B942）= 疑似，🔴 红底（#FF7B7B）= 确认排除
- 排序：蓝色在前，橙色居中，红色在后

**步骤 5：文件夹整理**

在 Markdown 文件夹同级目录下创建结构：
```bash
mkdir -p "<MD文件夹>/../全量核查筛选案例/确认纳入案例"
mkdir -p "<MD文件夹>/../全量核查筛选案例/疑似案例"
```

将对应 MD 文件复制到相应文件夹：
- 确认纳入 → `全量核查筛选案例/确认纳入案例/`
- 疑似 → `全量核查筛选案例/疑似案例/`
- 确认排除 → 不复制

**步骤 6：提示用户**

> "全量核查完成。核查报告已生成：`全量核查报告.xlsx`
>
> 文件已按结论分类复制至：
> - `全量核查筛选案例/确认纳入案例/`（X 个）
> - `全量核查筛选案例/疑似案例/`（X 个）
>
> 请进行人工复核：
> 1. 查看核查报告中的橙色（疑似）案例，确认是否纳入
> 2. 将所有最终确认纳入的案例放入「确认纳入案例」文件夹
> 3. 完成后，告知我进入第四阶段"

---

## 第四阶段：成果产出

### 目标

从「确认纳入案例」文件夹读取 MD 文件，提取结构化字段，生成最终案例库 Excel，并对 MD 文件按序号重命名。

### 执行步骤

**步骤 1：确认输入路径**

询问用户「确认纳入案例」文件夹的路径（默认为第三阶段产出的路径）。

检查是否存在上次未完成的临时文件：
```bash
python3 -c "
import os, tempfile
f = os.path.join(tempfile.gettempdir(), 'casesort_phase4_tmp.json')
print('⚠️ 发现上次的临时文件，是否清除后重新开始？' if os.path.exists(f) else '✅ 无残留临时文件')
"
```
若用户确认重新开始，执行后继续：
```bash
python3 -c "import os, tempfile; f=os.path.join(tempfile.gettempdir(),'casesort_phase4_tmp.json'); os.path.exists(f) and os.remove(f)"
```

**步骤 2：选择输出格式**

> "请选择输出格式：
>
> **格式 A：商业贿赂行政处罚案例**（17 字段，适用于《反不正当竞争法》商业贿赂相关案例）
>
> **格式 B：其他类型案例**（通用字段，适用于其他行政处罚、刑事案件等）"

**格式 A（商业贿赂）17 字段：**

| # | 字段 | 提取规则 |
|---|------|----------|
| 1 | 序号 | 按顺序递增 |
| 2 | 处罚时间 | 格式 YYYY.MM.DD，从文书中提取 |
| 3 | 案例标题 | MD 文件名（去掉 .md 后缀） |
| 4 | 行贿方 | 当事人全称（企业或个人） |
| 5 | 被行贿方 | 收受利益的单位或个人，无法提取填「未披露」 |
| 6 | 被行贿方角色 | 穿透式定性（见下方规则），四选一 |
| 7 | 行贿手段 | 用简短归纳词填写，如：现金、转账、回扣、免费提供设备、实物赠送、代付费用、人头费等；禁止使用完整描述句 |
| 8 | 所处行业 | 细分行业（参考文末行业分类表） |
| 9 | 主要事实 | AI 重写，200字以内（见写作规范） |
| 10 | 行贿金额 | 纯数字（元），无法确定填「/」 |
| 11 | 没收违法所得金额 | 纯数字（元），无则填「/」 |
| 12 | 罚款金额 | 纯数字（元） |
| 13 | 处罚机构 | 行政处罚机关全称 |
| 14 | 地域 | 省级行政区全称（参考文末地域映射表） |
| 15 | 案号 | 完整发文案号 |
| 16 | 法律依据 | 法条名称及具体条款，违法条款与处罚条款分行列示 |
| 17 | 备注 | 失效、补偿、积极配合等重要信息 |

**被行贿方角色穿透式定性规则（按顺序判断）：**
1. 提取利益最终流向的具体接收方
2. 若接收方是作为单位整体的交易合作方 → `交易相对方`
3. 若是合作方内部的普通员工、部门负责人，或医疗领域利用处方权等具体职务便利的普通医护人员 → `交易相对方的工作人员`
4. 若接收方凭借特定行政管理职级（院长、党委书记）或信息/社会关系优势（导游、司机、介绍人）干预引导了交易定向 → `利用职权或者影响力影响交易的单位或者个人`
5. 案卷事实缺失必要判定要素 → `未披露`

**法律依据提取规则（格式A）：**
- 文本含「2025修订」或「第二十四条」或「第8条（处罚条款）」→ 2025修订版
- 否则 → 2019修正版
- 2019修正版：商业贿赂禁止=第七条，处罚=第十九条
- 2025修订版：商业贿赂禁止=第八条，处罚=第二十四条
- 过渡期案例（违法行为在2025.10.15前，处罚在后）：违法条款用2019第七条，处罚条款用2025第二十四条
- 若涉及《药品管理法》：追加第八十八条和/或第一百四十一条

**格式 B（通用类型）基础字段：**

| # | 字段 | 说明 |
|---|------|------|
| 1 | 序号 | 按顺序递增 |
| 2 | 处罚/裁判时间 | 格式 YYYY.MM.DD |
| 3 | 案例标题 | MD 文件名 |
| 4 | 行为主体 | 实施该行为的人/单位及其身份 |
| 5 | 行为对方 | 相对方及其身份（如有） |
| 6 | 主要事实 | AI 重写，200字以内（见写作规范） |
| 7 | 地域 | 省级行政区全称 |
| 8 | 案号 | 完整发文案号 |
| 9 | 法律依据 | 法条名称及具体条款 |
| 10 | 备注 | 其他重要信息 |

选择格式 B 时，根据案件类型推荐额外字段，告知用户后确认：
> "本表格包含以上基础字段，根据本批案件类型，推荐加入 [XXX] 字段，请问是否同意？或有其他需要补充的字段？"

**步骤 3：批量提取（10 条/组，2 轮复核）**

将「确认纳入案例」中的 MD 文件按文件名排序，每 10 个为一组：

每批执行 2 轮：
- **第 1 轮**：逐一阅读 10 个 MD 文件，提取所有字段，「主要事实」使用 AI 重写（见写作规范）
- **第 2 轮**：复核提取结果，重点检查：金额是否为纯数字、案号是否完整、法条版本是否正确、主要事实是否在 200 字以内且格式达标

2 轮结束后，**必须先将 JSON 写入临时文件，再通过 `$(cat file)` 传入脚本**（同第三阶段，直接嵌入会因引号导致解析错误）：

```bash
# 第一步：将 JSON 写入临时文件
BATCH_FILE=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase4_batchN.json'))")
# （将 JSON 内容写入 $BATCH_FILE）

# 第二步：定位脚本并调用
CASESORT_SCRIPT=$(python3 -c "
import os, pathlib; h = pathlib.Path.home()
for r,d,f in os.walk(h):
    if len(pathlib.Path(r).relative_to(h).parts)>=8: d.clear()
    if 'phase4_excel.py' in f: print(os.path.join(r,'phase4_excel.py')); break
")
TMP4=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase4_tmp.json'))")
python3 "$CASESORT_SCRIPT" \
  --format A \
  --append "$(cat "$BATCH_FILE")" \
  --tmp "$TMP4"
```

> ⚠️ **禁止**在 `--append` 后直接写 JSON 字符串，原因同第三阶段。

每批完成后输出进度提示：
> "第 X 组（第 X-X 个案例）提取完成。"

**主要事实写作规范（格式A商业贿赂）：**
> 「当事人[XXX]系[行业]经营者，为谋取[交易机会/竞争优势]，于[时间段]以[手段]向[对象]给予[利益]共计[金额]，违反[法条]。[处罚机构]依法对其作出[处罚内容]的行政处罚。」
- 使用专业法律语言，200字以内
- 涵盖：主体 + 背景 + 行为动机 + 违法行为 + 关键手段 + 最终目的 + 处罚结果

**主要事实写作规范（格式B通用）：**
> 「[主体]系[身份]，于[时间]以[手段/方式]实施[违法行为]，[目的/结果]。[处罚/裁判机构]依据[法条]对其作出[处理结果]。」
- 使用专业法律语言，200字以内

**步骤 4：合并输出**

所有批次完成后，调用合并命令：

```bash
CASESORT_SCRIPT=$(python3 -c "
import os, pathlib; h = pathlib.Path.home()
for r,d,f in os.walk(h):
    if len(pathlib.Path(r).relative_to(h).parts)>=8: d.clear()
    if 'phase4_excel.py' in f: print(os.path.join(r,'phase4_excel.py')); break
")
TMP4=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase4_tmp.json'))")
python3 "$CASESORT_SCRIPT" \
  --merge "<确认纳入案例路径>/../案例库_最终版.xlsx" \
  --tmp "$TMP4"
```

**步骤 5：MD 文件重命名**

按最终 Excel 中的序号，对「确认纳入案例」文件夹中的 MD 文件加序号前缀：
- 格式：`001_原文件名.md`、`002_原文件名.md` …
- 序号补零至 3 位（不足 100 条）或 4 位（100 条以上）

重命名完成后报告：已重命名 X 个文件。

**步骤 6：询问是否进入第五阶段**

> "成果产出完成！案例库已保存至：[路径]
>
> 是否需要进入第五阶段，基于此案例库撰写法律案例分析报告？"

---

## 第五阶段：报告产出

### 目标

基于第四阶段产出的案例库 Excel，撰写结构化、专业且逻辑严密的法律案例分析报告，输出 Word 文档。

### 执行步骤

**步骤 1：确认数据来源**

询问用户案例库 Excel 的路径（默认为第四阶段产出的路径）。

**步骤 2：读取 Excel，规划报告结构**

读取全部案例数据后，在对话框输出拟定的报告分类方案，征求用户确认：

> "已读取 [X] 个案例。拟将案例分为以下类别：
>
> 1. [类别名称]（X 个案例）：[简要说明分类依据]
> 2. [类别名称]（X 个案例）：[简要说明]
> 3. [类别名称]（X 个案例）：[简要说明]
>
> 分类维度：[说明按何种维度划分，如行贿手段、被行贿方角色、行业等]
>
> 是否同意此分类方案？或需要调整？"

**步骤 3：撰写报告并生成 Word**

用户确认后，按以下结构撰写报告内容，构建 JSON 数据结构，写入临时文件，再调用 `phase5_report.py` 生成 Word 文档。

**报告 JSON 结构（按此格式构建数据）：**

```json
{
  "title": "XXX案例分析报告",
  "subtitle": "数据来源：XXX | 时间范围：XXXX年",
  "sections": [
    {
      "heading": "一、数据概览",
      "level": 1,
      "content": [
        {"type": "paragraph", "text": "本报告共收录案例 X 个，覆盖时间范围..."},
        {"type": "bullet", "items": ["地域分布：...", "主要法律：...", "处罚机构：..."]}
      ]
    },
    {
      "heading": "二、分类分析",
      "level": 1,
      "content": [
        {"type": "paragraph", "text": "按[维度]将案例分为以下类别："}
      ]
    },
    {
      "heading": "（一）[类别名称]（X 个案例）",
      "level": 2,
      "content": [
        {"type": "paragraph", "text": "该类案例的共性特征..."},
        {"type": "bullet", "items": ["法律风险1", "法律风险2"]},
        {"type": "table",
         "headers": ["序号", "案例标题", "案号", "处罚日期"],
         "rows": [["1", "案例标题", "案号", "日期"]]}
      ]
    },
    {
      "heading": "三、监管执法关键关注点",
      "level": 1,
      "content": [
        {"type": "bullet", "items": ["执法重点领域...", "高频违规模式...", "新出现违规形态..."]}
      ]
    },
    {
      "heading": "四、企业合规建议",
      "level": 1,
      "content": [
        {"type": "bullet", "items": ["制度层面：...", "操作层面：...", "高风险场景：..."]}
      ]
    }
  ]
}
```

构建完成后，将 JSON 写入临时文件，再调用脚本生成 Word：

```bash
# 第一步：将报告 JSON 写入临时文件
TMP5=$(python3 -c "import tempfile,os; print(os.path.join(tempfile.gettempdir(),'casesort_phase5_report.json'))")
# （将 JSON 内容写入 $TMP5）

# 第二步：定位脚本并生成 Word
CASESORT_SCRIPT=$(python3 -c "
import os, pathlib; h = pathlib.Path.home()
for r,d,f in os.walk(h):
    if len(pathlib.Path(r).relative_to(h).parts)>=8: d.clear()
    if 'phase5_report.py' in f: print(os.path.join(r,'phase5_report.py')); break
")
python3 "$CASESORT_SCRIPT" \
  --input "$TMP5" \
  --output "<案例库Excel同级目录>/案例分析报告.docx"
```

完成后告知用户文档路径。

---

## 地域推断参考（CITY_TO_PROVINCE）

提取「地域」字段时参考（从文书中检测到城市名后映射到省级行政区）：

北京市→北京市、上海市→上海市、天津市→天津市、重庆市→重庆市
广州/深圳/珠海/佛山/东莞/惠州/中山/汕头/湛江/茂名→广东省
南京/苏州/无锡/常州/南通/扬州/镇江/徐州/常熟/张家港/泰州→江苏省
杭州/宁波/温州/绍兴/嘉兴/台州/金华/义乌/湖州→浙江省
成都/绵阳/广元/南充/宜宾/泸州→四川省
武汉/宜昌/荆州/襄阳/十堰/恩施/保康→湖北省
长沙/株洲/湘潭/常德/永州/祁阳/沅陵→湖南省
济南/青岛/烟台/潍坊/临沂/济宁→山东省
郑州/洛阳/开封/新乡→河南省
西安/咸阳/宝鸡→陕西省
合肥/芜湖/马鞍山/六安/黄山/定远/望江/桐城/滁州→安徽省
福州/厦门/泉州/漳州→福建省
南昌/九江/赣州/上饶/吉安/泰和/余江/安义/兴国/靖安→江西省
沈阳/大连/鞍山→辽宁省、哈尔滨/齐齐哈尔/大庆→黑龙江省
长春/吉林→吉林省、太原/大同→山西省、石家庄/唐山/保定→河北省
呼和浩特→内蒙古自治区
南宁/桂林/柳州/河池/贺州/百色/玉林→广西壮族自治区
海口/三亚→海南省、昆明/大理/丽江→云南省、贵阳/遵义→贵州省
兰州/天水→甘肃省、西宁→青海省、银川→宁夏回族自治区
乌鲁木齐/喀什→新疆维吾尔自治区、拉萨→西藏自治区

---

## 行业分类参考（格式A商业贿赂专用，21 类）

| 行业 | 关键词 |
|------|--------|
| 医疗器械 | 医疗器械、器械、医疗科技、耗材、试剂 |
| 医药/制药 | 医药、药品、制药、药业 |
| 医疗服务 | 医院、诊所、卫生院、非急救转运、救护车、康复、护理院 |
| 建筑/工程 | 建筑、施工、工程、装修、装饰、水电安装、市政 |
| 房地产 | 房地产、置业、开发、物业 |
| 食品/饮料 | 食品、饮料、酒、食品科技 |
| 餐饮/外卖 | 餐饮、酒店、外卖、餐厅、加盟 |
| 汽车/交通 | 汽车、车辆、年检、4S店 |
| 金融/保险 | 银行、保险、证券、基金、金融 |
| 信息技术 | 软件、信息技术、科技、互联网、系统集成 |
| 工业设备 | 机械、设备、数控、工业自动化 |
| 化工/检验 | 化工、试剂、检验、化学品 |
| 旅游/文化 | 旅游、旅行社、景区、文化传播 |
| 教育/培训 | 教育、学校、培训、幼儿园 |
| 传媒/广告 | 广告、传媒、媒体、影视 |
| 零售/商贸 | 超市、零售、贸易、批发 |
| 农业 | 农业、农产品、种植、养殖 |
| 能源/矿产 | 能源、矿业、煤炭、石油、电力 |
| 物流/运输 | 物流、运输、快递、货运 |
| 环保/检测 | 环保、检测、认证、监测 |
| 其他 | 无法归类时使用 |
