# LocalSheetAI

LocalSheetAI 是一个面向本地 Excel 问题清单的 AI 辅助整理工具。项目会读取 `input/` 目录下匹配的工作簿，按配置处理指定 sheet，从第 3 行表头、第 4 行数据开始抽取问题描述和整行上下文，并调用本地 llama.cpp 兼容接口生成分类口径草案或逐行分类结果。

## 当前能力

- 按通配符匹配 `input/` 下的 Excel 文件，并按文件名日期和修改时间选择最新文件。
- 支持一个工作簿配置多个 sheet。
- 结构化读取 Excel 明细行，识别 E 列 `问题描述及原因初步判定`。
- 生成分类口径草案 Markdown：`output/category_skills_draft.md`。
- 按已生成的分类口径逐行调用本地 AI 分类。
- 在输出副本中追加最后一列 `AI整理归类项`，保存到 `output/`，不覆盖原始 Excel。
- 写回时优先使用本机 Excel COM 操作副本，尽量保留原始格式、图片、保护状态和现有内容。
- 可扫描指定目录下的审批单 PDF，并把对应设计变更编号汇总到既有审批单统计 Excel 中，默认保留原工作簿格式。

## 目录结构

```text
.
├── excel_ai.py                 # Excel AI 工作流入口
├── ai_self_check.py            # 本地 AI 自检入口
├── approval_pdf_summary.py     # 审批单 PDF 扫描与统计 Excel 更新入口
├── config.yaml                 # 主配置，读取 common.env 中的变量
├── common.env.example          # 环境变量示例，不包含本机私有路径
├── src/localai/
│   ├── flows/                  # 工作流编排
│   └── modules/                # Excel、分类、配置、本地 AI 客户端模块
├── input/                      # 本地输入 Excel，已被 git 忽略
├── output/                     # 输出文件，已被 git 忽略
├── log/                        # 日志，已被 git 忽略
└── vendor/                     # 本地运行依赖 DLL，已被 git 忽略
```

## 环境准备

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .\common.env.example .\common.env
```

编辑 `common.env`，配置本地 llama.cpp 服务、模型路径和 Excel 匹配规则。`common.env` 包含本机路径和运行参数，默认不提交。

常用配置示例：

```dotenv
EXCEL_WORKBOOK_SHEETS_JSON={"*质保问题清单-汇总*.xlsx":["02 机电-水","03 机电-暖","05 机电-排风"]}
EXCEL_HEADER_ROW=3
EXCEL_DATA_START_ROW=4
EXCEL_PROBLEM_COLUMN=E
EXCEL_CATEGORY_HEADER=AI整理归类项
CATEGORY_DRAFT_OUTPUT_PATH=./output/category_skills_draft.md
CATEGORY_SEED_CATEGORIES_JSON=["污水提升泵问题","隔油池装置问题","空调机组问题","消防风机问题","排风机问题","风阀问题","空调水阀问题","能量表问题"]
```

Windows 终端建议设置 UTF-8 输出，避免中文显示异常：

```powershell
$env:PYTHONIOENCODING='utf-8'
```

### 审批单 PDF 统计配置

审批单 PDF 扫描使用独立的本地配置文件 `.approval_pdf_summary.env`，用于保存目标 PDF 目录和统计 Excel 路径。该文件包含本机路径，默认不提交。

配置示例：

```dotenv
APPROVAL_PDF_DIR=D:\CloudStation\国会二期\12 北京院-B24地块\酒店设计变更\审批单扫描文件\识别整理后审批单
APPROVAL_EXCEL_PATH=D:\CloudStation\国会二期\12 北京院-B24地块\酒店设计变更\审批单扫描文件\识别整理后审批单\审批单统计.xlsx
APPROVAL_PDF_PREFIX=审批单_
APPROVAL_STATUS_TEXT=有审批单
APPROVAL_SCAN_RECURSIVE=false
APPROVAL_APPEND_MISSING=true
APPROVAL_CLEAR_MISSING=false
APPROVAL_HEADER_ROW=1
APPROVAL_ID_COLUMN=1
APPROVAL_STATUS_COLUMN=2
```

## 使用方式

准备阶段，只解析目标 Excel 并抽取结构化行，不调用本地 AI：

```powershell
$env:PYTHONIOENCODING='utf-8'
.\.venv\Scripts\python.exe -u .\excel_ai.py --mode prepare --no-llama --show-preview
```

生成分类口径草案：

```powershell
$env:PYTHONIOENCODING='utf-8'
.\.venv\Scripts\python.exe -u .\excel_ai.py --mode draft
```

逐行分类并输出 Excel 副本：

```powershell
$env:PYTHONIOENCODING='utf-8'
.\.venv\Scripts\python.exe -u .\excel_ai.py --mode classify
```

审批单 PDF 扫描预览，不写入 Excel：

```powershell
$env:PYTHONIOENCODING='utf-8'
.\.venv\Scripts\python.exe -u .\approval_pdf_summary.py --dry-run
```

审批单 PDF 正式汇总。脚本只在检测到实际变更时保存工作簿；保存前会自动生成时间戳备份：

```powershell
$env:PYTHONIOENCODING='utf-8'
.\.venv\Scripts\python.exe -u .\approval_pdf_summary.py
```

常用选项：

- `--dry-run`：只输出扫描和计划写入结果，不保存 Excel。
- `--no-append-missing`：PDF 中识别到、但 Excel 里不存在的编号不追加新行。
- `--clear-missing`：当 Excel 中已有审批单状态、但目录中找不到对应 PDF 时，清空该状态。默认不清空。
- `--no-backup`：保存前不生成备份文件。默认会在目标 Excel 同目录生成 `*.backup_YYYYMMDD_HHMMSS.xlsx`。
- `--pdf-dir`、`--excel`：临时覆盖 `.approval_pdf_summary.env` 中的目录和目标 Excel 路径。

输出文件命名示例：

```text
output/2026-04-28_质保问题清单-汇总_192335_AI整理归类.xlsx
```

## 工作流说明

1. `prepare` 解析配置中的工作簿和 sheet，抽取第 4 行起的有效行。
2. `draft` 按 sheet 汇总样本行，调用本地 AI 生成分类口径草案。
3. 用户确认或调整 `output/category_skills_draft.md`。
4. `classify` 读取分类口径，对每行单独调用本地 AI，只输出一个分类标签。
5. 程序复制原始 Excel 到 `output/`，在目标 sheet 最后一列追加 `AI整理归类项`。

## 安全约定

- 不覆盖 `input/` 中的原始 Excel。
- 不提交 `input/`、`output/`、`log/`、`.venv/`、`vendor/` 和 `common.env`。
- 不提交 `.approval_pdf_summary.env`，审批单 PDF 目录和目标 Excel 路径只保存在本地。
- 本地 AI 服务配置和模型路径只放在 `common.env`。
- 写回 Excel 时只修改输出副本中的新增分类列。
- 审批单 PDF 统计写回目标 Excel 前会自动备份；没有实际变更时不会保存改写工作簿。

## 主要命令

```powershell
# 本地 AI 自检
.\.venv\Scripts\python.exe -u .\ai_self_check.py

# Excel 准备
.\.venv\Scripts\python.exe -u .\excel_ai.py --mode prepare --no-llama

# 生成分类口径草案
.\.venv\Scripts\python.exe -u .\excel_ai.py --mode draft

# 生成带 AI 分类的新 Excel
.\.venv\Scripts\python.exe -u .\excel_ai.py --mode classify

# 预览审批单 PDF 统计
.\.venv\Scripts\python.exe -u .\approval_pdf_summary.py --dry-run

# 更新审批单统计 Excel
.\.venv\Scripts\python.exe -u .\approval_pdf_summary.py
```
