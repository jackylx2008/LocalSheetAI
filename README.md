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

## 目录结构

```text
.
├── excel_ai.py                 # Excel AI 工作流入口
├── ai_self_check.py            # 本地 AI 自检入口
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
- 本地 AI 服务配置和模型路径只放在 `common.env`。
- 写回 Excel 时只修改输出副本中的新增分类列。

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
```
