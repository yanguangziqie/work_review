# 工作总结生成系统（日报 / 周报 / 月报 / 季度总结）

根据日常工作内容文件（支持 **Word / Excel / TXT / Markdown / PDF / CSV**），结合可选模板，自动生成**日报 / 周报 / 月报 / 季度总结**，并支持导出 **Word / Markdown / TXT / Excel / HTML**。

## 功能

1. **多格式解析**：递归扫描输入目录，抽取 docx（含表格）、xlsx（全工作表、日期/数字格式清洗、单元格换行归一）、txt/md/csv、pdf 的文本；**工单/事项表自动识别表头语义**（名称/描述/状态/解决方法/时间），按结构化字段挖掘
2. **按报告周期过滤 + 自动推断**：日报=某天 / 周报=某周（2026W40）/ 月报=某月 / 季度=2026Q3，条目级过滤；时间范围支持 `2026-10-01`、`2026/9/21`、`2026年9月21日`、`2026W40`、`2026-10`、`2026Q3` 等写法（自动归一），**填了就严格按你填的**；留空时若目标周期内没有内容，自动从文件名/内容推断日期（取出现最多的日期）后生成，或忽略过滤全量纳入，每一步都写进执行日志
3. **工作事项挖掘**：自动识别「已完成 / 进行中」事项，识别系统/项目名称并归一化分组，提取量化数据
4. **模板开关（显式）**：
   - 📄 **使用模板**：上传模板（Word / Excel(.xls/.xlsx) / TXT / Markdown / CSV / PDF），按模板结构生成，输出格式跟随模板；
   - 📝 **不使用模板**：直接生成——填写了「生成要求」就按要求产出文档（Markdown，需 AI 模式理解文字要求；规则模式输出自由格式总结），不填按内置标准结构（一、总体情况 / 二、已完成事项 / 三、进行中事项）
5. **两种生成方式**：
   - **规则生成**：离线、快，抽取式组织（适合内容本身很规整的记录）
   - **AI 生成**：调用大模型归纳润色（支持 OpenAI 兼容接口 / 本地 Ollama）；「生成要求」作为提示词注入，如“侧重写数据中台成效、语言正式、800字以内”
6. **模板填充**：两类结构自动识别——**章节式**（一、总体情况…）按章节填充；**表格表单式**（绩效 PBC、KPI 承诺书等）识别列角色自动生成带表格的成品（权重平均分配、评价标准沿用模板默认值）
7. **多格式导出**：生成后可选导出 **Word (.docx) / Markdown (.md) / 纯文本 (.txt) / Excel (.xlsx) / 网页 (.html)**，成品之间互转（docx 表格内容会抽取进各格式）；不支持的格式/缺组件给出明确提示
8. **友好报错与日志**：任何错误返回中文提示 + 详细日志（每个文件被解析/过滤/跳过的原因、自动推断过程），不再出现看不懂的报错

## 快速开始

### 方式一：网页版（推荐）

```bash
cd quarterly_review
python3 webapp.py
# 浏览器打开 http://127.0.0.1:8899
# 局域网访问：python3 webapp.py --host 0.0.0.0，然后访问 http://<本机IP>:8899
```

页面流程：拖入工作记录文件 → 选**报告类型**（日报/周报/月报/季度总结）与**时间范围**（留空=当前周期；支持 `2026/9/21` 等写法）→ 选**是否使用模板**（不使用时可填**生成要求**直接按要求生成）→ 选**生成方式**（规则 / AI）→ 点「生成」→ 下载成品，或在下载区**选导出格式**（Word/Markdown/TXT/Excel/HTML）。

> 成品文件名 = 文档标题（填了就用）或报告类型名（如 `日报.md`、`周报.docx`），**不带日期**——日期以标题/正文内容为准，避免猜错日期误导。

### AI 生成配置

页面「生成方式」里选 🤖 AI 生成，两种接入方式：

| 方式 | 配置 |
|---|---|
| OpenAI 兼容接口（API Key） | 接口地址如 `https://api.openai.com/v1`、`https://api.deepseek.com/v1`；填 API Key 与模型名（如 `gpt-4o-mini`、`deepseek-chat`） |
| 本地 Ollama（无需 Key） | 接口地址 `http://localhost:11434`（默认）；模型下拉选择或手动输入，如 `qwen2.5:14b` |

页面提供「**测试连接并获取模型**」按钮：验证接口连通性并自动拉取模型列表（Ollama 读 `/api/tags`，OpenAI 兼容读 `/models`），模型可下拉选择。

**AI 输出可靠性**：Ollama 请求强制 `format:"json"`；解析端内置 JSON 修复器（未转义引号/尾逗号/注释自动修复）；解析失败自动重试一次。小型号偶发输出不合规 JSON 时可重试或换更大模型。无模板 + 生成要求的自由生成走对话接口，直接产出整篇 Markdown。

API Key 仅用于本次请求，不会落盘保存；也可用环境变量 `QSUM_API_BASE` / `QSUM_API_KEY` / `QSUM_MODEL` 配置。

命令行等价用法：

```bash
python3 generate_summary.py -i input -q 2026Q3 --ai \
    --provider openai --api-base https://api.deepseek.com/v1 \
    --api-key *** --model deepseek-chat

python3 generate_summary.py -i input -q 2026Q3 --ai \
    --provider ollama --model qwen2.5:14b
```

### 方式二：命令行

```bash
cd quarterly_review

# 把你的工作内容文件放进 input/ 目录（或任意目录），然后：
python3 generate_summary.py -i input -q 2026Q3                    # 季度总结
python3 generate_summary.py -i input --type week                 # 本周周报
python3 generate_summary.py -i input --type day -p 2026-10-01    # 某天日报
python3 generate_summary.py -i input --type month -p 2026-10     # 某月月报

# 输出：
#   output/2026年第三季度工作总结.docx    ← 成品（格式随模板）
#   output/..._工作内容摘要.md            ← 中间摘要，供核对/润色
```

依赖：Python 3.9+，`pip install python-docx openpyxl pdfplumber flask`（PDF 备选 `pypdf`）。
网页版导出 Excel/Word 需要 `openpyxl` / `python-docx`（上述依赖已含）；HTML/Markdown/TXT 导出无额外依赖；PDF 导出已移除（如需 PDF 可导出 Word 后另存）。

## 常用参数（命令行）

| 参数 | 说明 |
|---|---|
| `-i/--input` | 工作内容文件或目录，可多个（目录递归扫描） |
| `--type` | 报告类型：`day`/`week`/`month`/`quarter`（默认按时间格式自动识别） |
| `-p/--period` | 时间范围：`2026-10-01`（日）/`2026W40`（周）/`2026-10`（月）/`2026Q3`（季），默认当前 |
| `-q/--quarter` | 兼容旧用法，等价 `--type quarter --period <值>` |
| `-t/--template` | 模板路径（docx/xlsx/xls/txt/md/csv/pdf）；不传则按内置标准结构输出 docx |
| `--extra-prompt` | AI 生成的额外要求（提示词），不填走默认 |
| `-o/--output` | 输出路径，默认 `output/<周期标签><报告类型>.docx` |
| `--no-filter` | 不做时间过滤，纳入全部输入文件与事项 |
| `--ai` | 使用 AI 大模型生成（默认规则模式） |
| `--provider` | `openai`（OpenAI 兼容接口）或 `ollama`（本地） |
| `--api-base` / `--api-key` / `--model` | 接口配置，也可用环境变量 `QSUM_API_BASE` / `QSUM_API_KEY` / `QSUM_MODEL` |
| `--max-completed` / `--max-wip` | 条目数量上限（默认不限制） |
| `--title` | 覆盖文档标题 |
| `--no-digest` | 不输出 Markdown 摘要 |

示例：

```bash
# 指定多个来源文件、自定义输出与标题
python3 generate_summary.py -i 周报.xlsx 日志.md -i /path/to/docs \
    -q 2026Q3 -o output/三季度总结.docx --title "XX团队2026年三季度工作总结"

# 不传模板：内置标准结构输出 docx
python3 generate_summary.py -i input -q 2026Q3

# AI 生成 + 自定义要求
python3 generate_summary.py -i input -q 2026Q3 --ai --provider ollama --model qwen2.5:14b \
    --extra-prompt "侧重写数据中台成效，语言正式，800字以内"

# 纳入全部文件（不按时间过滤）
python3 generate_summary.py -i input --no-filter
```

## 工作原理

```
输入文件 ──▶ 多格式文本抽取 (qsum/extractors.py)
        ──▶ 周期过滤 + 事项挖掘/分类/项目分组 (qsum/miner.py)
        ──▶ 总体情况/已完成/进行中 内容生成 (qsum/generator.py)
        ──▶ 按模板样式填充输出 (qsum/template_filler.py)
        ──▶ 导出转换 Word/Markdown/TXT/Excel/HTML (webapp.py)
```

- **项目名识别**：以「系统/平台/应用/APP/中台…」为锚点提取，全局投票归一化（如「移动办公APP审批模块」「移动办公APP消息推送模块」自动归入「移动办公APP」）
- **状态判定**：按句中首个状态标记（“上线/验收/完成” vs “开发中/测试中/进行中/建设中…”），例如“开发中，预计9月底上线”正确判为进行中
- **表格行清洗**：Excel 行自动去掉状态列、重复列；工单表按表头语义取「名称/解决方法/状态」，描述列自动去噪（URL/电话/工号/对话短句）
- **量化数据**：识别“3000人/120份/40%”等指标，优先放入“关键数据”
- **周期自动推断**：时间范围留空且目标周期内无内容时，从文件名（如 `事情列表260920_260925`）与内容日期（`2026-09-21` / `2026年9月21日` 等）推断，取**出现最多的日期**锚定周期；无法识别则忽略过滤全量纳入，过程写入日志

## 生成内容的定位

自动生成的是**结构完整的初稿**：

- 时间、系统名、动作、量化数据均来自你的真实工作记录，不会编造
- 「（此处配3－5幅系统截图）」为配图占位，需自行插入截图
- 部分措辞可在成品中直接润色；也可以把生成的 `_工作内容摘要.md` 发给我，让我帮你润色改写

## 目录结构

```
quarterly_review/
├── webapp.py                  # Web 服务（python3 webapp.py）
├── webui/index.html           # Web 前端页面
├── generate_summary.py        # 命令行入口
├── qsum/
│   ├── extractors.py          # docx/xlsx/txt/md/csv/pdf 文本抽取
│   ├── miner.py               # 周期过滤、事项挖掘、项目分组
│   ├── generator.py           # 规则模式内容生成
│   ├── llm.py                 # AI 生成（OpenAI 兼容 / Ollama）
│   └── template_filler.py     # 模板填充
├── templates/工作总结模板.*   # 模板示例（docx/xlsx/txt/md + 日报/周报模板，按需替换）
├── input/                     # 示例输入（换成你的工作记录）
└── output/                    # 生成结果
```

## 常见问题

- **旧版 .doc/.xls**：不支持直接解析，请另存为 .docx/.xlsx
- **换模板**：页面选「使用模板」上传即可（输出格式跟随模板：md→md、Excel→xlsx、txt→txt，其它→docx）；两类结构自动识别：① 章节式（含「一、总体情况 / 二、已完成事项 / 三、进行中事项」类章节标题）；② 表格表单式（如绩效 PBC，含「指标/权重/评价标准」列头即可识别）
- **绩效模板使用**：用 PBC 等表格模板生成时，指标列填入工作事项，权重自动平均分配、评价标准沿用模板默认值——权重与标准需人工校准（系统不瞎编考核指标）
- **不想要固定结构**：选「不使用模板」+ 填写「生成要求」（AI 模式），直接按要求产出整篇文档；规则模式则输出自由格式总结
- **日期不对 / 想指定日期**：在时间范围里直接填（`2026/9/21` 即可），填了就严格按你填的；留空才会自动推断。成品文件名不带日期，日期以标题/正文为准
- **导出格式**：Word / Markdown / TXT / Excel / HTML 互转；提示“不支持/缺组件”时按提示安装对应包（`python-docx` / `openpyxl`）；PDF 导出已移除，可导出 Word 后另存为 PDF
- **Excel 质量**：表格行天然是碎片记录，规则模式产出偏“流水账”属正常；要质量用 AI 生成模式，或补充带叙述文字的文档类工作记录
- **AI 生成失败**：检查 API Key / 接口地址 / 模型名；Ollama 需先 `ollama pull` 对应模型；本地大模型生成耗时可能达数分钟，属正常；失败时可退回规则模式
- **文件被跳过**：看页面日志面板，会写明每个文件被解析/过滤/跳过的原因；目标周期没内容时系统会自动降级（按文件日期推断周期或忽略过滤），不会直接报错

## 变更记录（近版）

- **v5**：成品文件名不带日期（=标题或报告类型名）；时间范围支持 `2026/9/21` 等写法、填了严格生效；自动推断改取“出现最多的日期”；导出去掉 PDF、新增 HTML
- **v4**：新增「是否使用工作总结模板」显式开关；选“是”未上传文件给出提示
- **v3**：无模板 + 生成要求 → 直接按要求生成（不套默认模板）；下载区支持多格式导出，不支持的格式给明确提示
- **v2**：时间过滤自动降级（按文件日期推断周期 / 忽略过滤全量纳入）；修复 AI 模式周报/日报生成崩溃（周期标签解析问题）；错误返回 JSON + 堆栈、出错展开日志面板
