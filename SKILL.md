# 港股 IPO 基石投资者数据库 — 技能大纲

> 本项目由 Claude Code 主导开发，此文件作为全局规范指导所有代码设计与实现。
> 基于真实 PPTX（`2026基石_0622.pptx`）逆向分析得出数据结构。

---

## Skill 0: 项目总览

### 目标
构建一个**港股 IPO 基石投资者数据库**系统，将每周更新的 PPT 源数据结构化，提供多维度查询界面（按投资者、按项目、按行业等），支持定期增量更新。

### 核心用户
- 本部门同事（非技术背景）
- 使用场景：浏览器直接打开 HTML，输入投资者名称即可查询其参与的所有项目

### 核心数据流

```
ppts/ 文件夹（每周放入新 PPT）
     ↓  [Skill 2+3: PPT 提取 → data.js]
data.js（window.cornerstoneData — 基石投资数据）
     ├── [fetch_company_info.py: AKShare → company_info.js]
     │    股价、公司简介、首日涨跌幅等
     └── [compile_ipo_details.py 自动发现新公司 → ipo_details.js]
          │ 已有数据保留，新公司追加空条目
          └─ [Claude 通过 MCP 自动补充] 保荐人、认购倍数等
              ↓
hk_cornerstone_investors.html 加载全部数据文件
Skill 4: 响应式前端交互（浏览器直接打开，零安装）
```

### 数据源说明
- 基石投资数据：**Dealogic、HKEX**（交易所披露信息），通过 PPT 人工整理
- 公司基础信息：**AKShare** 自动抓取（股价、公司简介、上市日期等）
- IPO 补充详情：**东方财富 MCP**（保荐人、认购倍数、募资净额等）
- 汇率：HKD:CNY = 1:0.93；HKD:USD = 1:0.13
- 更新频率：约每周一次
- 文件命名：`<年份>基石_<MMDD>.pptx`，如 `2026基石_0622.pptx`

### 技术选型

| 维度 | 选择 | 理由 |
|------|------|------|
| 核心数据格式 | **`data.js`（`window.cornerstoneData` 数组）** | 浏览器 `file://` 协议直接可用，绕过 CORS |
| 公司详情数据 | **`company_info.js`（`window.companyInfo`）** | AKShare 预抓取，含股价/财务/简介 |
| IPO 补充数据 | **`ipo_details.js`（`window.ipoDetails`）** | 东方财富 MCP 查询，含保荐人/认购倍数等 |
| 前端 | **纯 HTML + Tailwind CSS CDN** | 零依赖、零安装，Tailwind 比 Bootstrap 更轻 |
| 数据提取 | **Python (python-pptx)** | 解析 PPT 表格和文本 |
| 版本管理 | **Git** | 追踪每周数据更新历史 |
| 投资者别名 | **JSON 映射文件** | 处理简繁体/空格/符号导致的同名不同写 |

---

## Skill 1: 文本解析与清洗（Text Parsing）

**目标**：将单元格内换行排列的复杂文本（如 `GIC（80）\nHHLRA 高瓴（50）`），精准拆分为独立的投资者和金额。

**核心逻辑**：

1. 以 `
` 和 `` 为候选分隔符切分行（`` 统一转为 `
`）
2. 对每一行，用正则提取投资者名称和最后一对括号内的金额
3. 不匹配正则的行**不丢弃**，缓存起来向前合并到下一匹配行的名称前
   - 解决 PPT 手工换行导致投资者名被截断的问题
4. 处理边缘情况：双层括号、中英双语名称、千分位逗号、全角括号
5. 金额字符串转浮点数
---

## Skill 2: PPT 矩阵表格提取（PPT Table Extraction）

**目标**：利用 `python-pptx` 库，读取标准化 PPT 中的表格数据，还原为内存中的二维结构。

### PPT 表格结构（实际逆向分析）

| 实际列 | 表头（第 1 行） | 子表头（第 2 行） | 说明 |
|--------|------------|-------------|------|
| 0 | **上市日期** | — | 格式不统一：`2026-06-17` / `2026/5/6` |
| 1 | **代码** | — | 后缀不统一：`6675` / `9981.HK` |
| 2 | **公司** | — | 中文名 / 英文名，含后缀 `-B`/`-W`/`-P` |
| 3 | **行业** | — | 主/子行业，换行分隔如 `元线传感\nSoc` |
| 4 | **预计上市总规模**（HKD 百万） | — | 含千分位逗号：`2,567` |
| 5 | **基石规模**（HKD 百万） | — | 浮点数格式 |
| 6 | **基石占比** | — | 如 `29%` |
| 7 | **基石投资者**（HKD 百万） | **上下游企业** | 产业上下游公司 |
| 8 | | **政府基金** | 政府引导基金、国资委、主权基金 |
| 9 | | **上市前股东** | Pre-IPO 股东加码 |
| 10 | | **战略合作** | 战略客户/渠道/合作伙伴 |
| 11 | | **其他投资者** | PE/VC、对冲基金、资管等金融机构 |
| 12 | **基石投资中非企业投资者占比** | — | 基石盘子里非产业资本的比例 |

### 核心逻辑

1. 遍历 `ppts/` 下所有 `.pptx`
2. 定位 slide 中的 `Table` 对象
3. 跳过前 2 行表头，从第 3 行开始循环
4. 列索引 **0-6** → 项目基本信息（日期、代码、公司、行业、发行规模、基石规模、占比）
5. 列索引 **7-11** → 5 类基石投资者文本（每格内可能是多行多投资者）
6. 列索引 **12** → 非企业投资者占比
7. 对列 7-11 的每格文本，调用 **Skill 1** 拆解为独立投资者
8. 输出二维数组：`[{project_meta}, {project_meta}, ...]`

### 已知注意事项

- **日期格式不统一**：需同时兼容 `yyyy-MM-dd` 和 `yyyy/M/d`
- **投资者名称含双重括号**：`隆威香港（保隆汽车）（3.06）` — 正则需从右向左匹配金额
  - **行业换行不视为子类分隔**：PPT 中的 `
` 是人工排版换行，不应拆分为子类。直接合并为连续字符串，不加空格或分隔符。例如 `元线传感
Soc` → `元线传感Soc`
- **投资者之间的真实分隔符是 `\x0b`（垂直制表符）**：这是 PPT 表格单元格内用于换行的真实条目分隔符。用 `\x0b` 切分候选段，段内的 `\n` 是手工换行（人工错误），合并为空格后统一解析
- **不匹配正则的行向前合并**：如果某行不包含 `名称（金额）` 格式，则视为上一行的续行，合并后再尝试解析。避免 `Libra Fixed Income One SP\n（庄家颖）` 被错误拆成两条
- **股票代码后缀**：`.HK` 有/无，统一规范（建议保留后缀）
- **公司名后缀含义**：`-B`=生物科技、`-W`=同股不同权、`-P`=未盈利

---

## Skill 3: 数据平铺与规范化（Data Flattening）

**目标**：将"一项目对多投资者"的二维表格，平铺为"一笔投资一条记录"的一维数组。

### 核心概念对比

| 传统嵌套结构（×） | 扁平结构（✓） |
|-------------------|--------------|
| 一个项目一条记录 | 一笔投资一条记录 |
| 投资者是项目的子数组 | 投资者和项目同级 |
| 查询需遍历嵌套层级 | 直接 `.filter()` 搞定 |

### 核心逻辑

1. 遍历提取到的每一行项目数据
2. 遍历该项目的 5 类基石投资者列表（列 7-11）
3. 每位投资者生成一条**包含完整项目信息**的独立记录
4. 添加 `investor_category` 字段标注来源类别（上下游企业/政府基金/上市前股东/战略合作/财务投资者）
5. 所有记录合并为单一数组

### 输出格式（`data.js`）

```js
window.cornerstoneData = [
  {
    // === 项目信息（每条记录完整携带） ===
    company: "琻捷电子",
    stock_code: "6675",
    listing_date: "2026-06-17",
    sector: "元线传感Soc",
    total_ipo_size_hkd_million: 981,
    cornerstone_size_hkd_million: 127,
    cornerstone_pct: 29,              // 基石占发行比例 %
    non_corporate_investor_pct: 22,   // 基石中非企业投资者占比 %

    // === 该笔投资信息 ===
    investor_name: "Oakwise",
    investor_category: "财务投资者",  // 上下游企业 | 政府基金 | 上市前股东 | 战略合作 | 财务投资者
    amount_hkd_million: 9.57,

    // === 来源追溯 ===
    source_ppt: "2026基石_0622.pptx",
    slide_no: 1
  }
  // ... 共 N 条记录（N = 所有投资者去重求和）
]
```

### 类型转换规则

| 原始 PPT 值 | 转换后 | 规则 |
|-------------|--------|------|
| `2026/5/6` | `2026-05-06` | 统一 `YYYY-MM-DD` |
| `2,567` | `2567` | 去千分位逗号 |
| `29%` | `29` | 百分比去 `%` 转 number |
| `5.20` | `5.2` | 金额浮点数 |
| 空单元格 | `null` | 缺失字段 |
| `\x0b`（单元格内换行） | `\n` | 统一换行符 |

### 补充文件

**投资者别名规范化**（`data/investor_aliases.json`）：

投资者名称来自 PPT 人工录入，存在同名不同写问题。规范化规则集中管理在此文件，每次提取时自动应用：

| 规则类型 | 说明 | 示例 |
|---------|------|------|
| `exact` | 精确匹配全名后替换 | `富国（富国基金&富国香港）` → `富国基金` |
| `contains` | 名称包含关键字即替换（不区分大小写） | `广发基金` → `广发基金`（含"广发基金"的自动统一） |
| `substring` | 名称中任意位置替换 | `（通过` → `（`，删除"通过" |
| `case_insensitive` | 不区分大小写全名匹配 | `Orbimed` → `奥博资本 OrbiMed` |

当前已设置的清洗规则详见 `data/investor_aliases.json`（**88 组、288 条精确匹配规则**），已于 2026-06-28 经用户逐项确认并应用到全部数据。覆盖以下类型的问题：

| 问题类型 | 示例 | 涉及组数 |
|---------|------|:--------:|
| 中英文混写统一 | `Hillhouse` / `HHLRA` / `高瓴资本` → `Hillhouse 高瓴` | ~40组 |
| TRS/掉期后缀归一 | `源峰` / `源峰（银河TRS）` → `CPE 源峰` | ~15组 |
| 投资通道说明归一 | `博裕` / `Aqua Ocean（博裕）` → `Boyu 博裕` | ~10组 |
| 拼写/错别字修正 | `Gallatlion` → `Gallantlion`；`雾淞` → `雾凇` | 2组 |
| 源数据错误修正 | `禾荣科技合富(中国)` → `禾荣科技` | 1处 |
| 子公司归入母公司 | `Aranda Investments` → `Temasek 淡马锡` | 1组 |
| 简繁体/空格统一 | `嘉实 基金`（含空格）→ `Harvest 嘉实` | 多处 |

> 规则命名格式：有英文名的统一为 `"英文名 中文名"`（空格分隔），仅中文的保留中文，仅英文的保留英文。
>
> **注意**：部分名称相似但实际为不同投资主体（如工银理财 ≠ 工银瑞信、景林系列各自独立），未做合并，详见规则文件。
>
> 此文件会持续积累，发现新的同名不同写时，添加规则后重新运行提取即可。

---

## Skill 4: 响应式前端交互（Reactive UI Filtering）

**目标**：构建单文件 HTML 看板，同事双击即用，支持秒级搜索与联动筛选。

### 核心技术

| 维度 | 方案 |
|------|------|
| 样式框架 | **Tailwind CSS**（CDN 引入，`https://cdn.tailwindcss.com`） |
| 数据加载 | `<script src="data.js">` — 直接挂到 `window.cornerstoneData` |
| 搜索 | 原生 JS `.filter()` + `.includes()`，中文模糊匹配 |
| 图标 | 可选 Heroicons SVG 内联 |

### 核心逻辑

1. **数据加载**：`<script src="data.js">` 加载 `window.cornerstoneData` 数组
2. **搜索过滤**：
   - 输入投资者名称 → `.filter(item => item.investor_name.includes(keyword))`
   - 输入股票/公司名 → `.filter(item => item.company.includes(keyword) || item.stock_code.includes(keyword))`
   - 同时支持行业下拉联动筛选
3. **聚合计算**（纯前端，零后端）：
   ```js
   // 按投资者聚合
   groupBy(data, 'investor_name') → {name, total_amount, project_count, projects: [...]}
   // Top N 排行
   sortBy(total_amount).slice(0, 5)
   ```
4. **自动渲染管理看板**（页面顶部，3 卡片商务风格）：
   - 📊 项目总数（含数据截止日期）
   - 📊 基石总额（单位：百万美元）
   - 📊 Top 3 最活跃基石投资者（按参与项目数排名，带序号）

### 功能列表

- **按投资者查询**：输入名称 → 显示该投资者参与的所有项目及金额
- **按项目查询**：输入公司名 / 股票代码 → 显示该项目的基石投资者完整列表
- **按行业筛选**：下拉选择行业 → 过滤结果
- **按上市年份筛选**：下拉选择 2025年/2026年/全部
- **按投资者类别筛选**：上下游企业 / 政府 / 老股东 / 战略 / 财务
- **结果排序**：按日期、金额、占比、解禁收益率等点击排序（表头点击切换升降序）
- **手动分页**：输入页码或点击上一页/下一页，每页20条
- **导出 CSV**：当前搜索结果一键导出（含上市日期、解禁收益率等字段）
- **投资者详情侧栏**：点击投资者名称 → 弹出侧栏展示其所有投资明细、总金额、参与项目数
- **公司详情侧栏**：点击公司名 → 弹出侧栏展示上市信息、发行数据、基石投资者列表、解禁收益率、公司简介

### 技术要求

- 纯前端，浏览器双击 `hk_cornerstone_investors.html` 即用（`file://` 协议）
- 搜索支持中文模糊匹配（原生 `.includes()` 即可）
- 无数据时显示友好提示（"未找到匹配结果"）
- 响应式布局（Tailwind 的 `sm:/md:/lg:`），电脑/平板/手机都能用
- 全量数据加载（~1,282 投资记录，151 项目），无分页压力

---

## Update 流程：每周数据更新（增量模式）

### 工作原理
extract_pptx.py 支持 **增量更新**——每次只处理新增或修改过的 PPT 文件，与历史数据自动合并去重。

```
首次运行: 所有 PPT → data.js + data.json + .pptx_index.json
                                      ↓
后续运行: 对比 .pptx_index.json，只提取新/改过的 PPT
          → 与 data.json 中的历史记录合并去重
          → 输出更新后的 data.js + data.json
          → 更新 .pptx_index.json
```

> ⚠️ **编码说明**：`更新数据.bat` 已全部使用纯英文输出，避免 cmd.exe 在中文 Windows 下编码解析错误。提示信息含义见下方。

### 方式 A（推荐 — 给同事用）
同事只需将**新 PPT** 放入 `ppts/` 文件夹，然后双击 `更新数据.bat` 即可：

```
ppts/ 放入新 PPT（保持旧 PPT 不变）
     ↓  [双击 更新数据.bat]
① 自动增量提取 PPT → 别名清洗 → 合并去重 → data.js
② 自动检查新名称：如有未覆盖的名称，保存 data/新增名称检查.md
③ 自动抓取 AKShare 公司详情 → company_info.js
④ 自动检测新公司 → IPO 详情追加空条目（不阻塞）
⑤ 自动生成单文件版（可分享领导）
⑥ 自动生成名称清洗确认清单 → 名称清洗清单.txt（同事可直接打开）
     ↓
刷新 hk_cornerstone_investors.html 即可查看最新数据
```

> 💡 `extract_pptx.py` 每次提取时自动从 `data/investor_aliases.json` 加载别名规则并清洗投资者名称。
> 💡 如果新 PPT 中出现未覆盖的名称，流程**不阻塞**，同事照常用；报告留给你后续处理。
> 💡 新公司的 IPO 详情（保荐人/认购倍数等）自动追加空条目。找 Claude 说"更新数据"，Claude 会运行批处理、自动通过 MCP 查新公司数据、补充完整。

### 股价每日自动更新

系统已设置 Windows 定时任务，**每天 17:30（港股收盘后）自动运行**：

```
Windows 任务计划程序
     ↓  每天 17:30
每日更新股价.bat
     ↓
① fetch_company_info.py — 刷新所有港股最新收盘价（~50秒）
② bundle_standalone.py — 同步更新 standalone 单文件版
     ↓
同事/领导查看时，股价始终是上一个交易日的最新数据
```

> 定时任务名称: `HK_IPO股价更新`
> 如需修改时间，运行: `schtasks /change /tn HK_IPO股价更新 /st 新时间`

> 💡 **同时放多个新 PPT 也没问题**：脚本会逐个处理所有新增文件，统一合并去重。
> 💡 **想加入 2025 年数据？** 把 `2025基石_*.pptx` 放进 `ppts/` 即可，跨年数据自动合并。

### 方式 B（命令行 — 给开发者用）
```
① 新 PPT 放入 ppts/
② python scripts/extract_pptx.py          # 增量模式（自动检测）
   python scripts/extract_pptx.py --force  # 强制全量重新提取
③ 查看新增名称建议（如有）：
   cat data/新增名称检查.md
④ 刷新 hk_cornerstone_investors.html
```

### ⚠️ 重要注意事项

| 场景 | 说明 |
|------|------|
| **保留旧 PPT** | 旧 PPT 文件**不要删除**。脚本需要它们来检查是否已处理过 |
| **改名旧 PPT** | 如果重命名了旧 PPT（如 `2026基石_0622.pptx` → `2026基石_0622_old.pptx`），下次运行会视为**新文件**重新导入，导致数据重复 |
| **修正数据** | 如果某个 PPT 数据有误、重新导出后覆盖原文件，脚本会检测到 **mtime 变化** 自动重新提取该文件 |
| **强制重新提取** | 如果数据乱了，删掉 `data/.pptx_index.json` 或加 `--force` 参数即可全量重来 |
| **2025 + 2026 合并** | 直接把 2025 年 PPT 放入 `ppts/`，它们不在索引中，会自动被增量处理 |

### 初始化同事环境（一次性）
1. 把整个 `hk_ipo_csdata/` 文件夹发给同事
2. 同事需要安装 Python（[https://www.python.org/downloads/](https://www.python.org/downloads/)），安装时勾选 **"Add Python to PATH"**
3. 打开 cmd / PowerShell 运行一次 `pip install python-pptx`
4. 之后每次更新只需：**放新 PPT → 双击 `更新数据.bat`**
5. 如果同事只是查看数据不改，连 Python 都不用装，双击 `hk_cornerstone_investors.html` 即可

> ⚠️ `更新数据.bat` 在部分中文 Windows 上可能出现编码解析错误（如 `'on' 不是内部或外部命令`）。
> 如果双击闪退，按 `Win+R` → 输入 `cmd` → 在黑窗口中运行：
> ```cmd
> cd /d C:\Users\xxx\Desktop\AI AUTO\hk_ipo_csdata
> 更新数据.bat
> ```
> 把报错文字发回来排查。

---

## Skill 5: 公司上市信息详情侧栏（Company Detail Sidebar）

**目标**：点击表格中的公司名称，右侧弹出侧栏展示该公司的上市信息、发行数据、基石投资者列表、上市表现和公司简介。

### 数据来源

| 数据项 | 来源 | 说明 |
|--------|------|------|
| `ipo_size_mn`（发行总规模） | **PPT 已有数据** | 取自 `cornerstoneData` 的 `ipo_size_mn` |
| `cs_size_mn`（基石规模） | **PPT 已有数据** | 取自 `cornerstoneData` 的 `cs_size_mn` |
| `cs_ratio`（基石占比） | **PPT 已有数据** | 取自 `cornerstoneData` 的 `cs_ratio` |
| `company_profile`（公司简介） | AKShare | 预抓取到 `company_info.json` |
| `board`（上市板块） | AKShare | 预抓取到 `company_info.json` |
| `listing_date`（上市日期） | AKShare | 预抓取到 `company_info.json` |
| `issue_price`（发行价） | AKShare | 预抓取到 `company_info.json` |
| `first_day_change_pct`（首日涨跌幅） | AKShare | 历史行情计算 |
| `change_6m_pct`（解禁收益率） | AKShare 计算 | 上市日期+180天后首日收盘 vs 发行价 |
| `sponsors`（保荐人） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |
| `stabilizing_agent`（稳价人） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |
| `public_sub_multiple`（公开认购倍数） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |
| `intl_sub_multiple`（国际认购倍数） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |
| `net_proceeds_mn`（募资净额） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |
| `lot_size`（每手股数） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |
| `overallotment_shares`（超额配售） | **东方财富 MCP** | 预抓取到 `ipo_details.json` |

> **解禁收益率口径**：上市日期 + 180 天 → 第一个交易日收盘价 → 对比发行价计算涨跌幅。对应基石投资者 6 个月锁定期满后的解禁收益。未到解禁期的股票显示为"-"。

### 数据流

```
① 基石投资数据（PPT 源）
   scripts/extract_pptx.py → data/data.js

② 公司基础信息（AKShare）
   scripts/fetch_company_info.py → data/company_info.json + company_info.js

③ IPO 补充详情（东方财富 MCP — 自动发现 + Claude 补充）
   data.json 有新代码 → compile_ipo_details.py 自动追加空条目
   Claude 通过 MCP 查询 → 填充到 ipo_details.json

前端加载：
   hk_cornerstone_investors.html
     ├── <script src="data/data.js">           → window.cornerstoneData
     ├── <script src="data/company_info.js">    → window.companyInfo
     └── <script src="data/ipo_details.js">     → window.ipoDetails
                          ↓
          用户点击公司名称 → 侧栏展示全部数据
```

### company_info.json 数据格式

```json
{
  "6675.HK": {
    "stock_code": "6675.HK",
    "company_name": "琻捷电子",
    "company_name_en": "SENASIC Electronics Technology Co., Ltd.",
    "industry": "半导体",
    "board": "H股",
    "listing_date": "2026-06-17",
    "issue_price": 18.36,
    "currency": "HKD",
    "total_ipo_size_mn": 981,
    "cs_size_mn": 127,
    "cs_ratio": "29%",
    "first_day_change_pct": 27.12,
    "change_6m_pct": null,
    "company_profile": "琻捷电子科技(上海)股份有限公司是全球领先的...",
    "data_status": {
      "profile": "ok",
      "issue_price": "ok",
      "listing_date": "ok",
      "first_day_return": "ok",
      "sponsor": "missing",
      "subscription_ratio": "missing"
    }
  }
}
```

### 港股代码转换规则

```python
# 数据中的格式: "6675.HK"
# AKShare 要求的格式: "06675"（5位数字）
def to_hk_code(code):
    return code.replace(".HK", "").zfill(5)
```

### 更新流程

```bash
# 在每周提取 PPT 数据之后运行
python scripts/fetch_company_info.py

# 输出 → data/company_info.json
# 刷新 HTML 即可看到公司详情
```

### 侧栏布局

```
┌────────────────────────────────┐
│ × 关闭                         │
│                                │
│ 🏢 琻捷电子                    │
│    6675.HK · 半导体            │
│                                │
│ ═══ 上市信息 ═══              │
│ 上市板块: H股                  │
│ 上市日期: 2026-06-17          │
│ 发行价: 18.36 HKD             │
│                                │
│ ═══ 发行规模 ═══              │
│ 发行总规模: 981 百万美元       │
│ 基石规模: 127 百万美元         │
│ 基石占比: 29%                 │
│                                │
│ ═══ 基石投资者 (N家) ═══      │
│   欣旺达香港         $5.20M    │
│   Oakwise            $9.57M    │
│   ...                         │
│                                │
│ ═══ 上市表现 ═══              │
│ 上市首日涨跌幅: +27.12%        │
│ 解禁收益率: -35.53%            │
│（锁定期6个月，未到解禁期显示"-")│
│                                │
│ ═══ 公司简介 ═══              │
│ 琻捷电子科技(上海)股份有限     │
│ 公司是全球领先的...            │
└────────────────────────────────┘
```

## Skill 6: 投资者别名管理与新增规则（Alias Management）

**目标**：管理投资者名称别名规则的生命周期——当新 PPT 中出现未覆盖的名称时，
检测、确认、入库、重应用。

### 别名规则文件

位置：`data/investor_aliases.json`

| 规则类型 | 说明 | 优先级 |
|---------|------|:------:|
| `exact` | 精确匹配全名后替换 | 1（最高） |
| `contains` | 名称包含关键字即替换（不区分大小写） | 2 |
| `substring` | 字符串替换（如删除名称中间的"通过"） | 3 |
| `case_insensitive` | 不区分大小写全名匹配后替换 | 4（最低） |

### 新增规则的触发时机

```
新 PPT → extract_pptx.py 增量提取
              ↓
        自动应用现有别名规则清洗
              ↓
        check_new_names.py 检测新名称
              ↓
    ┌─────────┴──────────┐
    ↓                     ↓
  全部已覆盖             有未覆盖名称
    ↓                     ↓
  静默通过         保存 data/新增名称检查.md
                   打印提醒（不阻塞流程）
```

### 新增规则的流程

**场景：同事跑完更新，你拿到 `data/新增名称检查.md` 报告**

同事双击 `更新数据.bat` 后会自动在末尾生成 **`名称清洗清单.txt`**（记事本可直接打开），
同事打开后每条填写"合并"或"保留"，把文件发给你即可。你也可以直接取 `data/新增名称检查.md` 处理。

```
① 获取报告：
   ┌─ 同事给 → 名称清洗清单.txt（已分类，带填空位）
   └─ 自己看 → data/新增名称检查.md（原始完整版）

② 判断每个名称：
   ├─ 确认为同一机构 → 在 investor_aliases.json 的 "exact" 中添加规则
   │  "exact": {
   │    "旧名称": "标准名称"
   │  }
   │
   ├─ 拿不准是不是同一机构 → 找 Claude 分析语义判断
   │  提供报告内容，Claude 可以理解中英文含义，
   │  比纯字符模糊匹配准确得多
   │
   └─ 确认是新机构 → 无需处理，下次自动跳过

③ 规则确认后，重新清洗已有数据：
   python scripts/apply_aliases.py

④ 提交版本：
   git add data/investor_aliases.json data/data.js data/data.json
   git commit -m "别名规则更新：新增 X 组"

⑤ （可选）重新运行名称清洗确认清单（覆盖旧清单）：
   python scripts/generate_checklist.py
```

### 工具脚本

| 脚本 | 功能 | 何时用 |
|------|------|--------|
| `check_new_names.py` | 检测新名称覆盖情况，出报告 | 集成在 `extract_pptx.py` 中自动调用 |
| `apply_aliases.py` | 已有数据重新执行别名清洗 | 新增/修改别名规则后手动运行 |

### 已知未覆盖名称的处理建议

对于 fuzzy matching 产生的候选建议，注意：

- 部分建议是准确的（如中文名相同的通常正确）
- 部分建议是纯字符巧合（如"Lake Bleu"和"Stoneylake Global"都含"lake"但不同实体）
- **拿不准的带着报告来找 Claude 分析**，语义理解比纯字符匹配可靠得多

### 别名规则更新记录

#### 2026-07-14 第二批批量更新（～40 组合并）

同事放入 `2026基石_0713.pptx` 后运行 `更新数据.bat`，自动生成 `data/新增名称检查.md`，
报告 260 个未覆盖名称。经人工确认后批量处理：

| 类别 | 内容 |
|------|------|
| **高瓴系列** | `HHLRA（高瓴）`、`HHLR（高瓴）`、`CPPIB（高瓴SMA）` → Hillhouse 高瓴 |
| **CPE 源峰系列** | `CPE Neem（CPE源峰）`、`CPE River（CPE源峰）`、`CPE源峰（CPE INVESTMENT XV）`、`源峰(银河TRS)` → CPE 源峰 |
| **标准名补充** | 博裕、橡树资本、贝莱德、鼎晖、高毅、太保、富达、Wind Sabre、大湾区基金、Huadeng 等变体合并 |
| **新建标准名** | `清池资本 Lake Bleu`（含旧数据 `清池资本` 统一）、`Barings 霸菱`、`Millennium 千禧` |
| **景林 TRS 统一** | `景林(中金TRS)`、`景林（国泰君安TRS）`、`上海景林(TRS-华泰)` → 景林资产 |
| **品牌/通道统一** | OPPO（含天进贸易）、TCL（含Metazone）、腾讯黄河、高盛资管、摩根资管 JPM AM |
| **工银瑞信** | 新增 `contains` 规则，任何含"工银瑞信"的名称自动归一 |
| **JPMorgan 改名** | `JPMorgan 摩根大通` → 全部改为 `摩根资管 JPM AM`（含历史数据，约 7 条记录） |
| **常春藤** | `IvyRock` + `常春藤` + `IvyRock 常春藤` → IvyRock 常春藤 |
| **Athos** | `Athos（ATHOS CAPITAL）` → `Athos` |
| **Hel Ved** | `Hel Ved（HEL VED MASTER FUND）` → `Hel Ved` |
| **保留原名** | `瀚亚（Eastspring）`（与 Eastspring 保诚不同机构）、`汇添富基金`（非富国） |

#### 2026-06-28 首批 88 组合并

详见下方 "别名清洗记录" 章节。

---

## 附录 A：当前数据统计摘要（2026-07-14）

| 项目 | 值 |
|------|-----|
| 数据来源文件 | `2025基石_1231.pptx` + `2026基石_0622.pptx` + `2026基石_0706.pptx` + `2026基石_0713.pptx` |
| 总项目数 | **161** 个 IPO 项目 |
| 总投资记录数 | **1,411** 条（一笔投资一条记录） |
| **清洗后** 唯一投资者数 | **743** 位（累计合并 100+ 组别名） |
| 别名规则总数 | **333 条**（330 exact + 3 contains + 2 substring + 1 case_insensitive） |
| 覆盖行业 | **100+** 个细分行业 |
| 基石总额 | **$28,477.70M** |
| 2025 年上市 | 约 90 家 |
| 2026 年上市 | 约 71 家 |
| 有解禁收益率数据 | 约 90 家（其余未到6个月锁定期） |
| 首家上市 | 2025年首批 |
| 最晚上市 | 2026-07-10（晶合集成） |
| 最大 IPO | 胜宏科技（2,567 HKD 百万） |
| 最高基石占比 | 曦智科技-P（72%） |

> 📊 别名清洗历史：2026-06-28 首批 88 组合并 → 2026-07-14 第二批 ~40 组合并（详见 Skill 6）

## 附录 B：基石投资者 5 分类速查

| 类别 | 典型机构 | 特点 |
|------|---------|------|
| **上下游企业** | 欣旺达、TCL、OPPO、小米、Glencore | 产业上下游，战略协同目的 |
| **政府基金** | 卡塔尔投资局 ADIA、GIC、江西国控、惠州国资委 | 政府/主权背景，金额较大 |
| **上市前股东** | MSIP、Focustar Capital | 已持股，IPO 加码 |
| **战略合作** | 腾讯、阿里巴巴、京东、联想 | 大客户/渠道/技术合作方 |
| **财务投资者** | HHLRA/高瓴、CPE/源峰、UBS GAM、BlackRock、泰康 | PE/VC/对冲基金/险资，追求回报 |

## 附录 C：项目目录结构

```
hk_ipo_csdata/
├── SKILL.md                  # 本文件
├── CLAUDE.md                 # Claude Code 项目指令
│
├── data/
│   ├── data.js               # ★ 核心数据（浏览器加载，window.cornerstoneData）
│   ├── data.json             # ★ 纯 JSON 格式（给 Python 增量读回用）
│   ├── company_info.json     # ★ 公司详情数据（AKShare 预抓取，Skill 5）
│   ├── .pptx_index.json      # ★ 已处理 PPT 文件索引（增量追踪用，自动维护）
│   └── investor_aliases.json # 投资者别名映射（人工维护）
│
├── ppts/                     # ★ PPT 源文件放这里（旧文件不要删！）
│   └── README.md
│
├── scripts/
│   ├── extract_pptx.py       # Skill 1+2+3：提取 → 解析 → 平铺 → data.js（支持增量/全量）
│   ├── apply_aliases.py      # 别名规则重应用（新增规则后重新清洗已有数据）
│   ├── check_new_names.py    # 新名称覆盖检测 + 疑似合并提醒（被 extract_pptx.py 调用）
│   ├── generate_checklist.py # 更新后自动生成名称清洗确认清单 .txt（给同事用）
│   ├── fetch_company_info.py # Skill 5：AKShare 批量抓取公司详情 → company_info.json
│   ├── bundle_standalone.py  # 打包单文件 HTML（数据内嵌，即开即用）
│   └── validate_data.py      # 数据校验 + 别名检测
│
├── 同事更新操作说明.md        # ★ 给同事的详细更新指南（无 Claude Code 也可用）
├── 名称清洗清单.txt          # ★ 更新后自动生成（同事可直接双击打开，填写确认）
├── 更新数据.bat              # ★ 双击即可增量更新（给同事用，纯英文防编码乱码）
├── 每日更新股价.bat           # 定时任务：每日17:30自动刷新股价 + 同步standalone
├── hk_cornerstone_investors.html              # Skill 4：前端看板（双击即用）
└── hk_cornerstone_investors_standalone.html   # ★ 单文件版（数据内嵌，可直接发领导）
```

## 附录 D：查找与修改指南

如需修改 SKILL.md 或项目的任何部分，随时打开本目录启动 Claude Code，告诉我：

- **"修改 SKILL.md，在 Skill 2 增加 XXX 列的处理"**
- **"帮我调整 hk_cornerstone_investors.html 的搜索逻辑"**
- **"更新数据结构，增加 XXX 字段到平铺输出"**
- **"分析新的 PPT 结构"**（当 PPT 模板有变化时）
- **"运行更新流程"**（放入新 PPT 后执行完整更新）
