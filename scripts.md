# 代码施工蓝图 (scripts.md)

> 本文件为项目代码的施工蓝图，定义了所有脚本的输出规范和前端功能架构。
> 严格遵循 SKILL.md 和 references.md 中定义的技能体系与数据规范。
> **先写蓝图，再写代码。**

---

## 1. 脚本体系总览

```
ppts/*.pptx
    │
    ▼
┌─────────────────────────────────────────────────┐
│  extract_pptx.py               (Skill 1+2+3)    │
│  ├─ 读取 ppts/ 下所有 .pptx                      │
│  ├─ 解析表格 → 二维数组          (Skill 2)       │
│  ├─ 正则拆解投资者文本           (Skill 1)       │
│  ├─ 平铺为"一笔投资一条记录"      (Skill 3)       │
│  └─ 输出 data/data.js                           │
└─────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────┐
│  validate_data.py                (数据校验)       │
│  ├─ 检查 data.js 完整性                          │
│  ├─ 检测疑似同名不同写投资者                      │
│  └─ 输出校验报告到命令行                          │
└─────────────────────────────────────────────────┘
    │
    ▼
data/data.js
    │  <script src="data/data.js">
    ▼
┌─────────────────────────────────────────────────┐
│  hk_cornerstone_investors.html (Skill 4)        │
│  ├─ 加载 data.js                                │
│  ├─ 管理看板（统计摘要）                          │
│  ├─ 多维筛选 + 搜索                             │
│  ├─ 结果列表 + 排序                             │
│  ├─ 投资者详情侧栏                              │
│  └─ 导出 CSV                                   │
└─────────────────────────────────────────────────┘
```

---

## 2. extract_pptx.py — 提取脚本规范

### 2.1 功能职责

一个脚本完成从 PPTX → `data.js` 的全链路：
1. 扫描 `ppts/` 目录下的所有 `.pptx` 文件
2. 解析每个 PPT 中的表格（Skill 2）
3. 对列 7-11 中的投资者文本执行正则解析（Skill 1）
4. 平铺为"一笔投资一条记录"的扁平数组（Skill 3）
5. 排序后输出为 `data/data.js`

### 2.2 输入 / 输出

| | 路径 |
|---|---|
| 输入目录 | `ppts/*.pptx` |
| 输出文件 | `data/data.js` |

### 2.3 输出格式（`data/data.js`）

```js
// data/data.js — 自动生成，请勿手动修改
// 生成时间: 2026-06-24 14:30:00
// 来源文件: 2026基石_0622.pptx
// 数据截止: 2026-06-21

window.cornerstoneData = [
  {
    // ========== 项目信息 ==========
    listing_date: "2026-06-17",         // String, YYYY-MM-DD, AKShare官方数据
    stock_code: "6675.HK",              // String, 统一带 .HK 后缀
    company_name: "琻捷电子",             // String
    industry: "元线传感Soc",             // String, 换行合并无分隔符
    ipo_size_mn: 981.0,                 // Float, 百万美元
    cs_size_mn: 127.0,                  // Float, 百万美元
    cs_ratio: "29%",                    // String, 保留百分号

    // ========== 该笔投资信息 ==========
    investor_name: "Oakwise",           // String
    investor_category: "财务投资者",      // String: 上下游企业 | 政府基金 | 老股东 | 战略合作 | 财务投资者
    amount_mn: 9.57,                    // Float, 百万美元

    // ========== 来源追溯 ==========
    source_ppt: "2026基石_0622.pptx",   // String
    slide_no: 1                          // Int, 所在幻灯片页码
  },
  // ... 数千条记录
];
```

### 2.4 列映射关系（对应 references.md Table Schema）

| PPT 列 | data.js 字段 | 处理逻辑 |
|--------|-------------|---------|
| 0 | `listing_date` | 取自 AKShare `stock_hk_security_profile_em`，格式 YYYY-MM-DD |
| 1 | `stock_code` | 统一补 `.HK`，`6675` → `6675.HK` |
| 2 | `company_name` | 直接取文本，保留 `-B`/`-W`/`-P` 后缀 |
| 3 | `industry` | 换行合并为连续字符串（无分隔符） |
| 4 | `ipo_size_mn` | 去千分位逗号，转 Float |
| 5 | `cs_size_mn` | 去千分位逗号，转 Float |
| 6 | `cs_ratio` | 保留原字符串，含 `%` |
| 7 | investor_corp | → 平铺，`investor_category="上下游企业"` |
| 8 | investor_gov | → 平铺，`investor_category="政府基金"` |
| 9 | investor_pre_ipo | → 平铺，`investor_category="老股东"` |
| 10 | investor_strategic | → 平铺，`investor_category="战略合作"` |
| 11 | investor_financial | → 平铺，`investor_category="财务投资者"` |
| 12（有则取） | 不输出此列 | 仅用于校验参考，不进入 data.js |

### 2.5 正则解析引擎（Skill 1 实现）

```
对列 7-11 每格文本：
1. 按换行符 \n 切分为行列表
2. 每行 trim 后跳过空行和 "-""
3. 对每行执行正则匹配最后一对括号:

   正则: /^(.+?)[（(]\s*([\d,]+\.?\d*)\s*[)）]$/

   - group(1) → investor_name（投资者名称）
   - group(2) → amount_mn（金额，去逗号后转 Float）

4. 处理边缘情况:
   - 多重括号: "Tembusu (David Su) (2.04)"
     → 从右向左匹配最后一对括号
     → name: "Tembusu (David Su)", amount: 2.04
   - 全角括号: "欣旺达香港（5.20）"
     → 正则中 [（(] 同时匹配全角和半角
```

### 2.6 输出顺序

按 `listing_date` **降序**排列（最新的在前），同日期内按 `company_name` 升序。

### 2.7 幂等性

多次运行同一 PPT 应产生完全一致的输出。脚本不应依赖任何可变状态（如当前时间）。

---

## 3. validate_data.py — 校验脚本规范

### 3.1 功能职责

1. 读取 `data/data.js`，检查每条记录必填字段是否完整
2. 字段类型校验（`ipo_size_mn` 应为 Number，`listing_date` 应匹配 `YYYY-MM-DD` 等）
3. 检测疑似同名不同写的投资者（如 `高瓴` vs `HHLRA 高瓴` vs `Hillhouse`）
4. 输出校验报告到命令行

### 3.2 校验规则

| 规则 | 说明 | 失败时 |
|------|------|--------|
| 必填字段非空 | `investor_name`, `company_name`, `stock_code` 等 | 报错 + 退出码 1 |
| 金额为正数 | `amount_mn` > 0 | 报错 + 退出码 1 |
| 日期格式 | `listing_date` 匹配 `^\d{4}-\d{2}-\d{2}$` | 报错 + 退出码 1 |
| 类别合法 | `investor_category` 在 5 类白名单中 | 报错 + 退出码 1 |
| 无重复主键 | `(stock_code, listing_date)` 组合不重复 | 警告（非阻断） |
| 疑似同名 | 名称相似度检测（Jaccard 或编辑距离） | 打印建议列表 |

### 3.3 输出示例

```bash
$ python scripts/validate_data.py

✅ data.js 校验通过
   - 总记录数: 347
   - 唯一项目数: 44
   - 唯一投资者数: 189
   - 覆盖行业数: 26

⚠️ 检测到以下疑似同名投资者（建议补充 investor_aliases.json）:
   "高瓴" (23次) ↔ "HHLRA 高瓴" (8次) ↔ "Hillhouse" (3次)
   "源峰" (5次) ↔ "CPE" (12次)
   "贝莱德" (6次) ↔ "BlackRock" (15次)
```

---

## 4. hk_cornerstone_investors.html — 前端功能架构

### 4.1 技术栈

| 维度 | 方案 |
|------|------|
| 样式 | **Tailwind CSS**（CDN: `https://cdn.tailwindcss.com`） |
| 数据 | `<script src="data/data.js">` → `window.cornerstoneData` |
| 图标 | Heroicons SVG（内联） |
| 交互 | 原生 JavaScript（零依赖） |

### 4.2 页面布局（从上到下）

```
┌──────────────────────────────────────────────────────┐
│  🏦 港股 IPO 基石投资者数据库                          │
│  数据截止: 2026-06-21 ｜ 共 44 个项目, 189 位投资者    │
├──────────────────────────────────────────────────────┤
│  📊 管理看板（3 卡片商务风格）                        │
│  ┌──────────────┐ ┌──────────────┐ ┌──────────────┐ │
│  │    项目总数   │ │   基石总额    │ │ 最活跃投资者  │ │
│  │     44       │ │   8,420      │ │ 1. 泰康(12)  │ │
│  │  个 IPO      │ │   百万美元    │ │ 2. UBS(10)   │ │
│  │ 数据截止...  │ │ 认购总规模    │ │ 3. 广发(9)   │ │
│  └──────────────┘ └──────────────┘ └──────────────┘ │
├──────────────────────────────────────────────────────┤
│  🔍 搜索与筛选                                       │
│  ┌──────────────────┐  ┌──────────┐  ┌─────────────┐│
│  │ 搜索基石投资者/公司│  │ 行业下拉  │  │ 金额 ≥/≤ 框 ││
│  └──────────────────┘  └──────────┘  └─────────────┘│
│  类别：[上下游企业] [政府基金] [老股东] [战略] [财务]    │
├──────────────────────────────────────────────────────┤
│  📋 结果列表（11 列可排序表格）                      │
│  ┌───┬──────┬──────┬────┬────┬──────┬──────┬────┬──────┬────┬──────┐ │
│  │ # │定价 │股票 │公司│行业│发行  │基石  │基石│基石  │类别│投资  │ │
│  │   │日期 │代码 │   │    │规模  │规模  │占比│投资者│    │金额  │ │
│  │   │     │     │   │    │(百万 │(百万 │    │      │    │(百万 │ │
│  │   │     │     │   │    │美元) │美元)  │    │      │    │美元) │ │
│  ├───┼──────┼──────┼────┼────┼──────┼──────┼────┼──────┼────┼──────┤ │
│  │ 1 │06-17│6675 │琻捷│... │ 981  │ 127  │29% │ ...  │... │ 9.57 │ │
│  │ 2 │ ... │ ... │... │... │ ...  │ ...  │... │ ...  │... │ ...  │ │
│  └───┴──────┴──────┴────┴────┴──────┴──────┴────┴──────┴────┴──────┘ │
│  [← 上一页]  第 1/35 页  [下一页 →]  每页 20 条     │
├──────────────────────────────────────────────────────┤
│  📥 底部操作栏                                       │
│  [导出筛选明细] [导出按投资者汇总] [重置筛选] [打印]    │
└──────────────────────────────────────────────────────┘

右侧滑出面板（点击投资者名称时）:
┌──────────────────┐
│ 泰康人寿          │
│ 参与 12 个项目    │
│ 总投资 $XXX M     │
│ ────────────────  │
│ ① 琻捷电子 5.20M │
│ ② 胜宏科技 30M   │
│ ③ ...            │
└──────────────────┘
```

### 4.3 功能模块详细说明

#### 模块 A：管理看板（统计卡片）

| 卡片 | 计算方式 | 示例 |
|------|---------|------|
| **项目总数** | `new Set(data.map(d => d.company_name)).size` | `44` |
| **基石总额** | `data.reduce((s, d) => s + d.amount_mn, 0)` 格式化 | `8,420.00` |
| **Top 3 最活跃** | 按投资者分组，按参与项目数降序取 3 | 泰康人寿(12次)… |

#### 模块 B：搜索与筛选

| 控件 | 类型 | 说明 |
|------|------|------|
| **搜索框** | 文本输入（`input`） | 同时匹配 `investor_name`、`company_name`、`stock_code`，实时过滤 |
| **行业下拉** | `<select>` | 动态从数据中提取行业列表，含"全部"选项 |
| **投资者类别** | 多选按钮或 `chips` | 上下游企业 / 政府基金 / 老股东 / 战略合作 / 财务投资者（可多选） |
| **金额范围** | 可选 min/max 输入 | 过滤 `amount_mn` 在区间内的记录 |

> 筛选逻辑：多个筛选条件之间为 **AND** 关系，同类别内为 **OR** 关系。

#### 模块 C：结果列表（11 列）

| 列名 | 数据字段 | 可排序 | 说明 |
|------|---------|--------|------|
| # | — | — | 序号（当前页起始序号） |
| 上市日期 | `listing_date` | ✅ | 完整日期 `YYYY-MM-DD`，AKShare 官方数据 |
| 股票代码 | `stock_code` | ✅ | 统一带 `.HK` 后缀 |
| 公司名称 | `company_name` | ✅ | — |
| 行业 | `industry` | ✅ | 显示主行业，hover 显示全部 |
| 发行规模<br>（百万美元） | `ipo_size_mn` | ✅ | 项目总发行规模 |
| 基石规模<br>（百万美元） | `cs_size_mn` | ✅ | 基石投资者认购总规模 |
| 基石投资者占比 | `cs_ratio` | ✅ | 基石规模占总发行规模比例 |
| 基石投资者 | `investor_name` | ✅ | **点击弹出详情侧栏** |
| 投资者类别 | `investor_category` | ✅ | 带颜色标签（badge） |
| 投资金额<br>（百万美元） | `amount_mn` | ✅ | 该笔投资金额 |

颜色标签映射（investor_category badge）：

| 类别 | 颜色 |
|------|------|
| 上下游企业 | `bg-blue-100 text-blue-800`（原"企业投资者"改名） |
| 政府基金 | `bg-purple-100 text-purple-800` |
| 老股东 | `bg-yellow-100 text-yellow-800` |
| 战略合作 | `bg-green-100 text-green-800` |
| 财务投资者 | `bg-orange-100 text-orange-800` |

> **表格样式规范**：所有单元格统一 `text-xs`、`text-gray-700`；数字列（金额/占比/规模等）用 `font-mono` **居中对齐**，文本列左对齐。表头 `py-2.5` 与数据行一致。
> 
> **多行表头排序箭头**：发行规模/基石规模/投资金额 三列用 `sortable-multi` 类，箭头 `<span class="srt-arrow">↕</span>` 内联在第一行右侧，JS 同步切换 `↕`→`↑`→`↓`；其他列仍用 CSS `::after`。

#### 模块 D：投资者详情侧栏

点击投资者名称时，从右侧滑出半屏面板：

```
┌────────────────────────────────┐
│ × 关闭                          │
│                                │
│ 🏢 泰康人寿                    │
│                                │
│   投资者类别: 财务投资者        │
│   参与项目: 12 个              │
│   总投资额: $380.50M           │
│   首次投资: 2026-01-07        │
│   最近投资: 2026-06-17        │
│                                │
│ ──── 参与项目列表 ────         │
│                                │
│ ① 琻捷电子          $5.20M    │
│    6675.HK · 元器件 · 2026-06 │
│                                │
│ ② 胜宏科技          $30.00M   │
│    2476.HK · PCB · 2026-04    │
│                                │
│ ③ ...                         │
└────────────────────────────────┘
```

#### 模块 E：导出 CSV（3 种导出方式）

| 按钮 | 位置 | 说明 |
|------|------|------|
| **导出筛选明细** | 顶部工具栏 | 导出当前筛选后的全部投资明细（一行一条），字段同表格列 |
| **导出按投资者汇总** | 顶部工具栏 | 按投资者分组汇总，含参与项目数、总投资额、项目列表文字 |
| **导出该投资者 CSV** | 侧栏底部 | 点击投资者名称打开详情后再点击，仅导出该投资者的项目明细 |

导出筛选明细示例：

```csv
listing_date,stock_code,company_name,industry,ipo_size_mn,cs_size_mn,cs_ratio,investor_name,investor_category,amount_mn
2026-06-17,6675.HK,琻捷电子,元线传感Soc,981,127,29%,欣旺达香港,上下游企业,5.20
2026-06-17,6675.HK,琻捷电子,元线传感Soc,981,127,29%,Oakwise,财务投资者,9.57
...
```

导出按投资者汇总示例：

```csv
investor_name,investor_category,project_count,total_amount_mn,project_list
易方达基金,财务投资者,3,18.5,长光辰芯（3 百万美元）；鸣鸣很忙（10 百万美元）；MINIMAX-WP（5.5 百万美元）
...
```

### 4.4 数据聚合函数（前端 JS 工具集）

```js
// hk_cornerstone_investors.html 中内置的聚合函数，供搜索/看板/详情侧栏调用

// 按字段分组
function groupBy(data, field)  // → { key: [records] }

// 按投资者聚合（含统计）
function aggregateByInvestor(data)
// → [{ investor_name, total_amount, project_count, first_date, last_date, projects: [...] }]

// 按公司聚合
function aggregateByCompany(data)
// → [{ company_name, stock_code, total_cs_amount, investor_count, investors: [...] }]

// Top N 排行
function topN(data, field, n)

// 格式化金额（始终保留两位小数）
function formatAmount(mn)  // → "9.57" 或 "1,230.00"

// CSV 序列化
function toCSV(data)
```

### 4.5 页面状态管理

所有筛选状态集中管理，每次变更触发统一渲染：

```js
const state = {
  keyword: "",           // 搜索关键词
  industry: "all",       // 行业筛选
  categories: [],        // 投资者类别（多选）
  amountMin: null,       // 金额下限
  amountMax: null,       // 金额上限
  sortField: "pricing_date",
  sortOrder: "desc",
  page: 1,
  pageSize: 20,
  selectedInvestor: null // 侧栏当前选中的投资者
};

function updateState(changes) { ... }  // 合并变更 → 触发 render()
function render() { ... }             // 根据 state 重新渲染所有模块
```

---

## 5. 别名清洗记录（2026-06-28 完成）

### 完成内容

2026年6月28日经用户逐项确认后，88 组别名规则已全部录入 `data/investor_aliases.json`，
并重新应用到全部数据。

### 操作步骤

```bash
# 1. 将别名规则应用到已有数据（无需重新解析 PPT）
python scripts/apply_aliases.py

# 2. 强制全量重跑（重建索引，确保一致性）
python scripts/extract_pptx.py --force
```

### 清洗效果

| 指标 | 清洗前 | 清洗后 |
|------|-------|-------|
| 总记录数 | 1,075 | 1,075 |
| 唯一投资者数 | 774 | **595**（减少 23%） |
| 别名规则 | — | 88 组（274 条 exact 匹配） |

### 关键合并示例

| 原始名称 | 规范化后 | 涉及变体数 |
|---------|---------|:---------:|
| HHLRA / 高瓴 / 高瓴资本 / 高瓴（HHLRA） | `Hillhouse 高瓴` | 11 个 |
| CPE / 源峰 / 源峰基金(CPE) / 源峰（银河TRS） | `CPE 源峰` | 6 个 |
| Boyu / 博裕 / Aqua Ocean（博裕） | `Boyu 博裕` | 6 个 |
| GIC / 新加坡政府投资公司 | `GIC 新加坡政府投资公司` | 4 个 |
| RBC / 加拿大皇家银行 | `RBC 加拿大皇家银行` | 5 个 |

### 以后更新流程

放入新 PPT → 双击 `更新数据.bat` → 增量提取 → `extract_pptx.py` **自动应用别名规则** → 刷新 HTML 即可。

> 注意：extract_pptx.py 在增量模式下会自动调用 `normalize_investor_name()` 清洗每条记录的投资者名称，
> 无需手动干预。如果新 PPT 中出现从未见过的投资者名称，它们会以原始名称通过，
> 需手动补充别名规则后重新运行。

```bash
# Step 1: 把新 PPT 放入 ppts/ 目录
# Step 2 (开发者): 命令行更新
python scripts/extract_pptx.py

# Step 2 (同事): 或直接双击 更新数据.bat

# Step 3: 输出 → data/data.js，刷新 HTML 即可

# Step 4: 提交版本（可选）
git add . && git commit -m "数据更新至 2026-06-28"

# Step 5: 用户刷新 hk_cornerstone_investors.html 即可
```

---

## 6. 文件清单（最终项目结构）

```
hk_ipo_csdata/
├── SKILL.md                  # 技能大纲（已建）
├── references.md             # 数据参考规范（已建）
├── scripts.md                # 代码施工蓝图（本文件）
├── CLAUDE.md                 # Claude Code 项目指令（后续生成）
│
├── data/
│   ├── data.js               # ★ 核心数据（由 extract_pptx.py 生成）
│   └── investor_aliases.json # 投资者别名映射（人工维护）
│
├── ppts/                     # ★ PPT 源文件放这里
│   └── README.md
│
├── scripts/
│   ├── extract_pptx.py       # PPT 提取脚本（增量/全量）
│   ├── apply_aliases.py      # 别名规则重应用（新规则确认后跑一次）
│   ├── check_new_names.py    # 新名称覆盖检测 + 疑似合并提醒
│   └── validate_data.py      # 数据校验脚本
│
├── 更新数据.bat              # ★ 双击即可更新数据（给同事用）
└── hk_cornerstone_investors.html    # 前端看板
```
