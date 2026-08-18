# HK IPO Cornerstone Investor Database

港股 IPO 基石投资者数据库。

本项目将定期更新的港股 IPO 基石投资者 PPT 转换为结构化数据，并通过
**GitHub Pages** 发布为可搜索的网页数据库。

## Website

网站通过 **GitHub Pages**
自动发布。更新数据库无需手动修改网页，也无需使用 Netlify。

------------------------------------------------------------------------

## How to Update the Database

### 日常更新：上传最新 PPT

进入 GitHub repository：

**`HK-IPO-DATA` → `ppts/`**

1.  点击 **Add file**
2.  点击 **Upload files**
3.  上传最新版本的基石投资者 PPT
4.  点击 **Commit changes**

上传完成后，GitHub Actions 会自动执行：

``` text
New PPT
   ↓
Extract PPT data
   ↓
Update structured database
   ↓
Generate website data
   ↓
Commit updated files
   ↓
GitHub Pages publishes updated website
```

正常情况下，不需要手动修改 `data.js`、`data.json` 或网页文件。

------------------------------------------------------------------------

## Checking an Update

上传 PPT 后进入：

**GitHub → Actions → Update website from PPT**

如果 workflow 显示绿色 **Success**，说明数据处理成功。GitHub Pages
随后会发布最新网站。

如果显示红色 **Failure**：

1.  点击失败的 workflow
2.  点击 `update-data`
3.  找到第一个红色步骤
4.  展开并查看最后的 error / traceback

黄色 warning 不等于 workflow 失败，应以最终 `Success / Failure`
状态为准。

------------------------------------------------------------------------

## Updating an Existing PPT

如果收到同一期 PPT 的修订版，建议使用**相同文件名替换原文件**。

内容发生变化后，系统会通过文件内容校验识别更新并重新处理。

## Adding a New PPT

如果是新一期数据，建议保留历史 PPT，并新增最新文件。

``` text
2026基石_0803.pptx
2026基石_0817.pptx
2026基石_0824.pptx
```

正常更新时，**建议保留历史 PPT，不要删除**。

## Important: Deleting PPTs

删除 `ppts/` 中的 PPT 不应被视为删除数据库历史记录的常规方法。

> 删除 PPT ≠ 一定自动删除该 PPT 已经写入数据库的数据。

如果某一期数据需要从数据库中彻底移除，应检查生成数据或执行完整的数据重建流程。

------------------------------------------------------------------------

## Repository Structure

``` text
HK-IPO-DATA/
├── .github/
│   └── workflows/
│       └── update-from-ppt.yml
├── ppts/
│   └── *.pptx
├── scripts/
│   ├── extract_pptx.py
│   ├── compile_ipo_details.py
│   ├── build_site.py
│   └── ...
├── data/
│   ├── data.json
│   ├── data.js
│   ├── company_info.json
│   └── ...
├── public/
├── index.html
├── requirements.txt
└── README.md
```

### `ppts/`

原始港股 IPO 基石投资者 PPT。日常更新主要操作这个文件夹。

### `scripts/`

负责 PPT 数据提取、清洗和网站数据生成的 Python
scripts。正常日常更新无需修改。

### `data/`

网站使用的结构化数据库文件。主要由程序自动生成，不建议日常手动修改。

### `.github/workflows/`

GitHub Actions 自动化配置。`update-from-ppt.yml` 负责 PPT
更新后的自动数据处理。

### `public/`

生成的网站发布文件。

------------------------------------------------------------------------

## Automated Update Workflow

``` text
Upload / Update PPT
        ↓
GitHub Actions
        ↓
Extract PPT data
        ↓
Compile IPO details
        ↓
Build website
        ↓
Commit generated data
        ↓
GitHub Pages publishes latest version
```

如果需要手动运行：

**GitHub → Actions → Update website from PPT → Run workflow → main → Run
workflow**

------------------------------------------------------------------------

## GitHub Pages

网站使用 **GitHub Pages** 托管。

Pages 设置位于：

**Repository → Settings → Pages**

网站更新流程不依赖 Netlify。

------------------------------------------------------------------------

## Recommended Weekly Workflow

``` text
Receive updated cornerstone PPT
            ↓
Upload to /ppts
            ↓
Commit changes
            ↓
Check GitHub Actions
            ↓
Confirm workflow = Success
            ↓
Check updated website
```

通常日常更新只需要：

> **Upload PPT → Commit → Check Actions → Check Website**

无需运行 Python、无需下载 GitHub Desktop，也无需手动重新部署网站。

## Notes

-   建议使用统一的 PPT 命名规则。
-   正常更新建议保留历史 PPT。
-   不建议手动修改自动生成的数据文件。
-   GitHub Action 失败时，应优先查看具体失败步骤，而不是反复重新运行。
-   GitHub Pages 是本项目当前正式的网站托管方式。
