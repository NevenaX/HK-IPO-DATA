[README.md](https://github.com/user-attachments/files/31128573/README.md)
# HK Cornerstone Investors — 网页部署与 PPT 更新

这个目录已经可以直接部署为静态网站。`index.html` 是网站首页；核心数据来自 `data/data.js`。

## 一次性上线

1. 在 GitHub 新建一个 repository。
2. 把本目录全部文件上传到 repository 根目录。
3. 在 Netlify 选择 **Add new site → Import an existing project → GitHub**，选择这个 repository。
4. Build command 和 Publish directory 会由 `netlify.toml` 自动处理；Netlify 只发布 `public/`，不会把 `ppts/` 源文件公开到网站。部署后 Netlify 会生成固定网址。
5. 可在 Netlify 的 Domain management 中修改 `*.netlify.app` 子域名。

## 以后怎么更新 PPT

只需要在 GitHub 打开 `ppts/` 文件夹，点击 **Add file → Upload files**，上传新的 `.pptx`，然后 Commit changes。

GitHub Actions 会自动执行：

`新 PPT → extract_pptx.py → data/data.json + data/data.js → compile_ipo_details.py → standalone HTML → 自动 commit`

Netlify 检测到这个自动 commit 后会重新部署，因此网站会更新。

## 重要规则

- 新 PPT 请放在 `ppts/`，不要放在根目录。
- 不要上传 Office 临时文件（文件名通常以 `~$` 开头）。
- 最好保持现有 PPT 表格的 13 列结构不变。
- `data/investor_aliases.json` 用于统一同一投资者的不同写法；修改它也会触发自动更新。
- `company_info.js/json` 的 AKShare 在线补充数据没有放进强制自动流程，避免外部接口故障导致 PPT 更新失败。需要刷新公司资料时可在本地单独运行 `python scripts/fetch_company_info.py`。

## 查看自动更新是否成功

GitHub repository → **Actions** → `Update website from PPT`。

绿色对勾表示数据已经处理并写回 repository；随后 Netlify 会按新的 commit 部署。
