# Excel 窗口标签（excel-window-tabs）

在 Mac / Windows 桌面版 Excel 中，把同时打开的多个工作簿窗口当作"标签页"来切换：任务窗格里显示所有已打开的工作簿，点击即可切换；开启**对齐模式**后，所有窗口会被叠放到同一位置与大小，切换时视觉上接近 WPS 的单窗口多标签。

## 功能

- 任务窗格中列出所有已打开的 Excel 窗口，高亮当前活动窗口，点击即切换（乐观高亮，无需等待轮询）。
- 文档名默认**完整换行显示**（最多 3 行，超出用悬停 tooltip 看全名），可在底部一键切换"紧凑"单行省略模式。
- 对齐模式：开启时记录各窗口原始位置，并把所有窗口设为同一几何尺寸；关闭时恢复原位置；另提供"重新对齐"与"以当前窗口为基准"。
- 自适应轮询：面板聚焦 0.8s、失焦 3s、面板隐藏时完全暂停；操作后立即刷新，连续失败自动退避。
- 键盘操作：`⌘1–9`（Windows 为 `Ctrl+1–9`）快速切换，`↑`/`↓`/`Home`/`End` 移动，`Enter`/`Space` 激活，`Esc` 清除提示。
- 顶部提示条：成功提示自动消失；错误提示停留更久，并带"重试"与"复制"按钮。
- 新建工作簿、关闭指定窗口（走 Excel 原生保存提示）、手动刷新。
- "在此工作簿中自动打开"：把当前工作簿标记为打开时自动显示标签栏（设置随文件保存）。
- 运行环境检查：不满足 `ExcelApiDesktop 1.1` 时给出提示而不是报错。

## 环境要求

- 桌面版 Excel（Microsoft 365 for Mac 或 Windows），版本支持 `ExcelApiDesktop 1.1`（本机验证版本：Excel for Mac 16.113.1）。
- Node.js 18+（本机使用 v24.19.0；若工具链报错可改用 Node 20 LTS）。

## 目录结构

```
excel-window-tabs/
├── manifest.xml              # 开发用清单（指向 https://localhost:3000）
├── webpack.config.js         # 构建与 dev server（HTTPS + 开发证书）
├── src/taskpane/             # 任务窗格 UI 与逻辑
│   ├── taskpane.html/css
│   ├── taskpane.ts           # 引导、轮询、交互
│   ├── excel-window-api.ts   # Excel 窗口 API 封装
│   ├── logic.ts              # 纯函数（可单测）
│   └── state.ts              # localStorage 状态
├── src/commands/             # FunctionFile（TaskpaneApp 要求）
├── scripts/                  # 生产清单生成 / 侧载脚本
└── assets/                   # 16/32/80 图标
```

## 本地开发

```bash
npm install
npm test          # 运行纯逻辑单元测试
npm run validate  # 校验 manifest.xml
npm start         # 启动 HTTPS dev server 并自动侧载到 Excel
npm stop          # 结束调试会话
```

`npm start` 首次运行会通过 `office-addin-dev-certs` 生成并信任本地开发证书（macOS 会弹出钥匙串授权）。

## 发布到 GitHub Pages（日常使用）

1. 在 GitHub 新建一个公开仓库，例如 `excel-window-tabs`，把本项目推送到 `main`。
2. 构建并生成生产清单（把 URL 换成你的账号与仓库名）：

```bash
npm run build
npm run manifest:prod -- https://<用户名>.github.io/excel-window-tabs
```

3. 把 `dist/` 发布到 `gh-pages` 分支：

```bash
npx gh-pages -d dist
# 或使用项目内置命令
npm run deploy
```

4. 在仓库 **Settings → Pages** 中确认发布分支为 `gh-pages`，随后访问
   `https://<用户名>.github.io/excel-window-tabs/taskpane.html` 验证可以打开。
5. 侧载生产清单：

```bash
npm run sideload -- dist/manifest.xml
```

6. 注册加载项（Excel 只在打开带注册信息的文档时才会生成功能区按钮）：

```bash
npm run register
```

该命令会生成一个临时工作簿并在 Excel 中打开它；看到"开始"选项卡出现**标签栏**按钮后即可关闭这个临时文档。
每次修改 manifest 版本号后都需要重新执行第 5、6 步。

## 手动侧载（不使用 npm start）

把清单复制到 Excel 的加载项目录，然后重启 Excel：

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

打开 Excel → **插入 → 我的加载项 → 开发人员加载项**，选择对应的加载项；随后在"开始"选项卡点击**标签栏**按钮打开任务窗格。

## 使用说明

1. 打开多个工作簿（每个工作簿一个窗口）。
2. 在任一窗口点击**标签栏**打开任务窗格，列表会显示所有窗口。
3. 点击某一行即可切换到对应窗口；点击行尾 `×` 关闭该窗口（若有未保存修改，Excel 会询问）。
4. 点击**对齐：关/开**切换对齐模式：开启后所有窗口叠放到相同位置与大小，切换时看起来像一个窗口换标签；关闭会尽量恢复各窗口原始位置。
5. 对齐开启后会多出两个按钮：`⇥` 重新对齐（把基准重新套用到所有窗口）、`◎` 以当前窗口为基准。
6. 点击底部**显示：完整/紧凑**切换名称展示方式（选择会被记住）。
7. 面板获得焦点时可用 `⌘1–9` / `Ctrl+1–9` 快速切换窗口，`↑`/`↓` 选择、`Enter` 激活。
8. 勾选**在此工作簿中自动打开**并保存工作簿，之后打开该文件时会自动显示标签栏。

## 已知限制

- 不是原生标签：所有窗口仍是独立窗口，Mission Control / ⌘` 依旧会看到多个窗口。
- 标签栏只存在于打开了任务窗格的窗口；新窗口需要手动点一次"标签栏"，或对文档勾选自动打开。
- 没有窗口事件可订阅：列表靠自适应轮询更新（聚焦 0.8s / 失焦 3s / 隐藏暂停），外部改动最多约 3 秒后反映。
- 对齐时会以开启对齐那一刻的活动窗口位置为准；跨显示器、全屏或不同 Space 时几何设置可能失败，此时会退化为"只切换不对齐"。
- 需要 `ExcelApiDesktop 1.1`（桌面版 Microsoft 365）；Excel 网页版和旧版本不支持。

## 测试

```bash
npm test
```

覆盖 `logic.ts` 的纯函数：窗口排序、重名标签、几何比较、窗口差异检测。

### 手动验收清单（在 Excel for Mac 上逐条验证）

1. 打开 1 / 2 / 5 个工作簿窗口，标签数量与名称正确，点击切换 < 500ms。
2. 开启对齐后所有窗口位置尺寸一致（±1pt），切换后保持一致；关闭后恢复各自原始位置。
3. 关闭有未保存修改的工作簿窗口时出现 Excel 原生保存提示；选择取消则窗口保留且列表刷新。
4. 最小化的窗口被激活时自动恢复为普通状态。
5. 点击"＋ 新建"后新工作簿出现在列表中。
6. 轮询期间手动开关窗口，列表在 2 秒内同步且无报错。
7. 跨显示器 / 全屏 / 不同 Space 时记录实际行为；几何设置失败后应降级为仅切换并提示。
8. 勾选"在此工作簿中自动打开"并保存后，重新打开该工作簿时任务窗格自动出现。
9. 在 Excel 网页版打开加载项时应显示"当前环境不受支持"，且控制台无未捕获异常。

## v2 计划

- 标签颜色与分组、拖拽排序、MRU 排序。
- 让标签条在所有窗口可见：shared runtime 自动打开窗格，或悬浮对话框切换器（需先验证 dialog 是否可跨窗口常驻置顶）。
