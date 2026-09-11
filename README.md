# Excel 水印工具

一个 Node 服务 + 网页界面：上传 `.xlsx` / `.xlsm`，自定义水印文字与样式，实时预览，一键下载带水印的文件。

直接在 OOXML（zip）层面写入水印，**不改动任何单元格、公式、样式、图表、批注和宏**。

```
┌──────────────────────────┐        ┌────────────────────────────────────┐
│  浏览器（public/）        │  HTTP  │  Node 服务（server.js）             │
│  上传 / 调参 / 实时预览   │ ─────► │  ├─ canvas 生成水印平铺图 (PNG)     │
│  下载带水印的 xlsx        │ ◄───── │  └─ JSZip 注入 OOXML 元素           │
└──────────────────────────┘        └────────────────────────────────────┘
```

## 功能

- **上传即用**：拖拽或点击选择文件，自动列出所有工作表并支持按表选择
- **水印内容可调**：文字（支持换行）、字体、字号、颜色、透明度、旋转角度、间距、加粗 / 斜体 / 边框
- **三种水印方式**
  - **覆盖水印（默认）**：写入一张透明底的浮动图片盖在单元格上方，普通视图 / 分页预览 / 页面布局都能看到，且会随文件打印；图片已锁定为不可选中，正常点击单元格不受影响
  - **背景水印**：写入工作表背景图，显示在单元格后方，完全不遮挡数据（即 Excel「页面布局 → 背景」的效果）；注意它只在**普通视图**可见、不能打印，因此开启时会把工作表的打开视图切回普通视图
  - **打印水印**：写入页眉居中文字，打印 / 页面布局视图每页显示，原有页眉内容自动保留
- **实时预览**：预览图由服务端用与生成文件完全相同的逻辑渲染，页面按 1:1 平铺模拟 Excel 效果，并显示水印单元尺寸与重复密度
- **无数据损失**：只新增/替换水印相关的 XML 片段与关系，其余部件原样保留（含 `.xlsm` 宏）
- **三种风格预设**：浅淡 / 标准 / 醒目

## 快速开始

```bash
npm install       # 安装依赖（@napi-rs/canvas 为预编译二进制，无需本地编译）
npm start         # 默认 http://127.0.0.1:3210
```

打开 <http://127.0.0.1:3210> 即可使用。

自定义端口 / 监听地址：

```bash
PORT=8080 HOST=0.0.0.0 npm start
```

开发模式（文件变更自动重启）：`npm run dev`

## 页面使用步骤

1. **选择 Excel 文件**：点击或拖拽 `.xlsx` / `.xlsm` 文件到虚线区域（单个文件最大 60MB）
2. **勾选工作表**：默认全选，可单独取消；图表工作表会被标记为「跳过」
3. **编辑水印内容**：输入文字（可用内置快捷词）、选择字体、调整字号 / 颜色 / 透明度 / 旋转角度 / 间距 / 样式
4. **选择水印方式**：背景水印与打印水印可单独或同时开启；打印水印可设置页眉字体与字号
5. **生成并下载**：点击「生成并下载」，浏览器会保存 `原文件名-水印.xlsx`

> 右侧预览区会实时刷新：上方模拟工作表背景水印，下方模拟打印页眉水印。
> 背景水印显示在单元格**后方**，如果单元格有填充色会遮住它——需要纸质文件带水印时请同时开启「打印水印」。

## HTTP 接口

水印参数统一通过查询参数 `c` 传递：`c = base64url(JSON.stringify(config))`。

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| GET | `/` | 网页界面 |
| GET | `/api/config` | 返回默认参数与字体建议列表 |
| GET | `/api/preview?c=...` | 返回水印平铺图（PNG，透明背景），响应头含 `X-Tile-Width` / `X-Tile-Height` |
| POST | `/api/analyze?name=...` | 请求体为 xlsx 原始字节，返回 `{ sheets: [{ name, state, kind, supported }] }` |
| POST | `/api/watermark?c=...&name=...` | 请求体为 xlsx 原始字节，返回加水印后的文件（`Content-Disposition` 为附件） |

成功响应头还会返回 `X-Watermark-Sheets`（已写入的工作表）与 `X-Watermark-Skipped`（跳过的工作表），均为 `encodeURIComponent` 编码。

出错时返回 JSON：`{ "ok": false, "error": "错误说明" }`。

### curl 示例

```bash
# 1) 生成参数（base64url）
CFG=$(python3 -c "import base64,json;print(base64.urlsafe_b64encode(json.dumps({
  'text':'内部资料 请勿外传','fontSize':36,'opacity':0.3,'rotate':-30,'gap':90,
  'color':'#c0392b','background':True,'printHeader':True,'sheets':'all'
},ensure_ascii=False).encode()).decode().rstrip('='))")

# 2) 查看工作表
curl -X POST --data-binary @报表.xlsx \
  "http://127.0.0.1:3210/api/analyze?name=报表.xlsx"

# 3) 加水印并保存
curl -X POST --data-binary @报表.xlsx -o "报表-水印.xlsx" \
  "http://127.0.0.1:3210/api/watermark?c=$CFG&name=报表.xlsx"

# 4) 只取预览图
curl -o tile.png "http://127.0.0.1:3210/api/preview?c=$CFG"
```

## 水印参数

| 字段 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | string | `内部资料` | 水印文字，支持 `\n` 换行（最多 6 行、200 字） |
| `fontFamily` | string | `sans-serif` | 渲染水印图片使用的字体族 |
| `fontSize` | number | `34` | 字号（px），8–200 |
| `color` | string | `#8a8a8a` | 颜色，`#rgb` 或 `#rrggbb` |
| `opacity` | number | `0.3` | 透明度，0.02–1（烘焙进 PNG 的 alpha 通道，兼容 Excel / WPS） |
| `rotate` | number | `-30` | 旋转角度，-90–90（负数逆时针） |
| `gap` | number | `90` | 水印间距（px），0–400，越大越稀疏 |
| `bold` / `italic` / `border` | boolean | `true` / `false` / `false` | 加粗 / 斜体 / 文字边框 |
| `overlay` | boolean | `true` | 是否写入覆盖水印（浮动图片，任何视图可见且可打印） |
| `background` | boolean | `false` | 是否写入工作表背景水印（仅普通视图可见） |
| `switchToNormalView` | boolean | `true` | 使用背景水印时，是否把分页预览/页面布局视图改回普通视图 |
| `printHeader` | boolean | `true` | 是否写入页眉文字水印 |
| `headerFont` | string | `宋体` | 页眉水印字体（Excel 字体名） |
| `headerFontSize` | number | `22` | 页眉水印字号（pt），6–72 |
| `sheets` | `'all'` \| string[] | `'all'` | 作用的工作表名称列表 |

服务端会统一做范围裁剪与转义，非法值自动回落到默认值。

## 实现原理

### 背景水印

1. 用 `@napi-rs/canvas` 渲染**一块瓷砖**：旋转后的文字 + 四周留白，透明背景，透明度直接写进 alpha 通道
2. 图片写入 `xl/media/imageN.png`，并在 `[Content_Types].xml` 中确保存在 `png` 的 Default 声明
3. 每个工作表 XML 中插入 `<picture r:id="rIdN"/>`，并在 `xl/worksheets/_rels/sheetN.xml.rels` 中新增图片关系

`<picture>` 的位置必须满足 SpreadsheetML 的元素顺序（`… drawing → legacyDrawing → legacyDrawingHF → picture → oleObjects → … → tableParts → extLst`）。实现使用一个轻量的顶层元素扫描器，把新元素插到「必须位于其后的第一个已存在元素」之前，因此不会因为文档里存在 `tableParts`、`extLst`、或 `sheetPr` 内部嵌套的 `extLst` 而插错位置。

如果工作表本来就有背景图，则复用原 `r:id` 并替换其关系目标，不会产生两个 `<picture>`。

### 覆盖水印

1. 按工作表已用区域（`<dimension>` + 列宽 `<cols>` + 行高 `defaultRowHeight`/`ht`）算出像素尺寸，用同一块瓷砖平铺出一整张透明底 PNG
2. 图片写入 `xl/media/`，并挂到工作表的 drawing 部件上：
   - 工作表原本没有 `<drawing>` → 新建 `xl/drawings/drawingN.xml` 与对应关系、内容类型
   - 已有 `<drawing>`（例如本来就有图表或图片）→ 复用该部件，直接追加锚点，不会产生第二个 `<drawing>`
3. 锚点使用 `oneCellAnchor`：锚定 A1、尺寸以 EMU 固定（1px = 9525 EMU），因此水印不会随列宽行高被拉伸变形
4. 图片设为 `noSelect="1"`，尽量让点击穿透到单元格；重复加水印时会替换旧锚点而不是叠加，并清理不再被引用的 `xl/media` 图片

### 字体与缺字处理

水印文字是服务端用 canvas 渲染成图片的，所以字体是否包含汉字字形非常关键。部分环境里
canvas 只能拿到日文/韩文字体（例如 macOS 上 `PingFang SC` 无法注册，请求 `sans-serif`
会回退到 Hiragino Sans），此时「请、传、资、样」这类**简体专用字**会渲染成方框（豆腐块）。

处理方式：

1. 启动时按平台尝试注册全字库中文字体（macOS：Arial Unicode / Hiragino Sans GB / Songti；
   Windows：微软雅黑 / 黑体 / 宋体；Linux：Noto Sans CJK / 文泉驿 等），注册名 `Watermark CJK`
2. 用「私用区字符 `U+E000` 的位图」作为豆腐块参照，逐字比对来判断某个字体是否覆盖文本
3. 渲染前先探测用户所选字体：覆盖不全就自动换成实测可用的中文字体，
   并在响应头 `X-Watermark-Font` / `X-Watermark-Font-Fallback` 里返回实际字体与替换标记，
   页面上会给出「已自动改用 XXX」的提示
4. `GET /api/config` 返回的字体列表**只包含本机实测可用的字体**，并标注「仅英文」的拉丁字体

因此即使在字形不全的环境里，水印文字也不会出现方框。

### 打印水印

在 `<headerFooter>` 的 `oddHeader`（必要时还有 `evenHeader` / `firstHeader`）中追加：

```
&C&"宋体,常规"&22&K8A8A8A内部资料
```

- `&C` 居中，`&"字体,字形"` 指定字体，`&22` 指定字号，`&K` 指定颜色
- 文字里字面量的 `&` 会写成 `&&`，避免被 Excel 当成控制符
- 原有页眉内容会保留，水印以新的一行追加在其后；`differentOddEven` / `differentFirst` 存在时会同步处理偶数页与首页
- 若工作表没有 `<pageMargins>`，会补充 Excel 的默认页边距，保证打印效果正常

## 项目结构

```
excel-watermark/
├── LICENSE                    # Apache License 2.0
├── server.js                  # HTTP 服务：静态页面 + API
├── src/
│   ├── config.js              # 参数定义、默认值与校验
│   ├── watermark-image.js     # canvas 渲染水印平铺图
│   ├── xlsx-watermark.js      # OOXML 注入（背景图 / 页眉水印）
│   └── xml.js                 # 轻量 XML 工具（顶层元素扫描、按序插入）
├── public/
│   ├── index.html             # 页面
│   ├── style.css              # 样式（含深色模式、响应式）
│   └── app.js                 # 上传、预览、下载交互
└── test/
    ├── make-fixtures.py       # 生成测试文件（openpyxl / xlsxwriter / 手工构造）
    ├── run-tests.mjs          # 核心与水印图片测试
    ├── ui-smoke.mjs           # 前端 jsdom 冒烟测试（含真实上传下载流程）
    └── verify_with_openpyxl.py
```

## 测试

```bash
npm test          # 核心 + 前端
npm run test:core # 只跑核心（结构、数据完整性、幂等、第三方读取器交叉验证）
npm run test:ui   # 只跑前端 jsdom 冒烟测试
```

`test/fixtures` 下的测试文件不存在时，核心测试会自动调用 `python3 test/make-fixtures.py` 生成（需要 `openpyxl`；`xlsxwriter` 可选，用于生成「已有背景图 + 表格 + 页眉」的样本）。测试覆盖：

- 生成的 XML 元素顺序、关系文件、内容类型、图片路径
- 单元格数据、公式、合并单元格、超链接、图表、批注保持不变
- 连续加水印两次不会插入重复元素（幂等）
- 只作用于指定工作表、名称不存在时的行为
- ExcelJS / SheetJS / openpyxl 三个独立读取器都能正常打开输出文件
- 覆盖水印：drawing 部件 / 内容类型 / `oneCellAnchor` 锚点 / 图片关系，已有图表的工作表也能正确追加
- 水印图片本身：透明背景、文字居中、汉字字形正常（不是缺字方框）
- 前端：页面初始化、参数联动、上传解析、预览刷新、生成下载与提示

## 已知限制

- **仅支持 `.xlsx` / `.xlsm`**：`.xls`（BIFF 旧格式）、`.csv`、加密文件会返回明确错误提示
- **图表工作表（Chart Sheet）** 会被跳过，接口会通过 `X-Watermark-Skipped` 告知
- **背景图不会被 Excel 打印**，且只在「普通视图」显示（Excel 自身的限制）；需要打印或需要在任意视图看到，请使用默认的「覆盖水印」
- 覆盖水印会随文件打印，且是浮动对象：如果某些环境不支持 `noSelect` 锁，点击水印所在区域会先选中图片，此时按 Esc 或改用「背景水印」模式
- 水印文字在服务端渲染成图片，字体取决于服务器上已安装的字体；跨平台部署时请确认目标字体存在
- 上传的文件只在内存中处理，不会写入磁盘，服务也不会保存任何副本

## 常见问题

**生成后看不到水印？**
按水印方式排查：
1. **覆盖水印（默认）**：任何视图都应可见。若看不到，请确认文件确实是用本工具重新生成的。
2. **背景水印**：只显示在「普通视图」。如果文件原本保存为「分页预览 / 页面布局」视图，打开时就不会显示（本工具默认会把它切回普通视图，可在页面取消勾选）。另外，单元格若有填充色会盖住背景图。
3. **打印水印**：只在打印预览 / 页面布局视图 / 实际打印时出现，普通视图看不到。

**水印里有汉字显示成方框 / 某些字不显示？**
这是渲染字体缺少该字字形导致的，已在水印文字渲染时自动处理：会换成实测可用的中文字体，
预览下方与下载提示里会显示「实际渲染字体」以及是否发生了替换。如果仍然出现方框，
说明该字符在本机所有字体里都缺字形（例如罕见的 CJK 扩展区汉字），可以换成常用字。

**怎么删掉水印？**
覆盖水印：`开始 → 查找和选择 → 选择窗格`，删除名为 `Watermark` 的对象（图片已设为不可直接选中）。
背景水印：`页面布局 → 删除背景`。打印水印：`页面布局 → 打印标题 → 页眉/页脚` 中删掉那一行文字。

**页眉里的文字被截断？**
页眉文字过长时会被页宽限制，建议缩短文字或减小页眉字号。

**端口被占用？**
`PORT=8080 npm start`，或先结束占用进程。

## 许可证

本项目基于 [Apache License 2.0](LICENSE) 授权，完整条款见仓库根目录的 `LICENSE` 文件。

Copyright 2026 bdoycn

第三方依赖的许可：

| 依赖 | 用途 | 许可 |
| --- | --- | --- |
| `@napi-rs/canvas` | 服务端渲染水印平铺图 | MIT |
| `jszip` | 读写 OOXML（zip）容器 | MIT（或 GPL-3.0-or-later，本项目按 MIT 使用） |
| `exceljs`（dev） | 测试中校验生成文件 | MIT |
| `xlsx`（dev） | 测试中校验生成文件 | Apache-2.0 |
| `jsdom`（dev） | 前端冒烟测试 | MIT |
