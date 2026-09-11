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
- **默认组合（推荐）**：
  - **屏幕**：工作表背景水印 —— 打开即普通视图可见，不遮挡数据、**不拦截鼠标**，单元格照常点击输入
  - **打印**：页眉图片水印 —— 按纸张尺寸铺满整页，每页都有，同样不是浮动物体
  - 全程**没有任何浮动对象**，所以鼠标点击/输入完全不受影响
- **可选方式**（各有取舍，按需开启）：
  - **覆盖·文字**：浮动艺术字文本框，护眼模式 / 分页预览也能看到，但浮动物体会吃鼠标点击
  - **覆盖·图片**：平铺图片浮在单元格上方，效果与背景一致，但整块区域都会接住鼠标
  - **页眉文字水印**：页眉居中一行文字
  - **保护工作表 + 锁定对象**：配合覆盖水印，让水印不能被选中/拖动（所有单元格会设为未锁定，仍可编辑）
  - **页面布局视图**（`viewMode: pageLayout`）：打开即页面布局，页眉图片水印直接可见
  - **WPS 私有元数据**（实验性，默认关闭）：见下文
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
| `overlay` | boolean | `false` | 是否写入覆盖水印（浮动图片，任何视图可见且可打印，但会拦截鼠标） |
| `background` | boolean | `true` | 是否写入工作表背景水印（仅普通视图可见，不拦截鼠标） |
| `printImage` | boolean | `true` | 是否写入页眉图片打印水印（用 `&G` + VML，按页面尺寸铺满，每页打印） |
| `switchToNormalView` | boolean | `true` | 使用背景水印时，是否把分页预览/页面布局视图改回普通视图 |
| `printHeader` | boolean | `false` | 是否写入页眉文字水印（与 `printImage` 同时开启时以图片为准） |
| `wpsWatermark` | boolean | `false` | 实验性：写入 WPS 私有水印元数据（实测 WPS 12.1 不读取外部写入的描述） |
| `lockObjects` | boolean | `false` | 保护工作表并锁定对象（水印选不中），同时解锁全部单元格 |
| `viewMode` | string | `auto` | 打开视图：`auto` / `keep` / `normal` / `pageLayout` / `pageBreakPreview` |
| `wpsInvalidateBgImgs` | boolean | `true` | WPS 元数据里是否标记背景图待重建（与 WPS 自身输出一致） |
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

### 各方案的边界（来自开源社区与实测）

xlsx 文件格式里没有「水印」这个概念——Apache POI 社区答主 Axel Richter 说得最直白：
*"Watermark functionality is not available in Microsoft Excel."*（[Stack Overflow 原帖](https://stackoverflow.com/a/68559074)），
所以下面所有方案都是对水印的「模仿」，各有硬限制：

| 方案 | 屏幕可见 | 打印 | 吃鼠标 | 结论 |
| --- | --- | --- | --- | --- |
| 工作表背景图 `<picture>` | 仅**普通视图** | ✗（[Microsoft：无法打印工作表背景图形](https://support.microsoft.com/en-us/topic/you-cannot-print-a-background-graphic-for-a-excel-worksheet-7b1bbe1b-c672-a1bc-99e1-fb9eb8af5c45)） | 不吃 | **屏幕首选** |
| 页眉图片 `&G` + VML | 页面布局 / 打印预览 | ✓ 整页 | 不吃 | **打印首选**（[Microsoft 官方推荐的水印做法](https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009)，[XlsxWriter 官方示例](https://xlsxwriter.readthedocs.io/example_watermark.html)采用的正是「页眉放图片」） |
| 浮动文字 / 图片 | 所有视图（含护眼模式） | ✓ | **吃** | 仅在「必须任意视图可见」时用 |
| 背景图 + 保护工作表锁对象 | 同上 | 同上 | 对象选不中，但点击仍被对象接住 | 缓解不能拖动/误选 |

已知限制：
- **WPS 护眼模式**会在单元格背景之上铺一层护眼色，把工作表背景图盖住；WPS 自己添加的水印之所以能在护眼模式显示，是因为它读私有元数据、在渲染层自己绘制（外部工具写不出被 WPS 认可的版本，本工具已实测：仅写元数据 → WPS 完全不显示）。
- 页眉图片水印在 **Normal 视图看不到**，这是 Excel/WPS 的设计；如需打开即可见，把 `viewMode` 设为 `pageLayout`，或使用背景水印。

### WPS 私有水印（customXml，实验性）

WPS 表格的水印**不是**标准 OOXML 功能，它同时写两样东西：

1. `xl/media/*.png` + `<picture r:id>` —— 工作表背景图，给「其他软件」看的（也就是 Excel 里普通视图能看到的那层）
2. `customXml/itemN.xml` —— 命名空间 `http://www.wps.cn/officeDocument/2017/etCustomData` 的水印描述，
   并且**必须**在 `xl/_rels/workbook.xml.rels` 里用 `.../relationships/customXml` 指向它（少了这条关系 WPS 会读不到水印描述、完全不画水印）：

```xml
<watermarks xmlns="http://www.wps.cn/officeDocument/2017/etCustomData">
  <watermark type="0">
    <text fontName="PingFang SC" angle="-45" fontSize="24.000000" opacity="0.500000"><v>内部资料</v></text>
  </watermark>
  <invalidBgImgs>
    <invalidBgImg stId="1" hash="41703da32350a6aafcfad6726a3444ad"/>
    ...
  </invalidBgImgs>
</watermarks>
```

WPS 读这段描述，自己把水印画在**普通视图、分页预览、页面布局和打印**上，所以它在 WPS 里表现得像原生水印；
Excel 完全不认识这个命名空间，只会用第 1 项的背景图。本工具按 WPS 自己的输出结构写入这段元数据
（`customXml/itemN.xml` + `itemPropsN.xml` + `customXml/_rels/itemN.xml.rels` + 内容类型 Override），
并生成签名写入 `invalidBgImgs`，让 WPS 打开时按新参数重新生成背景图，避免两套水印叠在一起。

- `wpsWatermark: false` 可完全关掉这一项
- `wpsInvalidateBgImgs: false` 可不写 `<invalidBgImgs>`（备用方案：若某些 WPS 版本因此不画水印时使用）

### 打印水印（页眉图片）

1. 读取 `<pageSetup>`（纸张 `paperSize`、`orientation`）与 `<pageMargins>`（left/right/header/bottom），
   算出可打印区域：宽 = 纸宽 − 左右边距，高 = 纸高 − 页眉边距 − 下边距
2. 按这个尺寸（磅 → 像素，96 DPI）渲染同一张水印平铺图，写入 `xl/media/`
3. 新建（或复用）`xl/drawings/vmlDrawingN.vml`，写入 `<v:shape id="CH" ... width:Npt;height:Npt>` 与 `<v:imagedata o:relid>`；
   工作表里加 `<legacyDrawingHF r:id>`，页眉写 `&C&G`
4. `headerFooter` 设置 `scaleWithDoc="0"`，这样即使工作表设置了缩放（例如 `scale="53%"`）或「调整为 1 页宽」，
   页眉水印仍按实际纸张尺寸打印

这条路径与 xlsxwriter 的页眉图片实现结构一致（xlsxwriter 3.2.9 的输出实测对比：`shapelayout → shapetype
_x0000_t75 → shape id="CH" → imagedata → lock` 的元素序列与属性集合完全相同），差异只有三处且均为有意为之：

| 差异 | xlsxwriter | 本工具 | 原因 |
| --- | --- | --- | --- |
| `o:title` | `_wm`（取自图片文件名） | `Watermark` | 用固定名称，便于在「选择窗格」里识别 |
| `width` / `height` | `270pt` / `150pt`（源图 360×200 px 在 96 DPI 下的原始尺寸） | 按纸张可打印区域计算 | 目标是整页铺满，不是角落放一个 logo |
| `scaleWithDoc` | 未写（Excel 默认 `1`） | `"0"` | 工作表缩放时页眉水印仍按实际纸张尺寸打印 |

并且它不是浮动对象，所以打印有图、鼠标不受影响。

### 打印水印（页眉文字）

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
- 打印水印（页眉图片）：VML 部件 / `CH` 形状 / `legacyDrawingHF` 关系 / `&G` 占位 / `scaleWithDoc` / 元素顺序，且与 xlsxwriter 参考结构一致
- WPS 原生水印：元数据结构与 WPS 自己导出的文件逐节点一致，重复生成不堆叠，可关闭，且不影响其它读取器
- 水印图片本身：透明背景、文字居中、汉字字形正常（不是缺字方框）
- 前端：页面初始化、参数联动、上传解析、预览刷新、生成下载与提示

## 已知限制

- **仅支持 `.xlsx` / `.xlsm`**：`.xls`（BIFF 旧格式）、`.csv`、加密文件会返回明确错误提示
- **图表工作表（Chart Sheet）** 会被跳过，接口会通过 `X-Watermark-Skipped` 告知
- **背景图不会被 Excel 打印**，且只在「普通视图」显示（Excel 自身的限制）；打印水印由页眉图片负责，屏幕上则由背景图负责
- 覆盖水印是浮动对象：Excel / WPS 实测都不支持 `noSelect` 锁，点击图片覆盖范围会先选中图片而不是单元格，因此默认不开启；需要「任何视图都可见 + 可打印」且表格不需要输入时再用它
- 水印文字在服务端渲染成图片，字体取决于服务器上已安装的字体；跨平台部署时请确认目标字体存在
- 上传的文件只在内存中处理，不会写入磁盘，服务也不会保存任何副本

## 常见问题

**生成后看不到水印？**
按水印方式排查：
1. **默认组合（背景水印 + 页眉图片水印）**：屏幕上由背景水印负责，看不到多半是因为当前不是「普通视图」——背景图只在普通视图显示，分页预览 / 页面布局都看不到；打印时看不到，按下面第 3 条排查。
2. **覆盖水印**：任何视图都应可见。若看不到，请确认文件确实是用本工具重新生成的。
3. **打印水印**：只在打印预览 / 页面布局视图 / 实际打印时出现，普通视图看不到。默认的「页眉图片水印」会按纸张尺寸铺满整页；如果打印出来只有顶部一小块，说明打印机/驱动忽略了 VML 中的尺寸，可改用「页眉文字水印」。
4. **背景水印单独排查**：如果文件原本保存为「分页预览 / 页面布局」视图，打开时就看不到（本工具默认会把它切回普通视图，可在页面取消勾选）；另外，单元格若有填充色会盖住背景图。

**水印里有汉字显示成方框 / 某些字不显示？**
这是渲染字体缺少该字字形导致的，已在水印文字渲染时自动处理：会换成实测可用的中文字体，
预览下方与下载提示里会显示「实际渲染字体」以及是否发生了替换。如果仍然出现方框，
说明该字符在本机所有字体里都缺字形（例如罕见的 CJK 扩展区汉字），可以换成常用字。

**开了覆盖水印后表格点不动、不能输入？**
这是浮动物体的正常行为（Excel/WPS 都会用图片接住鼠标）。两种处理：切到「位于单元格后方」重新生成，或在文件里按 `开始 → 查找和选择 → 选择窗格` 删除名为 `Watermark` 的对象。**默认组合（背景 + 页眉图片）不会出现这个问题。**

**为什么 WPS 里分页预览也有水印，Excel 里没有？**
WPS 读的是它自己写在 `customXml` 里的水印描述，会在所有视图里自己绘制水印；Excel 不认识这段元数据，
只能用工作表背景图（仅普通视图）和页眉水印（打印）。所以本工具两种都写：WPS 用户得到全视图水印，
Excel 用户得到「普通视图可见 + 打印可见」，互不干扰。

**怎么删掉水印？**
覆盖水印：`开始 → 查找和选择 → 选择窗格`，删除名为 `Watermark` 的对象（图片已设为不可直接选中）。
背景水印：`页面布局 → 删除背景`。打印水印：`页面布局 → 打印标题 → 页眉/页脚` 中删掉那一行文字。

**页眉里的文字被截断？**
页眉文字过长时会被页宽限制，建议缩短文字或减小页眉字号。

**端口被占用？**
`PORT=8080 npm start`，或先结束占用进程。

## 参考资料

水印相关实现与限制的出处（整理时已逐条访问核对，链接均为 200）：

| 资料 | 用途 |
| --- | --- |
| [Microsoft：Add a watermark in Excel](https://support.microsoft.com/en-us/office/add-a-watermark-in-excel-a372182a-d733-484e-825c-18ddf3edf009) | 官方给出的水印做法：在页眉/页脚插入图片 |
| [Microsoft：You cannot print a background graphic for a Excel worksheet](https://support.microsoft.com/en-us/topic/you-cannot-print-a-background-graphic-for-a-excel-worksheet-7b1bbe1b-c672-a1bc-99e1-fb9eb8af5c45) | 背景图无法打印（Excel 设计如此） |
| [Microsoft：Add or remove a sheet background](https://support.microsoft.com/en-us/office/add-or-remove-a-sheet-background-3577a762-8450-4556-96a2-cc265abc00a8) | 工作表背景图的行为说明 |
| [Stack Overflow：Apache POI - watermark in Excel（Axel Richter 的回答）](https://stackoverflow.com/a/68559074) | 「Excel 里没有水印功能」及各类模仿方案的取舍 |
| [Stack Overflow：Apache POI - adding watermark in Excel workbook（Axel Richter 的回答）](https://stackoverflow.com/a/51103756) | 用页眉/页脚图片模仿水印（打印水印）的做法 |
| [XlsxWriter：Example: Setting a Worksheet Watermark](https://xlsxwriter.readthedocs.io/example_watermark.html) | Python 版「页眉图片水印」参考实现；本工具用 xlsxwriter 3.2.9 生成对照文件做结构比对 |
| [rust_xlsxwriter：Adding a watermark as a header image](https://rustxlsxwriter.github.io/examples/watermark.html) | 同一做法的 Rust 实现，供交叉参考 |

> WPS 私有水印（`customXml` + `etCustomData`）没有公开规范，本文档中的结构说明来自对 WPS 自身输出文件的解包分析，属实测结论。

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
