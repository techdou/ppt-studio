# ppt-studio

全离线的 PPT / 课件 / 海报制作 skill。内置 python-pptx 本地引擎生成 `.pptx`，以 PPTD（YAML 语法）为中间格式，配合本机 PowerPoint COM 导出逐页截图做视觉 QA，附带零 npm 依赖的本地浏览器查看器。全链路无网络、无登录、无云服务依赖。

主要面向 AI coding agent 使用（`SKILL.md` 是 agent 的入口文档），也可以作为 Python 脚本库单独调用。

## 特性

- **双产物交付**：自包含 PPTD 项目目录（`.pptd` + `pages/` + `media/`）+ 本地引擎生成的 `.pptx`（默认逐页 fade 转场）
- **全离线**：生成、公式、图表、QA、人工查看全部在本地完成
- **讲义转课件**：成体系的讲义 / 教材 Markdown 一键批量转课件
- **公式**：LaTeX → MathML → OMML，走 Office 自带的 MML2OMML.XSL
- **图表**：plotly 原生渲染 bar / line / area / pie / doughnut / scatter / radar / bubble
- **视觉 QA**：本机 PowerPoint COM 导出逐页 PNG + 拼接总览图，交付前逐页检查排版
- **本地查看器**：浏览器里翻页预览、侧栏编辑文本元素（编辑即回写 `.page`），支持一键重导出

## 工作原理

PPTD 是 OOXML 的简化抽象层：用 YAML 描述主题、版式、元素位置与定义，去掉 Master 嵌套等复杂逻辑，每页自包含（完整格式定义见 [reference/pptd.md](reference/pptd.md)）。

```
输入（主题 / 文档 / 大纲 / 已有 pptx / 图片 / PDF / 讲义 Markdown）
        │
        ▼
  PPTD 项目目录  ◄──────── agent 生成 / 编辑 / 复刻 / 批量转换
        │
        ├──► scripts/export_pptx.py        python-pptx 本地引擎 → .pptx（fade 转场）
        ├──► scripts/export_images.py      PowerPoint COM 导图 + PIL 拼图 → 逐页 PNG / overview.jpg（视觉 QA）
        ├──► scripts/latex_omml.py         LaTeX 公式 → OMML（Office 自带 XSL）
        └──► ui/server.mjs                 本地查看器：预览 + 文本微调回写 .page
```

## 目录结构

```
ppt-studio/
├── SKILL.md                  # agent 入口：完整工作流与调用约定
├── reference/                # 设计参考文档
│   ├── pptd.md               # PPTD 格式定义
│   ├── fonts.md              # 字体体系
│   ├── shapes.md             # 形状库
│   ├── general-poster.md     # 海报场景
│   ├── slides_categories.md  # 场景设计索引
│   └── slides_categories/    # 7 个细分场景设计文档
├── scripts/                  # Python 工具链
│   ├── export_pptx.py        # PPTD → .pptx（python-pptx 本地引擎）
│   ├── export_pptx_local.py  # 引擎核心实现
│   ├── export_images.py      # PowerPoint COM 逐页导图（视觉 QA）
│   ├── handout_to_pptx.py    # 讲义 / 教材 Markdown → 课件批量转换
│   ├── latex_omml.py         # LaTeX 公式 → OMML
│   ├── page_text_tool.py     # .page 文本元素处理
│   └── requirements.txt
└── ui/
    ├── server.mjs            # 本地查看器服务（Node >= 18，零 npm 依赖）
    └── app/                  # 查看器前端
```

## 安装

```bash
# 1. 克隆到你的 agent 的 skills 目录（路径按实际 agent 调整，下同）
git clone https://github.com/techdou/ppt-studio.git ~/.agents/skills/ppt-studio

# 2. 安装 Python 依赖
pip install -r ~/.agents/skills/ppt-studio/scripts/requirements.txt
```

依赖说明：

- 核心：python-pptx >= 1.0、PyYAML
- 公式链：latex2mathml、lxml
- 图表链：plotly、kaleido、matplotlib（kaleido 出图需本机 Chrome / Chromium）
- QA 拼图：Pillow
- 图片 QA 需本机安装 Microsoft PowerPoint（COM）；无 PowerPoint 的环境跳过图片 QA，退化为结构化审查
- 本地查看器：Node.js >= 18，零 npm 依赖

## 使用

### 作为 agent skill

把本目录注册到 agent 的 skills 目录后，agent 读取 `SKILL.md` 即获得完整工作流：通读上下文 → 判定需求（Create / Edit / Replicate）→ 生成 PPTD → 校验与视觉 QA → 交付双产物。适用于：新建 PPT、编辑或美化已有 pptx、从图片 / PDF / 网页复刻幻灯片、讲义批量转课件、信息图与海报。

### 直接调用脚本

从 PPTD 项目生成 .pptx：

```bash
python scripts/export_pptx.py /abs/path/<项目名>/<项目名>.pptd \
  --output /abs/path/<项目名>/<项目名>.pptx --force
```

讲义 / 教材 Markdown 批量转课件：

```bash
python scripts/handout_to_pptx.py --input-dir <讲义根目录> --out-dir <输出根目录> \
  --label "小学 · 40 分钟"
```

导出逐页图片做视觉 QA（需本机 PowerPoint）：

```bash
python scripts/export_images.py /abs/path/<项目名>/<项目名>.pptd \
  --output /abs/path/<项目名>/.qa-images --force
```

### 本地查看器

```bash
node ui/server.mjs --project /abs/path/<项目名> --port 55280
```

打开 `http://127.0.0.1:55280/`：总览网格 + 单页大图 + 键盘翻页；侧栏列出当前页全部文本元素，编辑保存即回写 `.page`，改完点"重新导出预览"刷新。查看器只负责文本内容微调，布局 / 样式级修改（位置、颜色、字号）由 agent 直接改 `.page` 完成。

## 环境变量

| 变量 | 作用 |
|---|---|
| `PPT_STUDIO_MML2OMML` | 覆盖 Office 自带 MML2OMML.XSL 的定位路径 |
| `PPT_STUDIO_FA_SVGS` | Font Awesome svgs 目录（icon 渲染用）；默认向上搜索 `assets/fa/node_modules/@fortawesome/fontawesome-free/svgs` |

## 已知限制

- chart 的 waterfall / heatmap / treemap / sunburst / candlestick 输出占位框
- icon 以近似 emoji 渲染（约 60 个 FA 图标映射，未收录为 ★），非品牌图标
- LaTeX 公式退化为纯文本
- 贝塞尔曲线按折线渲染
- 云字体（如 MiSans）未安装时回退本机字体
- 图片 QA 依赖本机 Microsoft PowerPoint（COM）；kaleido 出图依赖 Chrome / Chromium

## 许可

尚未指定开源许可证，保留所有权利。如需在自有项目中使用，请先开 issue 联系。
