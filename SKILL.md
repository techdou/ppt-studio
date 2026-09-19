---
name: ppt-studio
description: "Create, edit, replicate, and export presentations (PPT / slides / 演示文稿) fully offline with a built-in python-pptx engine — no cloud service or login required. Use for any presentation task: 新建 PPT、幻灯片、汇报/答辩/课件 slides，编辑或美化已有 pptx，从图片/PDF/网页复刻幻灯片，讲义/教材 Markdown 批量转课件 handout-to-slides，信息图 infographic，海报 poster。Not for: 仅提取 PPT 文字内容、忠实翻译幻灯片、仅摘要或信息整理，除非用户明确要求处理演示文稿。"
metadata:
  version: "3.1.0"
  engine: "local (python-pptx) + plotly + OMML"
---

# 定义

ppt-studio 是围绕 PPTD 格式的演示文稿创作与导出 skill。PPTD 是 OOXML 的简化抽象层（YAML 语法），保留主题、版式、元素位置与定义等核心信息，去掉 Master 嵌套等复杂逻辑，每页自包含。完整格式定义见 `reference/pptd.md`。

**路径约定**：下文所有命令中的 `<skill_dir>` 指本 skill 的安装根目录（即本 SKILL.md 所在目录）。Agent 触发本 skill 时已知该绝对路径，调用时直接代入，不要照抄字面量。

全链路（生成、导出、QA、公式、人工查看）均在本机完成，无网络、无登录依赖：

| 环节 | 实现 | 网络 |
|---|---|---|
| PPTX 生成 | `scripts/export_pptx.py`（python-pptx 本地引擎） | 无 |
| 图片 QA | `scripts/export_images.py`（本机 PowerPoint COM + PIL 拼图） | 无 |
| 公式 | `scripts/latex_omml.py`（Office 自带 MML2OMML.XSL） | 无 |
| 人工查看/微调 | `ui/server.mjs` 本地查看器（Node 静态服务） | 无 |

## 双产物默认

除非用户明确选择退出，任何 PPT 任务交付：

1. 自包含 PPTD 项目目录（`.pptd` + `pages/` + `media/` 等全部引用依赖），永不只交一个孤立 manifest；
2. 本地引擎生成的 `.pptx`（默认逐页 fade 转场）。

已有 PPTX 可转换为 PPTD 后编辑，交付时同样双产物。

## PPT 生产工作流

### step1. 通读上下文
读用户上传的全部文件、给定 URL，以及格式指南 `reference/pptd.md`。

### step2. 判定需求
1. 目的：Create（新建）/ Edit（编辑已有 pptx）/ Replicate（从图片、PDF 等非 pptx 格式复刻为 pptd）
2. 设计方向：自主设计（需补全设计）/ Design system（用户给全套颜色、字体、版式规范）/ 使用模板 / 风格迁移（提供参考图或网页）
3. 输入类型：仅主题 / 完整文档 / 大纲——大纲与完整文档默认可检索扩充，除非用户明确禁止
4. 页数：用户指定优先；大纲匹配页数；仅主题按内容自定

路由：输入是成体系的讲义/教材 Markdown 且需批量出课件时，直接走下方「讲义/教材 → 课件 PPT（批量）」专用路径，不进本工作流。

### step3. 生成
生成前先读 `reference/pptd.md` 掌握格式与约束，按需分层：§1–4（全局约定、共享类型、入口与页面文件）必读；§5 Elements 占全文大头，按当页用到的元素类型读对应小节（text / shape / line / image / icon / table / chart），不必整章通读；step4 校验同理。

设计参考按需读：`reference/slides_categories.md`（场景设计索引，细分文档在同目录 `slides_categories/`）、`reference/fonts.md`（字体体系）、`reference/typography.md`（字号层级、行高与间距规范，调整字号必读）、`reference/shapes.md`（形状库）、`reference/general-poster.md`（海报场景）。

- **Replicate**：分析图片估计元素位置、字体与字号，尽量 1:1 复刻；无法用形状近似的照片/头像可从原图裁切为媒体资产。
- **Edit**：将上传的 pptx 转为 pptd（或直接用 python-pptx 编辑），审查转换后页面结构与关键视觉细节，只动目标范围，不碰范围外内容。
- **Generate**：按设计方向执行——自主设计先读对应场景文档；Design system 以用户设计文档为唯一风格依据；使用模板先转模板 pptd 再按模板风格生成；风格迁移分析参考源的视觉特征并复用其插图、字号层级与元素。

### step4. 校验
1. 对照 `reference/pptd.md` 校验生成的 pptd（必填字段、类型、边界、theme token、资源路径），多轮修复。
2. **视觉 QA（交付前必做）**——导出逐页图片审查：

   ```bash
   python <skill_dir>/scripts/export_images.py \
     /abs/path/<项目名>/<项目名>.pptd \
     --output /abs/path/<项目名>/.qa-images --force
   ```

   产出每页 PNG（`pages/1.png…N.png`）与拼接总览 `overview.jpg`。逐页检查：图片清晰不变形；文字不压关键画面；元素坐标不越界；边界与配色对比足够；排版统一（对齐、间距、页边距；字号层级对照 `reference/typography.md`）；文本不溢出文本框；内容不被上层元素遮挡。

   可疑页读全分辨率 `.qa-images/pages/<n>.png` 确认后修改对应 `.page`，`--force` 重跑复查，直到每页通过。

   无 PowerPoint 的环境跳过图片 QA，改为对 pages 的结构化审查（bounds/溢出/对比/层级），并在交付说明中注明跳过了图片 QA。

### step5. 交付
1. 项目目录自包含，永不交付孤立 manifest。**按内容自主命名**，不使用统一的 `deck`：

   命名规则：目录与文件取同一基名。优先用用户明确指定的名称；无指定名时，依内容主动拟定标题并写入 pptd manifest 的 `title`，取其作基名（新建场景即由 Agent 依上下文自主命名，不要等用户给名）；编辑已有 pptx 时沿用原文件基名。清洗 Windows 非法字符（`\/:*?"<>|`）与首尾空格，超 40 字符截断；确实无法确定标题时才回退 `deck`。

   ```text
   <项目名>/
     <项目名>.pptd
     pages/
       *.page
     media/
       *
     .qa-images/         # QA 导出图片（可选保留）
     <项目名>.pptx      # 默认生成
   ```

2. 用本地引擎生成 `.pptx`：

   ```bash
   python <skill_dir>/scripts/export_pptx.py \
     /abs/path/<项目名>/<项目名>.pptd \
     --output /abs/path/<项目名>/<项目名>.pptx
   ```

   选项：`--transition fade|none`（默认 fade，逐页写入）、`--force`（覆盖输出）。

3. 交付前核验：输出文件存在、页数正确、每页恰有一个 fade 转场（`p:transition/p:fade` 位于 `p:cSld` 之后）。用可点击的绝对路径链接交付：项目目录、`.pptd`、`pages/`、`media/`、`.pptx`。

4. **本地引擎支持范围**：text 全套富文本（p/span/strong/em/u/s/sub/sup/ul/ol/li/a/br、text-align/line-height/margin-top/font-size/color/background-color）；shape 40+ 常用形状 + custom path（M/L/H/V/C/Z）；image fit cover/contain/fill + crop + cropShape(roundRect/ellipse)；table 行列合并与 firstRowStyle/bodyStyles 样式链；line 直线/箭头；icon 以近似 emoji 渲染（约 60 个 FA 图标映射，未收录为 ★）；chart 原生支持 bar/line/area/pie/doughnut/scatter/radar/bubble；theme $ref 全链路；背景 solid/gradient；fade 转场；speaker notes。

5. **已知限制**（导出时自动打印告警）：chart 的 waterfall/heatmap/treemap/sunburst/candlestick 输出占位框；icon 为 emoji 近似非品牌图标；LaTeX 公式退化为纯文本；贝塞尔曲线按折线渲染；云字体（MiSans 等）未安装时回退本机字体。

## 讲义/教材 → 课件 PPT（批量）

**适用**：输入是成体系的讲义/教材 Markdown（H1 课名 + H2 环节节 + 图文交叉），需要批量生成每课课件。设计原则：**讲义即文档**——图片跟着内容走、穿插在正文相关段落后；课件直接从讲义结构生成，不需要在讲义文末维护"PPT 页面规划表"这类制作过程遗留物。

```bash
# 单文件
python <skill_dir>/scripts/handout_to_pptx.py --input 第1课讲义.md --out 第01课-课名.pptx \
  --label "小学 · 40 分钟" --author "课程名"

# 批量（根目录下每个含 第*课讲义.md 的数字子目录算一课，输出镜像目录结构）
python <skill_dir>/scripts/handout_to_pptx.py --input-dir <讲义根目录> --out-dir <输出根目录> \
  --label "小学 · 40 分钟"
```

转换规则：H1 → 封面页（正文首图作题图）；开头的无序列表 → 学习目标页；每个 H2 节 → 一页内容页（配图就近取节内第一张图，有图时正文占左半区）；"术语卡"节表格 → 术语页；不足 8 页自动补"课堂要点回顾"；输出 `第NN课-课名.pptx` 并写入 core title。

排版纪律（脚本已内置）：行预算截断 `fit_text`（有图 22 字/行 × 15 行、无图 32 字/行 × 12 行，超容量在句末收尾）；教师参考节与引用块默认不上课件（`--` 无参数时用内置 skip 列表，改 `DEFAULT_SKIP` 适配其他项目）；页码框宽度 ≥76px；页面坐标以 960×540 设计，1px = 9525 EMU。

**交付后 QA**：用 `scripts/export_images.py` 的 COM 链路导样张目检（导图前先杀残留 POWERPNT.EXE 进程并用 `DispatchEx` 新实例——残留进程会返回内存旧副本，导出图与磁盘文件不一致）。批量产物每课抽封面 + 1 张内容页即可。

需要逐页精修视觉（自定义版式/图表/动画级）时，仍走上方 PPTD 工作流；本脚本是"快、稳、批量"的讲义转课件专用路径。

## 本地查看器（人工查看与文本微调）

交付后想让用户在浏览器里看效果或微调文字时，启动本地查看器（Node >= 18，零 npm 依赖，完全离线）：

```bash
node <skill_dir>/ui/server.mjs --project /abs/path/<项目名> --port 55280
```

让用户打开 `http://127.0.0.1:55280/`。功能：

- **预览**：总览网格 + 单页大图 + 键盘左右翻页。数据源是项目 `.qa-images/pages/*.png`（COM 真实渲染，所见即所得）；没有 `.qa-images` 时界面上可一键触发导出。
- **文本微调**：侧栏列出当前页全部文本元素（elementId + 内容预览），编辑、保存即回写 `.page`（保留元素其他字段与富文本标签）；改完点"重新导出预览"刷新画面。

仅想让用户快速看效果时，不必启动查看器——直接交付 `scripts/export_images.py` 产出的 `overview.jpg` / `pages/*.png` 即可。布局/样式级修改（位置、颜色、字号）不是查看器职责，由 Agent 按用户要求直接改 `.page` 后重新导出。

## 环境依赖

```bash
pip install -r <skill_dir>/scripts/requirements.txt
```

- 核心：python-pptx>=1.0、PyYAML；公式链：latex2mathml、lxml（XSL 自动定位，可用环境变量 `PPT_STUDIO_MML2OMML` 覆盖）；图表链：plotly、kaleido、matplotlib；QA 拼图：Pillow
- FA 图标：环境变量 `PPT_STUDIO_FA_SVGS` 指向 FA svgs 目录；默认向上搜索 `assets/fa/node_modules/@fortawesome/fontawesome-free/svgs`
- 图片 QA 需本机 Microsoft PowerPoint（COM）；kaleido 出图需本机 Chrome/Chromium
- 本地查看器：Node.js >= 18，零 npm 依赖
- 命令平台差异：Windows 用 `python`（类 Unix 可用 `python3`）
- **依赖自恢复**：脚本运行报 `ModuleNotFoundError` / `ImportError` 时，先执行上面的 `pip install -r` 再重试，不逐个 pip install 猜包名。
