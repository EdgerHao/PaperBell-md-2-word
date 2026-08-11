---
tags:
  - Pandoc
  - 工具指南
created: 2026-07-06
---

# meta2style.lua 使用说明

> 把 Markdown frontmatter 的作者、单位、摘要、关键词，自动映射为 Word 自定义样式段落。

## 一、这是什么

`meta2style.lua` 是一个 Pandoc Lua 过滤器，用于把 Markdown frontmatter 中的标题块字段（`title` / `author` / `affiliation` / `abstract` / `keywords`）映射为带 `custom-style` 属性的 Word 段落，使这些内容在导出 `.docx` 时与正文（`Body Text` / `First Paragraph`）样式区分开。

与 LaTeX 路径思路一致（都是模板变量）：写作时只需在 frontmatter 填字段，正文干净，导出时自动套样式。

## 二、智能触发（不误伤其他文档）

仅当 frontmatter 含 `abstract` / `keywords` / `affiliation` **任一字段**时才接管标题块：

| 文档类型 | frontmatter 特征 | 行为 |
|---|---|---|
| 期刊手稿 | 有 `abstract`/`keywords` | ✓ 接管，输出独立样式段落 |
| 博士论文 | 只有 `title`/`author`，摘要写在正文 | ✗ 跳过，保持 Pandoc 默认 |
| 普通笔记 | 无这些字段 | ✗ 跳过 |

## 三、frontmatter 字段

### 3.1 字段→样式映射

| 字段 | 样式名 | 说明 |
|---|---|---|
| `title` | `Title` | 1 段 |
| `author` | `Author` | 合并为 1 行，多作者用 `, ` 分隔 |
| `affiliation` | `Affiliation` | 合并为 1 段，多单位用换行连接 |
| `abstract` | `Abstract` | 多段，自动加内联前缀 **Abstract: **（加粗） |
| `keywords` | `Keywords` | 自动加内联前缀 **Keywords: **（加粗），`; ` 分隔 |
| `abstractTitle` | （控制字段） | 控制摘要前缀，见 3.3 |

### 3.2 block scalar 写法（关键）

`author` / `affiliation` / `abstract` / `keywords` **必须用 `|` block scalar**，才能让 Pandoc 解析其中的 `<sup>`/`<sub>` 富文本。普通双引号字符串不解析 Markdown，角标会丢失。

```yaml
---
title: "Allicin-functionalized Fe<sub>3</sub>O<sub>4</sub> magnetic colloids"
author: |
  Wanxin Hao<sup>1,2,†</sup>, Yan Wu<sup>1,2,†</sup>, Gang Yan<sup>1,2,*</sup>
affiliation: |
  <sup>1</sup>Hubei Key Laboratory, Yangtze University, Wuhan 430100, China  
  <sup>2</sup>Key Laboratory of Exploration Technologies, Wuhan 430100, China  
  <sup>*</sup>Corresponding author: 202073036@yangtzeu.edu.cn
abstract: |
  Stable emulsified oil droplets are difficult to remove... capacity of 380 mg·g<sup>-1</sup>.
keywords: |
  Allicin; Fe<sub>3</sub>O<sub>4</sub> magnetic colloids; Emulsified oil droplets
---
```

> [!warning] Obsidian Properties 视图陷阱
> 在 Obsidian 的 Properties 视图里编辑含 `<sub>`/`<sup>` 的字段时，Obsidian 会把 `|` block scalar 转成双引号字符串，导致角标丢失。**这类字段请用文本模式编辑**，不要在 Properties 视图里动。

### 3.3 `abstractTitle` 控制字段

| 写法 | 效果 |
|---|---|
| 不写（缺省） | 自动加内联前缀 **Abstract: **（与 keywords 一致） |
| `abstractTitle: false` | 纯摘要，无任何前缀 |
| `abstractTitle: true` | 单独 `Abstract` 标题段（`AbstractTitle` 样式）+ 摘要正文 |
| `abstractTitle: "摘要"` | 单独 `摘要` 标题段 + 摘要正文 |

## 四、Word 模板样式配置（重要，使用者必读）

> [!important]
> Pandoc 对模板里没有的样式会自动创建**空格式**样式（等于没格式）。所以使用者**必须**在自己的 Word 模板中预先创建并设置以下样式的格式，否则导出的作者/单位/摘要/关键词将没有任何排版效果。

### 4.1 需要的样式清单

| 样式名 | 状态 | 用途 | 格式建议（英文期刊） | 格式建议（中文论文） |
|---|---|---|---|---|
| `Title` | 通常已有 | 论文标题 | 居中，Times New Roman，16-18pt，加粗 | 居中，黑体，二号，加粗 |
| `Author` | 通常已有 | 作者行 | 居中，TNR，11pt | 居中，宋体，小四 |
| `Affiliation` | **需新建** | 单位 | 居中，TNR，9pt，斜体 | 居中，宋体，五号 |
| `AbstractTitle` | **需新建** | 摘要标题词（单独段，用 `abstractTitle: true` 时） | 居中，TNR，10pt，加粗 | 居中，黑体，小四，加粗 |
| `Abstract` | 通常已有 | 摘要正文 | 两端对齐，TNR，10pt，左右缩进 | 两端对齐，宋体，小四，左右缩进 |
| `Keywords` | **需新建** | 关键词 | 两端对齐，TNR，10pt | 两端对齐，宋体，小四 |

> [!note] 样式名大小写敏感
> Pandoc 的 `custom-style` 匹配的是模板里样式的 **name**（显示名）。新建样式时，名称必须与上表完全一致（如 `Affiliation`，不是 `affiliation` 或 `Affiliations`）。

### 4.2 在 Word 中新建样式（以 Affiliation 为例）

1. 用 Word 打开你的模板（如 `word模板-常规.docx`）
2. 「开始」→ 样式面板 → 点击右下角箭头展开全部样式
3. 点击左下角「新建样式」按钮（带 A+ 图标）
4. 在弹出对话框中设置：
   - **名称**：`Affiliation`（精确匹配上表）
   - **样式类型**：段落
   - **样式基准**：无样式（避免继承链导致意外格式）
   - **后续段落样式**：正文 / Body Text
5. 点击左下角「格式」→「字体」：设为 Times New Roman，9pt，斜体
6. 点击「格式」→「段落」：对齐方式选「居中」
7. 确定 → 再确定
8. `Ctrl+S` / `Cmd+S` 保存模板

### 4.3 修改已有样式（以 Author 为例）

1. 样式面板找到 `Author` 样式
2. 右键 →「修改」
3. 调整字号、字体、对齐等格式
4. 确定 → 保存模板

> [!tip] 已有样式也要检查
> `Title` / `Author` / `Abstract` 虽然模板里通常已有，但格式可能不符合当前期刊要求。建议打开模板逐一检查，不符合就按 §4.3 修改。

### 4.4 验证样式生效

导出 Word 后打开检查：

- [ ] 鼠标点在标题段，样式面板高亮 `Title`
- [ ] 点在作者段，高亮 `Author`
- [ ] 点在单位段，高亮 `Affiliation`
- [ ] 点在摘要段，高亮 `Abstract`
- [ ] 摘要/关键词段开头的 **Abstract:** / **Keywords:** 加粗
- [ ] 上下标（Fe₃O₄、mg·g⁻¹）正确显示

## 五、安装与挂载

### 5.1 命令行

在 pandoc 命令中加 `--lua-filter`，放在 `sub-sup-html.lua` 之后：

```bash
pandoc input.md -s -o output.docx -t docx \
  --filter pandoc-crossref \
  --lua-filter sub-sup-html.lua \
  --lua-filter meta2style.lua \
  --lua-filter zotero.lua \
  --reference-doc word模板-常规.docx
```

### 5.2 obsidian-enhancing-export 插件

在插件 `data.json` 的 `Word (.docx)` 配置 `customArguments` 里追加：

```
--lua-filter="${pluginDir}/../../../0-辅助/IOTO/Pandoc/meta2style.lua"
```

放在 `replace-nbsp.lua` 之后即可。改完后重启 Obsidian 让插件重载配置。

## 六、故障排查

### Q: 导出后作者/单位/摘要没样式（显示为正文格式）
**A**：Word 模板里没有对应样式。按 §4.2 在模板里新建并设置格式。Pandoc 自动创建的是空格式样式。

### Q: AbstractTitle 没显示
**A**：缺省行为是加内联前缀 `Abstract: `（不显示单独标题段）。如要单独标题段，设 `abstractTitle: true`。

### Q: 摘要/关键词的 "Abstract:" / "Keywords:" 没加粗
**A**：确认用的是最新 `meta2style.lua`（前缀用 `pandoc.Strong` 包裹）。

### Q: 上下标丢失（Fe3O4 变成 Fe3O4，没有角标）
**A**：frontmatter 字段用了双引号字符串而非 `|` block scalar。含 `<sub>`/`<sup>` 的字段必须用 `|` block scalar，且命令链要挂 `sub-sup-html.lua`。常见原因是 Obsidian Properties 视图把 block scalar 转成了字符串。

### Q: 博士论文导出后标题块消失
**A**：智能触发机制下，博士论文（frontmatter 只有 `title`/`author`，无 `abstract`）会被跳过，保持默认。若标题块消失，检查 frontmatter 是否误加了 `abstract` 字段。

### Q: 作者/单位重复显示两遍
**A**：frontmatter 写了一份，正文里还保留着作者/单位内容。删掉正文的重复内容，只留 frontmatter 那份。

## 七、技术备忘（实现踩坑记录）

1. **必须用 `Div` 包裹 `Para`**：`pandoc.Para(content, attr)` 的 attr 参数会被忽略，必须用 `pandoc.Div({pandoc.Para(...)}, attr)`。
2. **必须用 `pandoc.List` 构造 `doc.blocks`**：用普通 Lua table 赋值会导致插入的段落被 docx writer 悄悄丢弃。
3. **必须清空 `doc.meta.title/author/date/abstract`**：否则 docx writer 会再渲染一遍默认标题块，导致重复。
4. **Pandoc 3.9 Meta API 变了**：不再用 `MetaInlines`/`MetaList` 等 wrapper 类型（`m.t` 是 `nil`），必须用 `pandoc.utils.type()` 判断类型。网上旧教程多基于 Pandoc 2.x，会踩此坑。

适配：Pandoc 3.9 + Lua 5.4
