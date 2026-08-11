---
title: PaperBell × Word：让 Markdown 写作拥有 Word 级排版
author: Shuang Song<sup>1,2</sup>, Wanxin Hao<sup>2,3</sup>, Sen Jiang <sup>3,*</sup>
affiliation: |
  <sup>1</sup>Max Planck Institute of Geoanthropology, Jena, Germany  
  <sup>2</sup>PaperBell Technology, Beijing, China  
  <sup>3</sup>Yangtze University, Wuhan 430100, China  
  <sup>*</sup>Corresponding author: paperbell@songshgeo.com
abstract: PaperBell 一直致力于让学术写作回归内容本身。传统工作流中，Markdown 写作灵活却难以输出符合期刊要求的 Word 文档；而 Word 写作又容易让人陷入格式调整的泥潭。本文介绍一种新的桥接方案：通过 meta2style.lua 过滤器，将 frontmatter 中的作者、单位、摘要、关键词自动映射为 Word 自定义样式段落，让 PaperBell 的 Markdown 写作直接产出排版专业的 Word 文档。你只需专注内容，格式由我们守护。
keywords: |
  PaperBell; Markdown to Word; Pandoc; 学术写作; 样式映射
acronym: PB-WORD
target: PaperBell Word Export
lang: zh
---

> 专注内容，我们帮你格式化。

## 为什么需要 Word 导出

PaperBell 默认通过 LaTeX 路径导出 PDF，这条路产出的文档排版专业、引用精准。但现实中，许多期刊仍要求提交 `.docx` 格式，合作者也常需要 Word 版本进行批注修订。

> PaperBell 的哲学是：**保留 Markdown 写作的灵活性，同时提供 Word 级别的专业输出**。

`meta2style.lua` 正是这条 Word 路径的关键——它让 frontmatter 里集中管理的元数据，自动变成 Word 里带独立样式的段落。

## 它如何工作

你在 frontmatter 里写好作者、单位、摘要、关键词，正文只管内容。导出 Word 时，过滤器自动把这些字段渲染成带 `custom-style` 的段落：

- `title` → `Title` 样式
- `author` → `Author` 样式
- `affiliation` → `Affiliation` 样式
- `abstract` → `Abstract` 样式（自动加 **Abstract:** 前缀）
- `keywords` → `Keywords` 样式（自动加 **Keywords:** 前缀）

> 你只需这样写 frontmatter：
> ```yaml
> author: |
>   Shuang Song<sup>1,2</sup>, Wanxin Hao<sup>2</sup>
> affiliation: |
>   <sup>1</sup>Max Planck Institute of Geoanthropology  
>   <sup>2</sup>PaperBell Technology
> ```
> 导出后，作者行的上标编号、单位的缩进，全部由 Word 样式自动控制，你不用手调一次。

## 智能适配，不越界

过滤器不会打扰不需要它的文档。只有当 frontmatter 含 `abstract` / `keywords` / `affiliation` 之一时，它才接管标题块；否则保持 Pandoc 默认行为。

> 这意味着你的博士论文（frontmatter 只有 `title`/`author`，摘要写在正文）不会被误伤——它依然按原来的方式导出。

## 一次性配置 Word 模板

为了让样式生效，你需要在 Word 模板里预先配置以下样式。这是一次性工作，之后所有文档都受益：

| 样式名 | 用途 | 建议格式 |
|---|---|---|
| `Title` | 论文标题 | 居中，加粗，二号 |
| `Author` | 作者行 | 居中，小四 |
| `Affiliation` | 单位 | 居中，五号，斜体 |
| `Abstract` | 摘要正文 | 两端对齐，小四，左右缩进 |
| `Keywords` | 关键词 | 两端对齐，小四 |

> 新建样式时「样式基准」选「无样式」，避免继承链导致意外格式。详细的操作步骤见 [[Readme-meta2style]]。

## 富文本也能保留

PaperBell 鼓励用 Markdown 管理元数据，而 `meta2style.lua` 配合 `sub-sup-html.lua`，能让 frontmatter 里的富文本正确保留：

- 化学式角标：`Fe<sub>3</sub>O<sub>4</sub>` → Fe₃O₄
- 上标编号：`<sup>1,2,†</sup>` → ¹·²·†
- 单位角标：`mg·g<sup>-1</sup>` → mg·g⁻¹
- 斜体：`*in-situ*` → *in-situ*

> 用文本模式编辑 frontmatter：在 Obsidian 的 Properties 视图里编辑含 `<sub>`/`<sup>` 的字段时，Obsidian 会把 block scalar 转成双引号字符串，导致角标丢失。这类字段请用**文本模式**编辑，保留 `|` block scalar。

## 可选的摘要标题词

缺省情况下，摘要段开头会自动加 **Abstract:** 前缀（与关键词的 **Keywords:** 一致）。如果你想要其他样式，用 `abstractTitle` 字段控制：

```yaml
# 缺省（不写）→ **Abstract:** 摘要内容...
abstractTitle: false      # 纯摘要，无前缀
abstractTitle: true       # 单独 "Abstract" 标题段 + 摘要正文
abstractTitle: "摘要"     # 单独 "摘要" 标题段 + 摘要正文
```

## 导出测试

这篇文档本身就是一份导出测试。用 PaperBell 的 Enhancing Export 插件导出为 Word，你应该看到：

- [ ] 标题段用了 `Title` 样式（居中加粗大字）
- [ ] 作者段用了 `Author` 样式，上标编号 ¹·²、† 保留
- [ ] 单位段用了 `Affiliation` 样式，上标 ¹ ² ³ 保留
- [ ] 摘要段用了 `Abstract` 样式，开头 **Abstract:** 加粗
- [ ] 关键词段用了 `Keywords` 样式，开头 **Keywords:** 加粗
- [ ] 正文标题、段落样式不受影响

## 结语

PaperBell 相信，写作在于把网状的思维变成线性的输出。而格式，不该打断这条线。`meta2style.lua` 把 Markdown 的轻灵和 Word 的严谨缝在一起——你写内容，样式自动就位。

> 以内容为核心，以自动化为手段，以专业输出为目标。

让我们继续探索这套强大的学术写作系统。
