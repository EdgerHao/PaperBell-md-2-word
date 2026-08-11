-- meta2style.lua
-- -----------------------------------------------------------------------------
-- 作用：把 Markdown frontmatter 中的 title / author / affiliation / abstract /
--       keywords 字段，映射为带 custom-style 属性的 Word 段落，使其在导出 .docx
--       时与正文（Body Text / First Paragraph）样式区分开。
--
-- 智能触发（避免误伤博士论文等只用 title 的文档）：
--   仅当 frontmatter 含 abstract / keywords / affiliation 任一字段时才接管标题块。
--   - 期刊手稿（有 abstract/keywords）→ 接管，输出 Title/Author/Affiliation/Abstract/Keywords
--   - 博士论文（只有 title/author，无 abstract）→ 跳过，保持 Pandoc 默认行为
--   - 普通笔记（无这些字段）→ 跳过
--
-- 接管时的字段→样式映射：
--   title       → Title            （1 段）
--   author      → Author           （合并为 1 行，多作者用 ", " 分隔）
--   affiliation → Affiliation      （合并为 1 段，多单位用换行连接）
--   abstract    → AbstractTitle + Abstract（多段）
--   keywords    → Keywords         （自动加 "Keywords: " 前缀，"; " 分隔）
--
-- 富文本支持：
--   author / affiliation / abstract 用 YAML block scalar (|) 书写，其中的
--   <sup>...</sup> / <sub>...</sub> 会被 Pandoc 解析为 AST，配合 sub-sup-html.lua
--   即可在 Word 中保留上下标。普通字符串字段（"xxx"）是纯文本，不解析 Markdown。
--
-- 配套要求：
--   1. Word 模板（word模板-常规.docx）中需预先创建并设置以下样式格式：
--      Title / Author / Affiliation / AbstractTitle / Abstract / Keywords
--      （Pandoc 会自动创建空样式，但格式为默认值，所以必须先在模板里定义好）
--   2. 导出命令中需同时挂载 sub-sup-html.lua（保证 <sup>/<sub> 角标不丢失）
--
-- 适配：Pandoc 3.9 + Lua 5.4（Meta API 已弃用 MetaInlines wrapper，改用 pandoc.utils.type 判断）
-- -----------------------------------------------------------------------------

local utils = require 'pandoc.utils'

-- 判断 x 是否是 Inline 节点
local function is_inline(x)
  return x ~= nil and utils.type(x) == 'Inline'
end

-- 判断 x 是否是 Block 节点
local function is_block(x)
  return x ~= nil and utils.type(x) == 'Block'
end

-- 把任意 meta 值统一转成 inlines（合并为一段）
-- 支持：string / inlines / list-of-inlines / list-of-Para(MetaBlocks)
local function to_inlines(m, sep)
  if m == nil then return nil end
  if type(m) == 'string' then return {pandoc.Str(m)} end
  if type(m) ~= 'table' or #m == 0 then return nil end

  if is_inline(m[1]) then
    -- m 本身就是 inlines
    local out = {}
    for _, i in ipairs(m) do out[#out + 1] = i end
    return out
  elseif is_block(m[1]) then
    -- MetaBlocks：把所有 Para 的 content 合并（段间用 SoftBreak 连接）
    local out = {}
    for idx, b in ipairs(m) do
      if idx > 1 then out[#out + 1] = pandoc.SoftBreak() end
      if b.t == 'Para' then
        for _, i in ipairs(b.content) do out[#out + 1] = i end
      end
    end
    return out
  else
    -- list-of-inlines（如 YAML 列表）：用 sep 合并为一段
    local out = {}
    for idx, sub in ipairs(m) do
      if idx > 1 then out[#out + 1] = pandoc.Str(sep or '; ') end
      if type(sub) == 'table' then
        if is_inline(sub[1]) then
          for _, i in ipairs(sub) do out[#out + 1] = i end
        elseif is_block(sub[1]) then
          for _, b in ipairs(sub) do
            if b.t == 'Para' then
              for _, i in ipairs(b.content) do out[#out + 1] = i end
            end
          end
        end
      end
    end
    return out
  end
end

-- 构造带 custom-style 的 Div（包裹一个 Para）
local function styled_div(inlines, style_name)
  return pandoc.Div(
    {pandoc.Para(inlines)},
    pandoc.Attr('', {}, {['custom-style'] = style_name})
  )
end

function Pandoc(doc)
  local m = doc.meta

  -- 智能触发：无 abstract/keywords/affiliation 时不接管，保持 Pandoc 默认行为
  local has_abstract = m.abstract ~= nil
  local has_keywords = m.keywords ~= nil
  local has_affiliation = m.affiliation ~= nil
  if not (has_abstract or has_keywords or has_affiliation) then
    return doc
  end

  local header = pandoc.List{}

  -- Title
  local title_il = to_inlines(m.title)
  if title_il and #title_il > 0 then
    header:insert(styled_div(title_il, 'Title'))
  end

  -- Author（合并为 1 行，多作者用 ", " 分隔）
  local author_il = to_inlines(m.author, ', ')
  if author_il and #author_il > 0 then
    header:insert(styled_div(author_il, 'Author'))
  end

  -- Affiliation（合并为 1 段，多单位用换行连接）
  local affil_il = to_inlines(m.affiliation)
  if affil_il and #affil_il > 0 then
    header:insert(styled_div(affil_il, 'Affiliation'))
  end

  -- AbstractTitle 控制字段（abstractTitle）：
  --   不写（缺省）        → abstract 段开头加【加粗】内联前缀 "**Abstract:** "
  --   false / "" / "off"  → 不加任何前缀（纯摘要）
  --   true                → 单独 "Abstract" 标题段 + 摘要正文（无内联前缀）
  --   "自定义文字"        → 单独标题段（自定义文字）+ 摘要正文
  local at_mode = 'inline'   -- 默认内联前缀
  local at_text = 'Abstract'
  if m.abstractTitle ~= nil then
    local s = utils.stringify(m.abstractTitle):gsub('^%s+', ''):gsub('%s+$', '')
    local sl = s:lower()
    if s == '' or sl == 'false' or sl == 'off' or sl == 'no' then
      at_mode = 'none'    -- 纯摘要
    elseif sl == 'true' or sl == 'on' or sl == 'yes' then
      at_mode = 'block'   -- 单独标题段
    else
      at_mode = 'block'
      at_text = s         -- 自定义标题段文字
    end
  end

  -- 给 inlines 列表前面加【加粗】的文字前缀
  -- 例如 "Abstract: " → 加粗的 "Abstract:" + 不加粗的空格 + 正文（符合 **Abstract:** text 习惯）
  -- 尾部空白不纳入加粗，避免 Strong 吞掉分隔空格
  local function prefix_inlines(il, prefix)
    local core = prefix:gsub('%s+$', '')            -- 去尾空白："Abstract: " → "Abstract:"
    local tail = prefix:sub(#core + 1)              -- 尾部空白：" "
    local out = {}
    out[#out + 1] = pandoc.Strong{pandoc.Str(core)}  -- 加粗前缀文字
    if tail ~= '' then
      out[#out + 1] = pandoc.Str(tail)               -- 不加粗的空格分隔
    end
    for _, x in ipairs(il) do out[#out + 1] = x end
    return out
  end

  -- Abstract：可选拨标题段 + 摘要正文（每段一个 Abstract 样式）
  if has_abstract then
    if at_mode == 'block' then
      header:insert(styled_div({pandoc.Str(at_text)}, 'AbstractTitle'))
    end
    if type(m.abstract) == 'table' and #m.abstract > 0 and is_block(m.abstract[1]) then
      -- MetaBlocks：每个 Para 一段；第一段视 at_mode 决定是否加内联前缀
      local first = true
      for _, b in ipairs(m.abstract) do
        if b.t == 'Para' then
          local content = b.content
          if first and at_mode == 'inline' then
            content = prefix_inlines(content, 'Abstract: ')
          end
          header:insert(styled_div(content, 'Abstract'))
          first = false
        end
      end
    else
      -- 单段 inlines
      local il = to_inlines(m.abstract)
      if il and #il > 0 then
        if at_mode == 'inline' then
          il = prefix_inlines(il, 'Abstract: ')
        end
        header:insert(styled_div(il, 'Abstract'))
      end
    end
  end

  -- Keywords（自动加【加粗】"Keywords:" 前缀，"; " 分隔）
  if has_keywords and type(m.keywords) == 'table' and #m.keywords > 0 then
    local il = {pandoc.Strong{pandoc.Str('Keywords:')}, pandoc.Str(' ')}
    if is_inline(m.keywords[1]) then
      -- 单个 inlines
      for _, x in ipairs(m.keywords) do il[#il + 1] = x end
    elseif is_block(m.keywords[1]) then
      -- block scalar（|）：MetaBlocks，提取所有 Para 的 content，<sub>/<sup> 会被解析
      for _, b in ipairs(m.keywords) do
        if b.t == 'Para' then
          for _, x in ipairs(b.content) do il[#il + 1] = x end
        end
      end
    else
      -- list（YAML 列表，元素为纯文本字符串，<sub>/<sup> 不解析）
      for idx, k in ipairs(m.keywords) do
        if idx > 1 then il[#il + 1] = pandoc.Str('; ') end
        local ki = to_inlines(k)
        if ki then for _, x in ipairs(ki) do il[#il + 1] = x end end
      end
    end
    if #il > 2 then header:insert(styled_div(il, 'Keywords')) end
  end

  -- 阻止 docx writer 重复渲染默认标题块
  doc.meta.title = nil
  doc.meta.author = nil
  doc.meta.date = nil
  doc.meta.abstract = nil

  -- 拼接：接管段落 + 原正文（必须用 pandoc.List，普通 table 会被丢弃）
  local nb = pandoc.List{}
  for _, b in ipairs(header) do nb:insert(b) end
  for _, b in ipairs(doc.blocks) do nb:insert(b) end
  doc.blocks = nb
  return doc
end
