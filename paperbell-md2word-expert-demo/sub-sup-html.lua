-- sub-sup-html.lua
-- -----------------------------------------------------------------------------
-- 作用：把正文里 raw HTML 写法的 <sub>...</sub> / <sup>...</sup>
--       转换成 Pandoc 原生 AST 的 Subscript / Superscript 节点。
--
-- 为什么需要它：
--   Pandoc 3.x 解析 Markdown 时，会把 <sub>2</sub> 拆成三个独立节点：
--     [ RawInline(html) "<sub>", Str "2", RawInline(html) "</sub>" ]
--   而导出 .docx 时，docx 写出器不支持 raw HTML，两个 RawInline 被直接丢弃，
--   只剩 Str "2"，导致化学式 CO2、Ca2+ 等的角标格式全部丢失（变成普通文本）。
--   本过滤器遍历 inline 序列，把开/闭标签配对，内层内容包成原生 Subscript/Superscript，
--   这样 docx 写出器就能正确生成为 <w:vertAlign w:val="subscript|superscript"/>。
--
-- 适配 Obsidian：Obsidian 实时预览原生渲染 <sub>/<sup>，Pandoc 原生 ~...~ / ^...^ 反而不渲染，
--                所以正文用 <sub>/<sup> + 本过滤器，是「Obsidian 预览 + Word 导出」双赢方案。
--
-- 用法：在 pandoc 命令里 --lua-filter="sub-sup-html.lua"，建议放在 zotero.lua 等过滤器之前。
-- -----------------------------------------------------------------------------

local function transform(inlines)
  local out, i = {}, 1
  while i <= #inlines do
    local el = inlines[i]
    if el.t == "RawInline" and el.format == "html" then
      local opening = el.text:match("^<(%a+)>$")   -- 匹配 <sub> 或 <sup>
      if opening == "sub" or opening == "sup" then
        local closing = "</" .. opening .. ">"
        local j, inner = i + 1, {}
        -- 向后查找匹配的闭合标签
        while j <= #inlines and not (inlines[j].t == "RawInline"
              and inlines[j].format == "html" and inlines[j].text == closing) do
          inner[#inner + 1] = inlines[j]
          j = j + 1
        end
        if j <= #inlines then
          -- 找到闭合标签，包成 Subscript / Superscript
          out[#out + 1] = (opening == "sub")
            and pandoc.Subscript(inner)
            or  pandoc.Superscript(inner)
          i = j + 1
        else
          -- 未找到闭合标签，原样保留（避免误吞后续内容）
          out[#out + 1] = el
          i = i + 1
        end
      else
        out[#out + 1] = el
        i = i + 1
      end
    else
      out[#out + 1] = el
      i = i + 1
    end
  end
  return out
end

function Inlines(inlines)
  return transform(inlines)
end
