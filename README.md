# MD2Doc

将 Markdown 文件转换为符合**公文格式规范**的 Word (.docx) 文件。

## 功能

- 基于 pandoc + python-docx 的 Markdown → Word 转换
- 内置中国公文格式规范（黑体标题、仿宋正文、固定行距 28 磅等）
- 支持通过格式要求文件自定义样式
- 支持 HTML 表格转换（含 `colspan`/`rowspan` 合并单元格）
- 自动处理东亚字体、图片行距、表格自动适配
- **MCP Server**：可在 Claude Code 中直接调用转换和分析工具
- **语义分析**：检查文档结构和内容，生成修改建议报告

## 安装

### 环境要求

- Python >= 3.12
- [pandoc](https://pandoc.org/installing.html)（需在 `~/.local/bin/pandoc` 或 PATH 中）

### 安装依赖

```bash
git clone <repo-url>
cd md2doc
uv sync
```

## CLI 命令行使用

### 基本用法

```bash
# 使用默认公文格式转换
uv run python scripts/md2docx.py <input.md>

# 指定输出路径
uv run python scripts/md2docx.py <input.md> -o output.docx

# 跳过格式设置（纯 pandoc 输出）
uv run python scripts/md2docx.py <input.md> --no-format
```

### 自定义格式

```bash
# 使用自定义格式要求文件
uv run python scripts/md2docx.py <input.md> -f 格式要求.md

# 叠加参考模板（模板样式覆盖格式要求）
uv run python scripts/md2docx.py <input.md> -f 格式要求.md -r template.docx
```

### 格式要求文件格式

参照项目根目录下的 `格式要求.md`：

```markdown
## H1（一级标题）
- 字体：黑体，三号，加粗
- 对齐方式：左对齐
- 行距：固定值28磅
- 间距：段前0行，段后0行
- 大纲级别：1级
- 特殊格式：无

## H2（二级标题）
- 字体：仿宋_GB2312，小三，加粗
- 对齐方式：左对齐
- 行距：固定值28磅
- ...

## 正文
- 字体：仿宋_GB2312，四号
- 对齐方式：两端对齐
- 行距：固定值28磅
- 特殊格式：首行缩进2字符

## 表格
- 字体：仿宋_GB2312，五号
- 对齐方式：居中对齐
- 行距：单倍行距
```

支持的字号：三号(16pt)、小三(15pt)、四号(14pt)、小四(12pt)、五号(10.5pt)、小五(9pt)

---

## MCP Server

md2doc 提供 MCP（Model Context Protocol）服务器，可在 Claude Code 中直接使用。

### 配置 MCP 服务器

在 `~/.claude.json` 的 `mcpServers` 中添加：

```json
{
  "mcpServers": {
    "md2doc": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/path/to/md2doc",
        "md2doc-mcp"
      ]
    }
  }
}
```

将 `/path/to/md2doc` 替换为实际的项目路径。配置后**重启 Claude Code** 即可生效。使用 `/mcp` 命令查看是否加载成功。

### 可用工具

#### 1. convert_markdown_to_docx — 转换 Markdown 为 Word

将 Markdown 文件转换为符合公文格式规范的 .docx 文件。

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `input_path` | string | ✓ | 输入的 Markdown 文件绝对路径 |
| `output_path` | string | | 输出 .docx 路径，默认为 `<输入目录>/<文件名>.docx` |
| `format_spec` | string | | 格式要求 Markdown 文件路径（如 格式要求.md） |
| `reference_docx` | string | | 参考 .docx 模板，样式会叠加在格式要求之上 |
| `extract_media` | string | | 提取媒体文件的目标目录 |
| `skip_format` | bool | | 设为 true 跳过所有格式设置（纯 pandoc 输出） |
| `no_cache_template` | bool | | 设为 true 强制重新生成缓存的参考模板 |

**格式设置的优先级**：未指定 `format_spec` 时，自动使用项目根目录下的 `格式要求.md`（可通过修改 `mcp_server.py` 中的 `DEFAULT_FORMAT_SPEC_PATH` 配置）；如果该文件也不存在，则使用内置的默认公文格式。

**使用示例对话**：

> 用户：帮我把 report.md 转成 Word 文档
>
> Claude：好的。你有特定格式要求吗？可以提供格式要求的 markdown 文件，或者参考的 docx 模板。
>
> 用户：没有特殊要求，用默认格式就行
>
> Claude：*调用 convert_markdown_to_docx，使用默认格式转换*

> 用户：用 custom-format.md 的格式转换
>
> Claude：*调用 convert_markdown_to_docx，format_spec="custom-format.md"*

#### 2. analyze_and_suggest — 文档语义分析

分析 Markdown 文档的结构和内容，在同目录下生成 `xx-修改建议.md` 修改建议报告。

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `input_path` | string | ✓ | 输入的 Markdown 文件绝对路径 |

**分析类别**（共 6 类）：

| 类别 | 检查内容 |
|------|----------|
| 标题层级问题 | 层级跳跃（H1→H3）、缺少主标题、多个 H1、标题过长 |
| 必要章节检查 | 公文常见要素：标题模式、发文机关、日期、附件/附录 |
| 格式问题 | 失效的图片引用、HTML 表格标签不匹配 |
| 内容问题 | 空章节、占位符文本、重复段落、段落过长 |
| 编号连续性 | 编号间隙、重新编号、格式不一致 |
| 中文标点规范 | 半角标点、英文引号、缺少句末标点 |

**输出文件格式**：

```markdown
# 修改建议 — report.md

> 生成时间: 2026-06-28 14:30:00
> 分析文件: /path/to/report.md
> 总行数: 1523
> 发现问题: 12 项（错误 2 项，警告 7 项，提示 3 项）

---

## 一、标题层级问题（2项）

| 行号 | 严重度 | 问题描述 |
|------|--------|----------|
| 42   | 警告   | 标题层级跳跃: H1 → H3，缺少 H2 |
| --   | 错误   | 缺少一级标题 (H1) |

## 二、必要章节检查（1项）

| 行号 | 严重度 | 问题描述 |
|------|--------|----------|
| --   | 警告   | 未检测到「日期信息」，建议检查是否遗漏 |
...
```

**使用示例对话**：

> 用户：帮我检查一下这篇文档有什么问题
>
> Claude：*调用 analyze_and_suggest，分析后返回修改建议文件路径和问题摘要*

---

## MCP 服务器配置项

默认格式要求文件的路径在 `src/md2doc/mcp_server.py` 中配置：

```python
# 修改此行即可更改默认格式
DEFAULT_FORMAT_SPEC_PATH: str | None = str(
    Path(__file__).parent.parent.parent / "格式要求.md"
)
```

- 设为其他路径：使用该路径的格式要求文件
- 设为 `None`：跳过加载默认格式文件，直接使用内置的硬编码默认格式

---

## 项目结构

```
md2doc/
├── scripts/
│   └── md2docx.py          # CLI 转换脚本（独立运行）
├── src/md2doc/
│   ├── __init__.py          # 包元数据
│   ├── converter.py         # 核心转换包装器（异常驱动的 API）
│   ├── analyzer.py          # 语义分析引擎
│   └── mcp_server.py        # MCP 服务器入口
├── 格式要求.md               # 默认公文格式规范
├── 测试用例/                 # 测试文档
│   ├── 测试文档.md
│   ├── 模板.docx
│   └── images/
└── pyproject.toml
```

## 默认公文格式规范

| 样式 | 字体 | 字号 | 加粗 | 对齐 | 行距 | 缩进 |
|------|------|------|------|------|------|------|
| H1（一级标题） | 黑体 | 三号 16pt | ✓ | 左对齐 | 固定 28pt | 无 |
| H2（二级标题） | 仿宋_GB2312 | 小三 15pt | ✓ | 左对齐 | 固定 28pt | 无 |
| H3（三级标题） | 仿宋_GB2312 | 四号 14pt | ✓ | 左对齐 | 固定 28pt | 无 |
| 正文 | 仿宋_GB2312 | 四号 14pt | | 两端对齐 | 固定 28pt | 首行缩进 2 字符 |
| 表格 | 仿宋_GB2312 | 五号 10.5pt | | 居中 | 单倍行距 | 无 |

## 许可证

MIT
