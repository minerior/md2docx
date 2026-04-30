"""Markdown -> JSON 转换器.

解析 Obsidian 风格的 Markdown 文件，输出符合 DocumentRequest 格式的 JSON.

用法:
    python src/markdown_parser.py input.md [--output output.json]
"""

import argparse
import json
import re
from pathlib import Path


def parse_markdown(text: str, title: str = "") -> dict:
    """解析 Markdown 文本，返回 DocumentRequest 格式的字典.

    Args:
        text: Markdown 文本内容
        title: 文档标题（如果为空，使用默认值）

    保底机制：任何无法识别的内容都会回退为 body_text.
    """
    lines = text.splitlines()

    # 1. 过滤 YAML frontmatter (Obsidian 笔记属性)
    lines = _strip_frontmatter(lines)

    # 2. 解析内容
    content = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        # 表格
        if stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            table_item = _parse_table(table_lines)
            if table_item:
                content.append(table_item)
            else:
                # 表格解析失败，回退为正文
                content.append({
                    "type": "body_text",
                    "text": "\n".join(table_lines),
                })
            continue

        # 标题
        heading_match = re.match(r"^(#{1,7})\s+(.+)", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            heading_text = heading_match.group(2).strip()
            content.append({
                "type": f"heading_{level}",
                "text": heading_text,
            })
            i += 1
            continue

        # 正文段落（合并连续的非空行）
        paragraph_lines = [stripped]
        i += 1
        while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("|") and not re.match(r"^#{1,7}\s", lines[i].strip()):
            paragraph_lines.append(lines[i].strip())
            i += 1
        content.append({
            "type": "body_text",
            "text": " ".join(paragraph_lines),
        })

    # 3. 构建 DocumentRequest
    return {
        "title": title or "未命名文档",
        "subtitle": "副标题",
        "content": content,
    }


def parse_markdown_file(md_path: str) -> dict:
    """解析 Markdown 文件，返回 DocumentRequest 格式的字典.

    文件名（不含扩展名）作为文档标题.
    """
    md_path = Path(md_path)
    text = md_path.read_text(encoding="utf-8")
    return parse_markdown(text, title=md_path.stem)


def _strip_frontmatter(lines: list[str]) -> list[str]:
    """过滤 YAML frontmatter（Obsidian 笔记属性）.

    如果文档以 '---' 开头，则找到下一个 '---' 并删除两者之间的内容.
    """
    if not lines:
        return lines

    if lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return lines[i + 1:]
        # 没有找到结束标记，返回全部（异常情况）
        return lines

    return lines


def _parse_table(table_lines: list[str]) -> dict | None:
    """解析 Markdown 表格.

    Obsidian 表格格式:
        | 字段名称 | 说明 | 示例 |
        | ------ | ---- | ---- |
        | 数据1  | 数据2 | 数据3 |
    """
    if len(table_lines) < 3:
        return None

    # 第一行：表头
    header = [cell.strip() for cell in table_lines[0].split("|")[1:-1]]
    if not header:
        return None

    # 第二行：分隔线，跳过
    # 第三行及以后：数据
    rows = []
    for line in table_lines[2:]:
        cells = [cell.strip() for cell in line.split("|")[1:-1]]
        if cells and any(cells):
            rows.append(cells)

    if not rows:
        return None

    return {
        "type": "table",
        "header": header,
        "rows": rows,
    }


def main():
    parser = argparse.ArgumentParser(description="Markdown to JSON converter")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("--output", "-o", help="输出 JSON 文件路径（默认输出到 stdout）")
    args = parser.parse_args()

    result = parse_markdown_file(args.input)
    json_text = json.dumps(result, ensure_ascii=False, indent=2)

    if args.output:
        Path(args.output).write_text(json_text, encoding="utf-8")
        print(f"✅ 已保存到 {args.output}")
    else:
        print(json_text)


if __name__ == "__main__":
    main()
