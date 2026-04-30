---
name: obsidian-to-docx
description: >
  将 Obsidian Markdown 笔记转换为符合 GB/T 9704-2012 格式的 Word 文档（.docx），
  下载并重命名后移动到合适目录。
  当用户说以下任意内容时触发：
  - "把这篇笔记转成 word / docx"
  - "导出 obsidian 笔记为文档"
  - "生成公文格式的文档"
  - "把 md 转成 docx"
  - "下载转换后的文档"
  - 任何涉及将 Markdown/Obsidian 笔记转为 Word 文档的请求。
  服务端点默认运行在 http://localhost:7002，可通过 .env 中的 BASE_URL 修改。
---

# Obsidian → DOCX

## 完整工作流

```
用户提供 .md 文件路径
    │
    ▼
上传文件到 /api/markdown-to-docx
    │
    ▼
获取 download_url
    │
    ▼
下载 docx 到临时位置
    │
    ▼
重命名为合适名称（默认：原文件名去掉 .md 加 .docx）
    │
    ▼
移动到目标目录（默认：与原 .md 同级目录）
    │
    ▼
向用户报告最终路径
```

## 服务端点

```
POST http://<host>:7002/api/markdown-to-docx
Content-Type: multipart/form-data
```

> `<host>` 默认为 `localhost`，如果服务部署在远程服务器上则替换为对应 IP。

## 完整命令

### 1. 上传并获取下载链接

```bash
curl -s -X POST http://localhost:7002/api/markdown-to-docx \
  -F "file=@/path/to/note.md"
```

返回：
```json
{
  "download_url": "http://localhost:7002/api/download/xxx.docx",
  "filename": "xxx.docx",
  "path": "output/xxx.docx"
}
```

### 2. 下载并保存到指定位置

```bash
# 默认保存为原文件名.docx，放在与原 .md 同级目录
curl -s -X POST http://localhost:7002/api/markdown-to-docx \
  -F "file=@/path/to/note.md" | \
  python3 -c "
import json, sys, subprocess, os
result = json.load(sys.stdin)
url = result['download_url']
md_path = '/path/to/note.md'
dst_dir = os.path.dirname(md_path)
dst_name = os.path.splitext(os.path.basename(md_path))[0] + '.docx'
dst_path = os.path.join(dst_dir, dst_name)
subprocess.run(['curl', '-s', '-o', dst_path, url], check=True)
print(dst_path)
"
```

## 特性

- 自动过滤 Obsidian YAML frontmatter（`---` 包裹的属性块）
- 自动识别标题层级（`#` → Heading 1, `##` → Heading 2 ... `#######` → Heading 7）
- 自动识别 Markdown 表格
- 无法识别的内容回退为正文
- 文件名（不含 `.md`）作为文档标题
- 副标题固定为 "副标题"
- 生成 GB/T 9704-2012 格式文档（方正小标宋/黑体/楷体/仿宋、奇偶页不同页码等）

## 其他可用接口

```bash
# JSON 直接生成 docx
POST http://localhost:7002/api/create-docx
Content-Type: application/json

# 健康检查
GET http://localhost:7002/health
```

## 故障排查

若服务未响应：
```bash
# systemd 方式运行
systemctl --user status docx
systemctl --user restart docx

# 或直接启动
cd <项目目录>
source venv/bin/activate
python main.py
```
