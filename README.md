# Docx Service

基于 FastAPI + python-docx 的公文文档生成服务，严格遵循 **GB/T 9704-2012《党政机关公文格式》**。

核心设计理念：**样式由模板控制，代码只负责内容编排**。

项目的目标是基于.md文件生成一个格式严谨的文档，但这个文档不应是最终版本，用户可以根据此文档进行后续的格式微调。因此项目中还包括了Normal.dotm

---

## 架构

```
Agent / 用户
    │
    ├─ HTTP POST /api/create-docx          (JSON → docx)
    ├─ HTTP POST /api/markdown-to-docx     (Markdown → docx)
    │
    ▼
FastAPI (main.py)
    ├─ DocumentGenerator (src/generator.py)
    │   ├─ 加载 Normal.docx（模板样式）
    │   ├─ 第一节：封面 + 目录（无页码）
    │   ├─ 第二节：正文（页码从 1 开始，奇右偶左）
    │   └─ 输出 docx
    │
    └─ md_to_json.py（Markdown 解析器）
        └─ 过滤 Obsidian 属性 / 提取标题、正文、表格

Normal.docx（Windows 维护，服务器直接使用）
    ├─ Heading 1    → 黑体 三号
    ├─ Heading 2    → 楷体 三号
    ├─ Heading 3~7  → 仿宋 三号
    ├─ Normal       → 仿宋 三号（首行缩进两字符）
    ├─ Title        → 方正小标宋简体 二号（居中）
    ├─ 目录标题      → 居中
    ├─ 表头         → 黑体（无缩进）
    └─ No Spacing   → 仿宋（无缩进，用于表格内容）
```

---

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 准备模板
# 将包含 GB/T 9704-2012 样式的 Word 模板命名为 Normal.docx 放在项目根目录
# 或使用仓库中的默认模板

# 3. 启动服务
python main.py
# 或
uvicorn main:app --host 0.0.0.0 --port 7002

# 4. 测试
curl http://localhost:7002/health
```

---

## API 端点

### 1. JSON 生成文档

```http
POST /api/create-docx
Content-Type: application/json
```

**请求体**：
```json
{
  "title": "关于XXX的请示",
  "subtitle": "XX市人民政府",
  "content": [
    {"type": "heading_1", "text": "一、工作背景"},
    {"type": "body_text", "text": "根据ABC-123号文件要求..."},
    {"type": "table", "header": ["项","内容"], "rows": [["1","数据"]]}
  ]
}
```

**支持的 `type`**：`title`, `subtitle`, `heading_1~7`, `body_text`, `page_break`, `table`

---

### 2. Markdown 生成文档

```http
POST /api/markdown-to-docx
Content-Type: multipart/form-data
```

**参数**：`file`（上传 `.md` 文件）

**特性**：
- 自动过滤 Obsidian YAML frontmatter（`---` 包裹的属性块）
- 自动识别标题 `#`、表格 `|`、正文
- 文件名作为文档标题
- 无法识别的内容自动回退为正文

```bash
curl -X POST http://localhost:7002/api/markdown-to-docx \
  -F "file=@document.md"
```

---

### 3. 下载文件

```http
GET /api/download/{filename}
```

---

### 4. 健康检查

```http
GET /health
```

---

## 配置

| 文件 | 说明 |
|---|---|
| `.env` | `BASE_URL`：控制返回的下载链接前缀 |
| `gongwen.yaml` | 封面空行数、目录开关、样式映射等 |
| `Normal.docx` | **Word 模板**，所有字体/字号/页边距由它控制 |

---

## 项目结构

```
.
├── main.py                          # FastAPI 入口
├── gongwen.yaml                     # 服务配置
├── Normal.docx                      # Word 样式模板
├── .env.example                     # 环境变量模板
├── src/
│   ├── __init__.py
│   ├── generator.py                 # 文档生成器核心
│   ├── config.py                    # 配置加载
│   ├── models.py                    # Pydantic 模型
│   └── markdown_parser.py           # Markdown → JSON 转换器
├── tests/
│   ├── __init__.py
│   └── test_generator.py            # 单元测试
└── output/                          # 生成文件留存目录（.gitignore）
```

---

## 样式规范（GB/T 9704-2012）

| 元素 | 字体 | 字号 |
|---|---|---|
| 封面标题 | 方正小标宋简体 | 二号（22pt） |
| 一级标题 | 黑体 | 三号（16pt） |
| 二级标题 | 楷体 | 三号（16pt） |
| 三~七级标题 | 仿宋 | 三号（16pt） |
| 正文 | 仿宋 | 三号（16pt），首行缩进两字符 |
| 表格表头 | 黑体 | 三号 |
| 表格内容 | 仿宋 | 三号 |
| 页码 | Times New Roman | 四号（14pt），奇右偶左，格式 `— N —` |

---

## 字体依赖

本项目生成的文档使用以下字体，以符合 **GB/T 9704-2012** 规范：

| 用途 | 字体 | 说明 |
|---|---|---|
| 封面标题 | 方正小标宋简体 | 商业字体，需自行获取授权 |
| 一级标题 | 黑体 | 系统自带 |
| 二级标题 | 楷体（方正楷体_GBK） | 商业字体，可用系统楷体替代 |
| 三~七级标题 / 正文 | 仿宋（仿宋_GB2312） | 商业字体，可用系统仿宋替代 |
| 页码 | Times New Roman | 系统自带 |
| 表格表头 | 黑体 | 系统自带 |

> ⚠️ **合规提示**：上述商业字体文件（`.ttf`/`.otf`）**受版权保护，请勿放入开源仓库**。`Normal.docx` 模板仅引用字体名称，本身不包含字体文件，可安全分发。
>
> 若目标机器未安装对应字体，Word 会自动回退到系统默认字体，文档仍可正常打开和编辑，仅显示效果略有差异。

### Linux 服务器安装示例

```bash
# 将字体文件复制到用户字体目录
mkdir -p ~/.local/share/fonts
cp 方正小标宋简体.ttf ~/.local/share/fonts/
cp 仿宋_GB2312.ttf ~/.local/share/fonts/
cp 方正楷体_GBK.ttf ~/.local/share/fonts/
fc-cache -fv
```

---

## 维护模板

仓库中建议同时保留两个文件：

```
├── Normal.dotm     ← 可编辑源文件（Windows Word 打开）
└── Normal.docx     ← 服务运行时的加载模板
```

### 修改样式流程

1. 在 Windows Word 中打开 `Normal.dotm`
2. 通过「开始 → 样式」修改所需样式（字体、字号、段落格式等）
3. **文件 → 另存为 → 选择 `Normal.docx`**，覆盖项目根目录下的同名文件
4. 重启服务（或热重载）即可生效

> `Normal.dotm` 作为可编辑源文件放入仓库，方便协作者修改样式后导出 `.docx`。

---

## 许可

MIT License
