"""FastAPI 入口."""

import os
from pathlib import Path
from uuid import uuid4

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from src.markdown_parser import parse_markdown
from src.generator import DocumentGenerator
from src.models import DocumentRequest, DocumentResponse

# 加载 .env
load_dotenv()

app = FastAPI(title="Docx Service", version="0.1.0")

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

BASE_URL = os.getenv("BASE_URL", "")


def _build_response(filename: str, output_path: Path) -> DocumentResponse:
    """构建标准响应."""
    download_url = f"/api/download/{filename}"
    if BASE_URL:
        download_url = f"{BASE_URL.rstrip('/')}{download_url}"

    return DocumentResponse(
        download_url=download_url,
        filename=filename,
        path=str(output_path),
    )


@app.post("/api/create-docx", response_model=DocumentResponse)
async def create_docx(request: DocumentRequest):
    """接收 JSON 请求，生成 docx 文件."""
    filename = f"{uuid4().hex}.docx"
    output_path = OUTPUT_DIR / filename

    generator = DocumentGenerator()
    generator.generate(request, str(output_path))

    return _build_response(filename, output_path)


@app.post("/api/markdown-to-docx", response_model=DocumentResponse)
async def markdown_to_docx(file: UploadFile = File(...)):
    """上传 Markdown 文件，自动转换为 docx.

    处理流程：Markdown -> JSON -> docx.
    无法识别的 Markdown 内容会回退为正文.
    """
    # 1. 读取 Markdown 内容
    content = await file.read()
    text = content.decode("utf-8")

    # 2. 提取文件名（不含扩展名）作为标题
    title = Path(file.filename).stem if file.filename else "未命名文档"

    # 3. Markdown -> JSON
    doc_request = parse_markdown(text, title=title)

    # 4. JSON -> docx
    filename = f"{uuid4().hex}.docx"
    output_path = OUTPUT_DIR / filename

    generator = DocumentGenerator()
    generator.generate(DocumentRequest(**doc_request), str(output_path))

    return _build_response(filename, output_path)


@app.get("/api/download/{filename}")
async def download_docx(filename: str):
    """下载生成的 docx 文件."""
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.get("/health")
async def health():
    return {"status": "ok"}

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=7002)