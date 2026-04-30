"""Pydantic 请求/响应模型."""

from typing import Literal

from pydantic import BaseModel, Field


class ContentItem(BaseModel):
    """文档内容元素."""

    type: Literal[
        "title",
        "subtitle",
        "heading_1",
        "heading_2",
        "heading_3",
        "heading_4",
        "heading_5",
        "heading_6",
        "heading_7",
        "body_text",
        "page_break",
        "table",
    ]
    text: str = ""
    # table 专用字段
    header: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)


class DocumentRequest(BaseModel):
    """生成文档请求."""

    title: str
    subtitle: str = ""
    date: str = ""  # 空则使用今天（中文格式）
    content: list[ContentItem] = Field(default_factory=list)


class DocumentResponse(BaseModel):
    """生成文档响应."""

    download_url: str
    filename: str
    path: str
