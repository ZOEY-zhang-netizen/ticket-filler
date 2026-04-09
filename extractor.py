"""Unified text extraction for DOCX / PDF / image uploads."""

from __future__ import annotations

import base64
import os
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def _normalize_text(text: str) -> str:
    text = text.replace('\r', '\n')
    text = re.sub(r'\u3000', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r' *\n *', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def extract_text_from_docx(path: str) -> str:
    """Read DOCX while preserving paragraph and table row boundaries."""
    with zipfile.ZipFile(path, 'r') as zf:
        xml = zf.read('word/document.xml')
    root = ET.fromstring(xml)
    lines: list[str] = []
    for elem in root.iter():
        tag = elem.tag
        if tag == f'{W_NS}tr':
            cells = []
            for tc in elem.findall(f'{W_NS}tc'):
                parts = [t.text or '' for t in tc.iter(f'{W_NS}t')]
                cell_text = ''.join(parts).strip()
                if cell_text:
                    cells.append(cell_text)
            if cells:
                lines.append('\t'.join(cells))
        elif tag == f'{W_NS}p':
            parts = [t.text or '' for t in elem.iter(f'{W_NS}t')]
            para = ''.join(parts).strip()
            if para:
                lines.append(para)
    return _normalize_text('\n'.join(lines))


def extract_text_from_pdf(path: str) -> str:
    try:
        import pdfplumber
    except ImportError as exc:
        raise ImportError('PDF 解析需要 pdfplumber') from exc
    pages: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            if text.strip():
                pages.append(text)
    return _normalize_text('\n\n'.join(pages))


def extract_text_from_image(path: str) -> str:
    """OCR via Anthropic Vision when ANTHROPIC_API_KEY is configured."""
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        raise EnvironmentError('图片 OCR 需要 ANTHROPIC_API_KEY 环境变量')
    try:
        from anthropic import Anthropic
    except ImportError as exc:
        raise ImportError('图片 OCR 需要 anthropic') from exc

    media_type = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
    }.get(Path(path).suffix.lower(), 'image/jpeg')

    with open(path, 'rb') as fh:
        data = base64.b64encode(fh.read()).decode('utf-8')

    client = Anthropic(api_key=api_key)
    message = client.messages.create(
        model='claude-3-5-sonnet-latest',
        max_tokens=4096,
        messages=[{
            'role': 'user',
            'content': [
                {
                    'type': 'image',
                    'source': {
                        'type': 'base64',
                        'media_type': media_type,
                        'data': data,
                    },
                },
                {
                    'type': 'text',
                    'text': '请完整提取图片中的所有中文表格和正文内容，保留原有顺序、换行和关键字段名称，不要总结。',
                },
            ],
        }],
    )
    parts = [getattr(item, 'text', '') for item in message.content]
    return _normalize_text('\n'.join([p for p in parts if p]))


def extract_text(path: str) -> str:
    suffix = Path(path).suffix.lower()
    if suffix == '.docx':
        return extract_text_from_docx(path)
    if suffix == '.pdf':
        return extract_text_from_pdf(path)
    if suffix in {'.jpg', '.jpeg', '.png'}:
        return extract_text_from_image(path)
    raise ValueError(f'不支持的文件格式: {suffix}')
