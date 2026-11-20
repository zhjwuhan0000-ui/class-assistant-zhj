#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2docx.py <in.md> <out.docx>

把一个 Markdown 文件转成 Word（.docx）。
支持：标题 #/##/###、无序列表（含缩进层级）、引用 >、加粗 **x** / 斜体 *x*、
```mermaid ... mindmap ... ``` 代码块（解析为原生分级大纲，而非代码文本）。
忽略 --- 分隔线、HTML 注释、空行。

依赖：python-docx（DSH 捆绑 Python 自带）。
"""
import re
import sys

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

TOKEN_RE = re.compile(r'(\*\*[^*]+\*\*|\*[^*]+\*)')


def set_east_asia(run, east='宋体'):
    """给 run 设置中文字体，避免个别 Word 版本中文显示异常。"""
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), east)


def add_inline(paragraph, text):
    """解析 **bold** / *italic* 并逐段添加 run。"""
    for tok in TOKEN_RE.split(text):
        if not tok:
            continue
        if tok.startswith('**') and tok.endswith('**') and len(tok) >= 4:
            r = paragraph.add_run(tok[2:-2])
            r.bold = True
        elif tok.startswith('*') and tok.endswith('*') and len(tok) >= 2:
            r = paragraph.add_run(tok[1:-1])
            r.italic = True
        else:
            paragraph.add_run(tok)


def mindmap_text(s):
    """清洗 mindmap 节点行：去掉 root 前缀与各种包裹符号，返回纯文字。"""
    t = s.strip()
    t = re.sub(r'^root\s*', '', t)
    for _ in range(3):
        if t.startswith('((') and t.endswith('))'):
            t = t[2:-2].strip()
        elif t.startswith('))') and t.endswith('(('):
            t = t[2:-2].strip()
        elif t.startswith('[') and t.endswith(']'):
            t = t[1:-1].strip()
        elif t.startswith('(') and t.endswith(')'):
            t = t[1:-1].strip()
        elif t.startswith('{') and t.endswith('}'):
            t = t[1:-1].strip()
        else:
            break
    return t


def is_mindmap(code_lines):
    first = next((ln.strip() for ln in code_lines if ln.strip()), '')
    return 'mindmap' in first


def add_mindmap(doc, code_lines):
    """把 mindmap 代码块渲染成原生分级大纲（根加粗 + 分级项目符号）。"""
    nodes = []
    for ln in code_lines:
        s = ln.rstrip('\n')
        stripped = s.strip()
        if not stripped or stripped == 'mindmap':
            continue
        indent = len(s) - len(s.lstrip(' '))
        text = mindmap_text(stripped)
        if not text:
            continue
        nodes.append((indent, text))
    if not nodes:
        return
    min_indent = min(n[0] for n in nodes)
    for indent, text in nodes:
        level = max(0, (indent - min_indent) // 2)
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Pt(level * 18)
        p.paragraph_format.space_after = Pt(2)
        if level == 0:
            r = p.add_run(text)
            r.bold = True
        else:
            r = p.add_run('• ' if level == 1 else '– ')
            r.bold = True
            add_inline(p, text)


def convert(md_path, docx_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.read().split('\n')

    doc = Document()
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()

        # HTML 注释 / 空行
        if stripped.startswith('<!--') or not stripped:
            i += 1
            continue

        # 分隔线
        if re.match(r'^-{3,}$', stripped):
            i += 1
            continue

        # mermaid 代码块
        if stripped.startswith('```'):
            fence_lang = stripped[3:].strip().lower()
            if fence_lang == 'mermaid':
                i += 1
                code = []
                while i < n and not lines[i].strip().startswith('```'):
                    code.append(lines[i])
                    i += 1
                i += 1  # 跳过收尾 ```
                if is_mindmap(code):
                    add_mindmap(doc, code)
                continue
            else:
                # 普通代码块：跳过（本项目只关心 mermaid）
                i += 1
                while i < n and not lines[i].strip().startswith('```'):
                    i += 1
                i += 1
                continue

        # 标题
        m = re.match(r'^(#{1,6})\s+(.*)$', line)
        if m:
            level = len(m.group(1))
            p = doc.add_heading('', level=min(level, 4))
            r = p.add_run(m.group(2).strip())
            set_east_asia(r)
            i += 1
            continue

        # 引用
        if line.lstrip().startswith('>'):
            text = line.lstrip()[1:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(18)
            add_inline(p, text)
            i += 1
            continue

        # 无序列表（含缩进层级）
        m = re.match(r'^(\s*)[-*]\s+(.*)$', line)
        if m:
            indent = len(m.group(1))
            level = indent // 2
            text = m.group(2)
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.left_indent = Pt(level * 18)
            add_inline(p, text)
            i += 1
            continue

        # 普通段落
        p = doc.add_paragraph()
        add_inline(p, stripped)
        i += 1

    doc.save(docx_path)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.stderr.write('usage: md2docx.py <in.md> <out.docx>\n')
        sys.exit(2)
    convert(sys.argv[1], sys.argv[2])
