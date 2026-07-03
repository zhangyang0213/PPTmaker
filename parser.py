"""Markdown / 纯文本解析器 - 自动分页与标题层级识别"""

import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class SlideContent:
    """单页幻灯片的内容结构"""
    layout: str              # 布局类型: title, content, section, image_text, ...
    title: str = ""          # 标题
    subtitle: str = ""       # 副标题
    body_lines: List[str] = field(default_factory=list)   # 正文行列表
    bullet_items: List[str] = field(default_factory=list)  # 列表项
    level: int = 0           # 标题层级 0=H1, 1=H2, 2=H3, 3=H4
    keywords: List[str] = field(default_factory=list)      # 提取的关键词
    image_query: str = ""    # 配图搜索词


def _clean_text(text: str) -> str:
    """清理文本中的Markdown符号"""
    # 去除加粗 **text** 或 __text__
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    # 去除斜体 *text* 或 _text_（但保留列表项的 - 和 *）
    text = re.sub(r'(?<!\w)\*(?!\s)(.+?)(?<!\s)\*(?!\w)', r'\1', text)
    # 去除删除线 ~~text~~
    text = re.sub(r'~~(.+?)~~', r'\1', text)
    # 去除行内代码 `code`
    text = re.sub(r'`(.+?)`', r'\1', text)
    # 去除链接 [text](url)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    return text.strip()


def parse_markdown(text: str) -> List[SlideContent]:
    """解析Markdown文本，生成幻灯片内容列表

    规则:
    - H1 (#) → 封面页 (title layout)
    - H2 (##) → 章节页 / 内容页 (section/content layout)
    - H3/H4 → 内容页的标题 (content layout)
    - 列表项 (-, *, 1.) → bullet_items
    - 普通段落 → body_lines
    - 连续短段落合并为一页
    """
    slides = []
    lines = text.strip().split('\n')
    i = 0

    while i < len(lines):
        line = lines[i].rstrip()

        # 跳过空行
        if not line.strip():
            i += 1
            continue

        # H1 → 封面页
        if line.startswith('# ') and not line.startswith('## '):
            title = _clean_text(line[2:].strip())
            subtitle_lines = []
            # 收集H1下面所有非空非标题行作为副标题（如姓名/学号/老师等）
            j = i + 1
            while j < len(lines):
                next_line = lines[j].strip()
                if not next_line:
                    j += 1
                    continue
                if next_line.startswith('#'):
                    break
                # 短行视为副标题行
                if len(next_line) < 80:
                    subtitle_lines.append(_clean_text(next_line))
                    j += 1
                else:
                    break
            subtitle = "\n".join(subtitle_lines) if subtitle_lines else ""
            i = j
            slides.append(SlideContent(
                layout="title",
                title=title,
                subtitle=subtitle,
                level=0,
            ))
            continue

        # H2 → 章节页或内容页
        if line.startswith('## ') and not line.startswith('### '):
            title = _clean_text(line[3:].strip())
            # 收集该H2下的内容
            body_lines, bullet_items, consumed = _collect_body(lines, i + 1)
            i = i + 1 + consumed

            if not body_lines and not bullet_items:
                # 无正文 → 章节分隔页
                slides.append(SlideContent(
                    layout="section",
                    title=title,
                    level=1,
                ))
            else:
                # 有正文 → 内容页
                # 如果内容过多，自动分页
                chunks = _auto_paginate(title, body_lines, bullet_items, level=1)
                slides.extend(chunks)
            continue

        # H3 → 内容页
        if line.startswith('### ') and not line.startswith('#### '):
            title = _clean_text(line[4:].strip())
            body_lines, bullet_items, consumed = _collect_body(lines, i + 1)
            i = i + 1 + consumed
            chunks = _auto_paginate(title, body_lines, bullet_items, level=2)
            slides.extend(chunks)
            continue

        # H4 → 内容页
        if line.startswith('#### '):
            title = _clean_text(line[5:].strip())
            body_lines, bullet_items, consumed = _collect_body(lines, i + 1)
            i = i + 1 + consumed
            chunks = _auto_paginate(title, body_lines, bullet_items, level=3)
            slides.extend(chunks)
            continue

        # 无标题的纯文本段落 → 内容页（标题取首行关键词）
        if line.strip():
            body_lines, bullet_items, consumed = _collect_body(lines, i)
            i += consumed
            # 取前几个字作为标题
            first_line = body_lines[0] if body_lines else ""
            title = first_line[:20] + ("..." if len(first_line) > 20 else "")
            chunks = _auto_paginate(title, body_lines, bullet_items, level=2)
            slides.extend(chunks)
            continue

        i += 1

    # 如果没有封面页，把第一个内容页提升为封面
    if slides and slides[0].layout != "title":
        first = slides[0]
        slides[0] = SlideContent(
            layout="title",
            title=first.title,
            subtitle=first.body_lines[0] if first.body_lines else "",
            level=0,
        )

    # 添加结束页
    if slides:
        slides.append(SlideContent(layout="end", title="感谢聆听"))

    return slides


def parse_plain_text(text: str) -> List[SlideContent]:
    """解析纯文本，通过启发式规则自动分页

    规则:
    - 短行(< 30字) + 前后有空行 → 可能是标题
    - 长行 → 正文
    - 连续短句 → 可能是列表
    - 空行 → 段落/页分隔
    - 仅含列表项的段落合并到前一个段落
    """
    slides = []
    paragraphs = re.split(r'\n\s*\n', text.strip())

    if not paragraphs:
        return slides

    # 预处理：将仅含列表项的段落合并到前一个非空段落
    merged = []
    for para in paragraphs:
        lines = [l.strip() for l in para.split('\n') if l.strip()]
        if not lines:
            continue
        # 判断是否全部是列表项
        all_bullets = all(
            re.match(r'^[\-\*]\s', l) or re.match(r'^\d+[\.、)]\s', l)
            for l in lines
        )
        if all_bullets and merged:
            # 合并到前一个段落
            merged[-1] = merged[-1] + '\n' + para
        else:
            merged.append(para)

    paragraphs = merged

    # 第一段作为封面
    first_para = paragraphs[0].strip()
    first_lines = [l.strip() for l in first_para.split('\n') if l.strip()]

    title = first_lines[0] if first_lines else "未命名演示"
    subtitle = first_lines[1] if len(first_lines) > 1 else ""
    slides.append(SlideContent(
        layout="title",
        title=title,
        subtitle=subtitle,
        level=0,
    ))

    # 后续段落 → 内容页
    for para in paragraphs[1:]:
        lines = [l.strip() for l in para.split('\n') if l.strip()]
        if not lines:
            continue

        # 判断是否有标题行（短行且非列表项）
        first_is_bullet = re.match(r'^[\-\*]\s', lines[0]) or re.match(r'^\d+[\.、)]\s', lines[0])
        if len(lines[0]) < 30 and not first_is_bullet:
            slide_title = lines[0]
            body = lines[1:]
        else:
            slide_title = lines[0][:20] + "..." if not first_is_bullet else ""
            body = lines

        # 区分列表和正文
        bullet_items = []
        body_lines = []
        for line in body:
            if re.match(r'^[\-\*]\s', line):
                bullet_items.append(re.sub(r'^[\-\*]\s*', '', line))
            elif re.match(r'^\d+[\.、)]\s', line):
                bullet_items.append(re.sub(r'^\d+[\.、)]\s*', '', line))
            else:
                body_lines.append(line)

        # 自动分页
        chunks = _auto_paginate(slide_title, body_lines, bullet_items, level=1)
        slides.extend(chunks)

    # 添加结束页
    slides.append(SlideContent(layout="end", title="感谢聆听"))
    return slides


def _collect_body(lines: List[str], start: int) -> tuple:
    """从指定位置开始收集正文内容，直到遇到下一个标题或文本结束

    Returns:
        (body_lines, bullet_items, consumed_count)
    """
    body_lines = []
    bullet_items = []
    i = start
    while i < len(lines):
        line = lines[i].rstrip()
        # 遇到新标题则停止
        if line.startswith('#'):
            break
        if not line.strip():
            i += 1
            continue
        # 列表项
        if re.match(r'^[\-\*]\s', line):
            bullet_items.append(_clean_text(re.sub(r'^[\-\*]\s*', '', line)))
        elif re.match(r'^\d+[\.、)]\s', line):
            bullet_items.append(_clean_text(re.sub(r'^\d+[\.、)]\s*', '', line)))
        else:
            body_lines.append(_clean_text(line.strip()))
        i += 1

    return body_lines, bullet_items, i - start


def _auto_paginate(
    title: str,
    body_lines: List[str],
    bullet_items: List[str],
    level: int = 1,
    max_bullets_per_page: int = 6,
    max_body_lines_per_page: int = 8,
) -> List[SlideContent]:
    """自动分页：当内容过多时拆分为多页

    规则:
    - 列表项超过 max_bullets_per_page → 拆分
    - 正文行超过 max_body_lines_per_page → 拆分
    - 同时有正文和列表 → 用图文页布局
    """
    slides = []

    # 如果有图片配图需求（有正文也有列表），用图文页
    if body_lines and bullet_items:
        # 正文放左侧，列表作为要点
        all_items = bullet_items
        for start in range(0, len(all_items), max_bullets_per_page):
            chunk = all_items[start:start + max_bullets_per_page]
            # 第一页放正文，后续页只放列表
            if start == 0:
                slides.append(SlideContent(
                    layout="image_text",
                    title=title if start == 0 else f"{title}（续）",
                    body_lines=body_lines[:max_body_lines_per_page],
                    bullet_items=chunk,
                    level=level,
                ))
            else:
                slides.append(SlideContent(
                    layout="content",
                    title=f"{title}（续）",
                    bullet_items=chunk,
                    level=level,
                ))
    elif bullet_items:
        # 只有列表
        for start in range(0, len(bullet_items), max_bullets_per_page):
            chunk = bullet_items[start:start + max_bullets_per_page]
            slides.append(SlideContent(
                layout="content",
                title=title if start == 0 else f"{title}（续）",
                bullet_items=chunk,
                level=level,
            ))
    elif body_lines:
        # 只有正文
        for start in range(0, len(body_lines), max_body_lines_per_page):
            chunk = body_lines[start:start + max_body_lines_per_page]
            slides.append(SlideContent(
                layout="content",
                title=title if start == 0 else f"{title}（续）",
                body_lines=chunk,
                level=level,
            ))
    else:
        # 无内容 → 章节页
        slides.append(SlideContent(
            layout="section",
            title=title,
            level=level,
        ))

    return slides


def extract_keywords(slides: List[SlideContent]) -> List[SlideContent]:
    """为每页幻灯片提取关键词，用于配图搜索"""
    try:
        import jieba
        import jieba.analyse
        USE_JIEBA = True
    except ImportError:
        USE_JIEBA = False

    for slide in slides:
        text = slide.title + " " + " ".join(slide.body_lines) + " ".join(slide.bullet_items)
        if USE_JIEBA:
            keywords = jieba.analyse.extract_tags(text, topK=5)
        else:
            # 简单关键词提取：取标题词
            keywords = [w for w in re.findall(r'[\u4e00-\u9fff]{2,}', slide.title)][:3]
        slide.keywords = keywords
        # 取前2个关键词作为图片搜索词
        slide.image_query = " ".join(keywords[:2]) if keywords else slide.title

    return slides
