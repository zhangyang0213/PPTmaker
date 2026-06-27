"""PPTX生成器 - 使用python-pptx生成精美幻灯片"""

import os
import tempfile
from io import BytesIO
from typing import List, Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from styles import (
    get_theme, get_style_name, SLIDE_WIDTH, SLIDE_HEIGHT,
    STYLE_CATEGORIES, LAYOUT_NAMES,
)
from parser import SlideContent
from image_search import UnsplashSearcher, translate_keywords


# ── 布局坐标常量 (16:9, EMU) ────────────────────────────────
MARGIN = Cm(1.5)
CONTENT_TOP = Cm(3.0)
TITLE_TOP = Cm(1.2)
TITLE_LEFT = Cm(1.5)
TITLE_WIDTH = SLIDE_WIDTH - 2 * MARGIN
BODY_LEFT = Cm(1.5)
BODY_WIDTH = SLIDE_WIDTH - 2 * MARGIN
BODY_TOP = Cm(3.5)

# 图文布局中图片和文字的比例
IMAGE_FRACTION = 0.45


def create_presentation(
    slides: List[SlideContent],
    style_id: str,
    unsplash_key: str = "",
    enable_images: bool = True,
) -> Presentation:
    """创建完整的PPT演示文稿

    Args:
        slides: 解析后的幻灯片内容列表
        style_id: 风格编号 "1"-"5"
        unsplash_key: Unsplash API Access Key
        enable_images: 是否启用配图

    Returns:
        Presentation对象
    """
    theme = get_theme(style_id)
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # 使用空白布局
    blank_layout = prs.slide_layouts[6]  # 空白

    # 初始化图片搜索器
    searcher = UnsplashSearcher(unsplash_key) if unsplash_key else None

    for idx, slide_data in enumerate(slides):
        slide = prs.slides.add_slide(blank_layout)

        # 设置背景
        _set_background(slide, theme["bg_color"])

        # 根据布局类型绘制
        layout = slide_data.layout
        if layout == "title":
            _draw_title_slide(slide, slide_data, theme)
        elif layout == "section":
            _draw_section_slide(slide, slide_data, theme)
        elif layout == "content":
            _draw_content_slide(slide, slide_data, theme)
        elif layout == "image_text":
            _draw_image_text_slide(slide, slide_data, theme)
        elif layout == "toc":
            _draw_toc_slide(slide, slide_data, theme)
        elif layout == "end":
            _draw_end_slide(slide, slide_data, theme)
        else:
            _draw_content_slide(slide, slide_data, theme)

        # 添加装饰元素
        _add_decorations(slide, theme, layout)

        # 添加页码（非封面和结束页）
        if layout not in ("title", "end"):
            _add_page_number(slide, idx, len(slides), theme)

    # 异步下载并插入图片
    if enable_images and searcher:
        _insert_images(prs, slides, searcher, theme)

    return prs


# ── 各布局绘制函数 ────────────────────────────────────────────

def _draw_title_slide(slide, data: SlideContent, theme: dict):
    """封面页"""
    # 装饰横线
    line_top = Cm(5.0)
    _add_shape(slide, MARGIN, line_top, SLIDE_WIDTH - 2 * MARGIN, Cm(0.06),
               theme["accent_color"])

    # 主标题
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(2.5), TITLE_WIDTH, Cm(3.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]
    p.alignment = PP_ALIGN.CENTER

    # 副标题
    if data.subtitle:
        txBox2 = _add_textbox(slide, TITLE_LEFT, Cm(5.5), TITLE_WIDTH, Cm(2.0))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = data.subtitle
        p2.font.size = Pt(20)
        p2.font.color.rgb = theme["accent_color"]
        p2.font.name = theme["body_font"]
        p2.alignment = PP_ALIGN.CENTER


def _draw_section_slide(slide, data: SlideContent, theme: dict):
    """章节分隔页"""
    # 左侧装饰条
    _add_shape(slide, 0, 0, Cm(0.8), SLIDE_HEIGHT, theme["accent_color"])

    # 章节标题
    txBox = _add_textbox(slide, Cm(2.5), Cm(2.0), SLIDE_WIDTH - Cm(4.0), Cm(5.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT

    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = theme["section_size"]
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 底部装饰线
    _add_shape(slide, Cm(2.5), Cm(7.5), Cm(5.0), Cm(0.05), theme["accent2_color"])


def _draw_content_slide(slide, data: SlideContent, theme: dict):
    """文字内容页"""
    # 顶部装饰条
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.4), theme["accent_color"])

    # 标题
    txBox = _add_textbox(slide, TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, Cm(1.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 标题下划线
    _add_shape(slide, TITLE_LEFT, Cm(2.7), Cm(4.0), Cm(0.04), theme["accent_color"])

    # 正文内容
    body_top = Cm(3.2)
    if data.bullet_items:
        txBox2 = _add_textbox(slide, BODY_LEFT, body_top, BODY_WIDTH, Cm(12.0))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, item in enumerate(data.bullet_items):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = f"  •  {item}"
            p.font.size = Pt(18)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(8)

    elif data.body_lines:
        txBox2 = _add_textbox(slide, BODY_LEFT, body_top, BODY_WIDTH, Cm(12.0))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, line in enumerate(data.body_lines):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = line
            p.font.size = Pt(16)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(6)


def _draw_image_text_slide(slide, data: SlideContent, theme: dict):
    """图文页：左侧图片区，右侧文字"""
    # 顶部装饰条
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.4), theme["accent_color"])

    # 标题
    txBox = _add_textbox(slide, TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, Cm(1.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 图片占位区域（左侧）
    img_left = BODY_LEFT
    img_top = Cm(3.2)
    img_width = int((SLIDE_WIDTH - 3 * MARGIN) * IMAGE_FRACTION)
    img_height = Cm(12.0)
    # 绘制灰色占位框
    _add_shape(slide, img_left, img_top, img_width, img_height,
               RGBColor(0xE0, 0xE0, 0xE0))
    # 在占位框中写"配图区"
    txBox_img = _add_textbox(slide, img_left, img_top + img_height // 2 - Cm(0.5),
                              img_width, Cm(1.0))
    tf_img = txBox_img.text_frame
    p_img = tf_img.paragraphs[0]
    p_img.text = "配图区"
    p_img.font.size = Pt(14)
    p_img.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    p_img.alignment = PP_ALIGN.CENTER

    # 右侧文字区域
    text_left = img_left + img_width + MARGIN
    text_width = SLIDE_WIDTH - text_left - MARGIN
    txBox2 = _add_textbox(slide, text_left, Cm(3.2), text_width, Cm(12.0))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True

    if data.bullet_items:
        for i, item in enumerate(data.bullet_items):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = f"  •  {item}"
            p.font.size = Pt(16)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(8)

    if data.body_lines:
        for line in data.body_lines:
            p = tf2.add_paragraph()
            p.text = line
            p.font.size = Pt(14)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(6)


def _draw_toc_slide(slide, data: SlideContent, theme: dict):
    """目录页"""
    # 标题
    txBox = _add_textbox(slide, TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, Cm(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = data.title or "目录"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 目录项
    _add_shape(slide, TITLE_LEFT, Cm(2.7), Cm(4.0), Cm(0.04), theme["accent_color"])

    txBox2 = _add_textbox(slide, BODY_LEFT, Cm(3.2), BODY_WIDTH, Cm(12.0))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    for i, item in enumerate(data.bullet_items):
        if i == 0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = f"  {i+1}.  {item}"
        p.font.size = Pt(20)
        p.font.color.rgb = theme["body_color"]
        p.font.name = theme["body_font"]
        p.space_after = Pt(12)


def _draw_end_slide(slide, data: SlideContent, theme: dict):
    """结束页"""
    # 装饰横线
    _add_shape(slide, MARGIN, Cm(5.0), SLIDE_WIDTH - 2 * MARGIN, Cm(0.06),
               theme["accent_color"])

    # 感谢文字
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(3.0), TITLE_WIDTH, Cm(4.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]
    p.alignment = PP_ALIGN.CENTER

    # 副文字
    txBox2 = _add_textbox(slide, TITLE_LEFT, Cm(7.0), TITLE_WIDTH, Cm(2.0))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = "THANK YOU"
    p2.font.size = Pt(18)
    p2.font.color.rgb = theme["accent_color"]
    p2.font.name = theme["body_font"]
    p2.alignment = PP_ALIGN.CENTER


# ── 装饰元素 ─────────────────────────────────────────────────

def _add_decorations(slide, theme: dict, layout: str):
    """根据风格添加装饰元素"""
    border_style = theme["border_style"]

    if border_style == "ink":
        # 水墨风：底部淡色矩形 + 右上角小色块
        if layout not in ("title", "end"):
            _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.6), SLIDE_WIDTH, Cm(0.6),
                       theme["decorative_colors"][0], alpha=0.3)
        # 右上角装饰
        _add_shape(slide, SLIDE_WIDTH - Cm(3.0), 0, Cm(3.0), Cm(0.8),
                   theme["decorative_colors"][2], alpha=0.15)

    elif border_style == "organic":
        # 自然风：底部绿色条 + 左侧小圆角矩形
        if layout not in ("title", "end"):
            _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5),
                       theme["decorative_colors"][0], alpha=0.4)
        _add_shape(slide, 0, Cm(1.0), Cm(0.3), Cm(3.0),
                   theme["decorative_colors"][1], alpha=0.3)

    elif border_style == "chinese":
        # 国潮风：上下红色细线 + 角落金色方块
        if layout not in ("title",):
            _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.3), SLIDE_WIDTH, Cm(0.3),
                       theme["decorative_colors"][0], alpha=0.7)
        _add_shape(slide, SLIDE_WIDTH - Cm(1.2), 0, Cm(1.2), Cm(1.2),
                   theme["decorative_colors"][1], alpha=0.15)

    elif border_style == "bold":
        # 热血风：底部粗红条 + 左侧黄色竖条
        if layout not in ("title", "end"):
            _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.8), SLIDE_WIDTH, Cm(0.8),
                       theme["decorative_colors"][0], alpha=0.8)
        _add_shape(slide, 0, 0, Cm(0.4), SLIDE_HEIGHT,
                   theme["decorative_colors"][1], alpha=0.6)

    elif border_style == "wood":
        # 书香风：底部木色条 + 右侧边线
        if layout not in ("title", "end"):
            _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5),
                       theme["decorative_colors"][0], alpha=0.5)
        _add_shape(slide, SLIDE_WIDTH - Cm(0.15), 0, Cm(0.15), SLIDE_HEIGHT,
                   theme["decorative_colors"][1], alpha=0.3)


# ── 图片插入 ─────────────────────────────────────────────────

def _insert_images(prs: Presentation, slides: List[SlideContent],
                   searcher: UnsplashSearcher, theme: dict):
    """为图文页和内容页插入Unsplash配图"""
    for i, slide_data in enumerate(slides):
        if slide_data.layout not in ("image_text", "content"):
            continue
        if not slide_data.image_query:
            continue

        # 翻译关键词
        query = translate_keywords(slide_data.keywords)
        if not query:
            query = slide_data.image_query

        # 搜索图片
        results = searcher.search_images(query, per_page=1, orientation="landscape")
        if not results:
            continue

        # 下载图片
        img_data = searcher.download_image(results[0]["url"])
        if not img_data:
            continue

        try:
            img_stream = BytesIO(img_data)
            slide = prs.slides[i]

            if slide_data.layout == "image_text":
                # 图文页：替换占位区域
                img_left = int(MARGIN)
                img_top = int(Cm(3.2))
                img_width = int((SLIDE_WIDTH - 3 * MARGIN) * IMAGE_FRACTION)
                img_height = int(Cm(12.0))
                slide.shapes.add_picture(
                    img_stream, img_left, img_top, img_width, img_height
                )
            else:
                # 内容页：右下角小图
                img_width = int(Cm(5.0))
                img_height = int(Cm(3.5))
                img_left = SLIDE_WIDTH - MARGIN - img_width
                img_top = SLIDE_HEIGHT - Cm(1.0) - img_height
                slide.shapes.add_picture(
                    img_stream, img_left, img_top, img_width, img_height
                )
        except Exception as e:
            print(f"插入图片失败: {e}")
            continue


# ── 辅助函数 ─────────────────────────────────────────────────

def _add_textbox(slide, left, top, width, height):
    """添加文本框"""
    return slide.shapes.add_textbox(
        int(left), int(top), int(width), int(height)
    )


def _add_shape(slide, left, top, width, height, fill_color, alpha=1.0):
    """添加矩形装饰形状"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        int(left), int(top), int(width), int(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()  # 无边框

    # 设置透明度
    if alpha < 1.0:
        _set_shape_alpha(shape, alpha)

    return shape


def _set_shape_alpha(shape, alpha: float):
    """设置形状透明度 (0.0-1.0)"""
    try:
        fill = shape.fill._fill
        srgb = fill.find(qn('a:solidFill')).find(qn('a:srgbClr'))
        if srgb is not None:
            alpha_elem = srgb.makeelement(qn('a:alpha'), {})
            alpha_elem.set('val', str(int(alpha * 100000)))
            srgb.append(alpha_elem)
    except Exception:
        pass  # 透明度设置失败不影响基本功能


def _set_background(slide, color: RGBColor):
    """设置幻灯片背景色"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_page_number(slide, idx: int, total: int, theme: dict):
    """添加页码"""
    txBox = _add_textbox(
        slide,
        SLIDE_WIDTH - Cm(2.5),
        SLIDE_HEIGHT - Cm(1.0),
        Cm(2.0),
        Cm(0.6)
    )
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"{idx + 1} / {total}"
    p.font.size = Pt(10)
    p.font.color.rgb = theme["accent_color"]
    p.font.name = theme["body_font"]
    p.alignment = PP_ALIGN.RIGHT


def save_presentation(prs: Presentation, filepath: str):
    """保存演示文稿到文件"""
    prs.save(filepath)


def export_to_bytes(prs: Presentation) -> bytes:
    """导出演示文稿为字节流"""
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()
