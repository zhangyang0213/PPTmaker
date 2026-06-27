"""PPTX生成器 - 全页布局，模板上传深度支持"""

import os
import tempfile
from io import BytesIO
from typing import List, Optional, Tuple
from copy import deepcopy
import math

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
MARGIN_SM = Cm(1.0)
IMAGE_FRACTION = 0.45


def create_presentation(
    slides: List[SlideContent],
    style_id: str,
    unsplash_key: str = "",
    enable_images: bool = True,
    template_prs: Optional[Presentation] = None,
) -> Presentation:
    """创建完整的PPT演示文稿"""
    theme = get_theme(style_id)

    if template_prs is not None:
        return _create_from_template(slides, theme, template_prs, style_id,
                                      unsplash_key, enable_images)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    blank_layout = prs.slide_layouts[6]
    searcher = UnsplashSearcher(unsplash_key) if unsplash_key else None

    for idx, slide_data in enumerate(slides):
        slide = prs.slides.add_slide(blank_layout)
        _set_background(slide, theme["bg_color"])
        layout = slide_data.layout

        if layout == "title":
            _draw_title_slide(slide, slide_data, theme)
        elif layout == "section":
            _draw_section_slide(slide, slide_data, theme)
        elif layout == "content":
            _draw_content_slide(slide, slide_data, theme)
        elif layout == "image_text":
            _draw_content_slide(slide, slide_data, theme)  # 当图文页处理为内容页
        elif layout == "toc":
            _draw_toc_slide(slide, slide_data, theme)
        elif layout == "end":
            _draw_end_slide(slide, slide_data, theme)
        else:
            _draw_content_slide(slide, slide_data, theme)

        _add_decorations(slide, theme, layout)

        if layout not in ("title", "end"):
            _add_page_number(slide, idx, len(slides), theme)

    if enable_images and searcher:
        _insert_images(prs, slides, searcher, theme)

    return prs


# ══════════════════════════════════════════════════════════════
#  模板上传生成 - 深度读取模板布局
# ══════════════════════════════════════════════════════════════

def _create_from_template(
    slides: List[SlideContent],
    theme: dict,
    template_prs: Presentation,
    style_id: str,
    unsplash_key: str,
    enable_images: bool,
) -> Presentation:
    """基于用户上传的模板生成PPT - 深度读取模板布局"""
    prs = Presentation()
    prs.slide_width = template_prs.slide_width
    prs.slide_height = template_prs.slide_height

    sw = prs.slide_width
    sh = prs.slide_height

    # ── 步骤1: 分析模板的slide layouts，识别布局类型 ──
    layout_info = _analyze_template_layouts(template_prs)

    # ── 步骤2: 逐页生成内容 ──
    searcher = UnsplashSearcher(unsplash_key) if unsplash_key else None

    for idx, slide_data in enumerate(slides):
        # 选择最匹配的模板layout
        best_layout = _pick_best_layout(slide_data.layout, layout_info, template_prs)
        slide = prs.slides.add_slide(best_layout)

        # 清除所有placeholder中的旧文本
        _clear_all_placeholders(slide)

        # 获取模板中各placeholder的位置信息
        ph_positions = _get_placeholder_positions(slide)

        # 根据布局类型填入内容
        _fill_slide_content(slide, slide_data, theme, ph_positions, sw, sh)

        # 添加页码
        if slide_data.layout not in ("title", "end"):
            _add_page_number(slide, idx, len(slides), theme)

    # 插入配图
    if enable_images and searcher:
        _insert_images(prs, slides, searcher, theme)

    return prs


def _analyze_template_layouts(template_prs: Presentation) -> dict:
    """分析模板中每种layout的特征，返回布局信息"""
    info = {}
    for i, layout in enumerate(template_prs.slide_layouts):
        ph_count = len(layout.placeholders)
        ph_types = []
        for ph in layout.placeholders:
            ph_types.append({
                "idx": ph.placeholder_format.idx,
                "type": ph.placeholder_format.type,
                "has_text": ph.has_text_frame,
            })
        # 简单启发式判断
        has_title = any(p["type"] == 1 for p in ph_types)  # TITLE=1
        has_body = any(p["type"] == 2 for p in ph_types)   # BODY=2
        has_subtitle = any(p["type"] == 3 for p in ph_types) # SUBTITLE=3

        if has_title and has_subtitle and not has_body:
            layout_type = "title"
        elif has_title and has_body:
            layout_type = "content"
        elif has_title and not has_body:
            layout_type = "section"
        else:
            layout_type = "blank"

        info[i] = {
            "layout_type": layout_type,
            "ph_count": ph_count,
            "ph_types": ph_types,
            "has_title": has_title,
            "has_body": has_body,
            "has_subtitle": has_subtitle,
        }

    return info


def _pick_best_layout(slide_layout: str, layout_info: dict,
                       template_prs: Presentation):
    """根据幻灯片类型选择最匹配的模板layout"""
    # 映射关系
    type_preference = {
        "title": ["title", "section", "content", "blank"],
        "section": ["section", "title", "content", "blank"],
        "content": ["content", "section", "blank"],
        "image_text": ["content", "section", "blank"],
        "toc": ["content", "section", "blank"],
        "end": ["title", "section", "blank"],
    }

    prefs = type_preference.get(slide_layout, ["content", "blank"])

    for pref in prefs:
        for i, info in layout_info.items():
            if info["layout_type"] == pref:
                return template_prs.slide_layouts[i]

    # fallback
    return template_prs.slide_layouts[0] if len(template_prs.slide_layouts) > 0 else template_prs.slide_layouts[-1]


def _clear_all_placeholders(slide):
    """清除slide上所有placeholder中的文本，保留格式"""
    for ph in slide.placeholders:
        if ph.has_text_frame:
            for para in ph.text_frame.paragraphs:
                for run in para.runs:
                    run.text = ""
            # 也清除直接在paragraph上的文本
            for para in ph.text_frame.paragraphs:
                if para.text:
                    para.text = ""


def _get_placeholder_positions(slide) -> dict:
    """获取模板中各placeholder的位置和尺寸"""
    positions = {}
    for ph in slide.placeholders:
        pf = ph.placeholder_format
        positions[pf.idx] = {
            "left": ph.left,
            "top": ph.top,
            "width": ph.width,
            "height": ph.height,
            "type": pf.type,
        }
    return positions


def _fill_slide_content(slide, data: SlideContent, theme: dict,
                         ph_positions: dict, sw: int, sh: int):
    """根据内容类型填充幻灯片"""
    layout = data.layout

    # 判断是否有模板placeholder可用
    has_ph = len(list(slide.placeholders)) > 0

    if has_ph:
        _fill_with_placeholders(slide, data, theme, ph_positions)
    else:
        # 无placeholder时用自由文本框填充
        _fill_with_textboxes(slide, data, theme, sw, sh)


def _fill_with_placeholders(slide, data: SlideContent, theme: dict,
                             ph_positions: dict):
    """利用模板的placeholder填入内容"""
    for ph in slide.placeholders:
        pf = ph.placeholder_format
        idx = pf.idx

        if not ph.has_text_frame:
            continue

        ph_type = pf.type
        # TITLE = 1, BODY = 2, SUBTITLE = 3, CENTER_TITLE = 4, CENTER_BODY = 5

        if ph_type in (1, 4):  # 标题类
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = data.title
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.color.rgb = theme["title_color"]
            p.font.name = theme["title_font"]

        elif ph_type == 3:  # 副标题
            tf = ph.text_frame
            tf.clear()
            if data.subtitle:
                p = tf.paragraphs[0]
                p.text = data.subtitle
                p.font.size = Pt(18)
                p.font.color.rgb = theme["accent_color"]
                p.font.name = theme["body_font"]
            elif data.layout == "end":
                p = tf.paragraphs[0]
                p.text = "THANK YOU"
                p.font.size = Pt(18)
                p.font.color.rgb = theme["accent_color"]
                p.font.name = theme["body_font"]

        elif ph_type in (2, 5):  # 正文类
            tf = ph.text_frame
            tf.clear()
            tf.word_wrap = True
            items = data.bullet_items if data.bullet_items else data.body_lines
            prefix = "  •  " if data.bullet_items else ""
            for i, item in enumerate(items):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = f"{prefix}{item}"
                p.font.size = Pt(16)
                p.font.color.rgb = theme["body_color"]
                p.font.name = theme["body_font"]
                p.space_after = Pt(8)
                p.space_before = Pt(2)

        elif idx >= 2:  # 其他placeholder也尝试填充
            tf = ph.text_frame
            tf.clear()
            # 如果有副标题且还没填
            if data.subtitle and data.layout == "title":
                p = tf.paragraphs[0]
                p.text = data.subtitle
                p.font.size = Pt(18)
                p.font.color.rgb = theme["accent_color"]
                p.font.name = theme["body_font"]


def _fill_with_textboxes(slide, data: SlideContent, theme: dict,
                          sw: int, sh: int):
    """无placeholder时用自由文本框填充整页"""
    layout = data.layout
    margin = int(Cm(1.5))

    if layout == "title":
        # 大标题 - 居中偏上
        tb = slide.shapes.add_textbox(margin, int(Cm(2.5)), sw - 2*margin, int(Cm(3.5)))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data.title
        p.font.size = Pt(44)
        p.font.bold = True
        p.font.color.rgb = theme["title_color"]
        p.font.name = theme["title_font"]
        p.alignment = PP_ALIGN.CENTER

        if data.subtitle:
            tb2 = slide.shapes.add_textbox(margin, int(Cm(6.5)), sw - 2*margin, int(Cm(2.0)))
            tf2 = tb2.text_frame
            tf2.word_wrap = True
            p2 = tf2.paragraphs[0]
            p2.text = data.subtitle
            p2.font.size = Pt(22)
            p2.font.color.rgb = theme["accent_color"]
            p2.font.name = theme["body_font"]
            p2.alignment = PP_ALIGN.CENTER

    elif layout == "end":
        tb = slide.shapes.add_textbox(margin, int(Cm(3.0)), sw - 2*margin, int(Cm(4.0)))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data.title
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.color.rgb = theme["title_color"]
        p.font.name = theme["title_font"]
        p.alignment = PP_ALIGN.CENTER

    else:
        # 内容页 - 标题占上方，正文占下方大部分
        # 标题
        tb_title = slide.shapes.add_textbox(margin, int(Cm(0.8)), sw - 2*margin, int(Cm(1.5)))
        tf = tb_title.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data.title
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = theme["title_color"]
        p.font.name = theme["title_font"]

        # 正文 - 占满下方区域
        body_top = int(Cm(2.5))
        body_height = sh - body_top - int(Cm(1.5))
        tb_body = slide.shapes.add_textbox(margin, body_top, sw - 2*margin, body_height)
        tf2 = tb_body.text_frame
        tf2.word_wrap = True

        items = data.bullet_items if data.bullet_items else data.body_lines
        prefix = "  •  " if data.bullet_items else ""
        for i, item in enumerate(items):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = f"{prefix}{item}"
            p.font.size = Pt(18)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(10)
            p.space_before = Pt(4)


# ══════════════════════════════════════════════════════════════
#  内置风格绘制 — 全页填充布局
# ══════════════════════════════════════════════════════════════

def _draw_title_slide(slide, data: SlideContent, theme: dict):
    """封面页 — 全页居中大标题"""
    dc = theme["decorative_colors"]

    # 底部渐变色块
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(5.0), SLIDE_WIDTH, Cm(5.0),
               theme["bg_color2"], alpha=0.6)

    # 顶部装饰细线
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.15), dc[0])

    # 主标题 - 大字居中
    txBox = _add_textbox(slide, Cm(2.0), Cm(2.0), SLIDE_WIDTH - Cm(4.0), Cm(4.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(48)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]
    p.alignment = PP_ALIGN.CENTER

    # 装饰横线
    _add_shape(slide, SLIDE_WIDTH // 2 - Cm(3.0), Cm(6.8), Cm(6.0), Cm(0.06),
               theme["accent_color"])
    _add_shape_diamond(slide, SLIDE_WIDTH // 2 - Cm(0.2), Cm(6.65),
                       Cm(0.4), Cm(0.4), theme["accent_color"])

    # 副标题
    if data.subtitle:
        txBox2 = _add_textbox(slide, Cm(2.0), Cm(7.3), SLIDE_WIDTH - Cm(4.0), Cm(2.5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = data.subtitle
        p2.font.size = Pt(22)
        p2.font.color.rgb = theme["accent_color"]
        p2.font.name = theme["body_font"]
        p2.alignment = PP_ALIGN.CENTER


def _draw_section_slide(slide, data: SlideContent, theme: dict):
    """章节页 — 左侧强调条 + 大标题填充大部分页面"""
    dc = theme["decorative_colors"]

    # 左侧宽装饰条
    _add_shape(slide, 0, 0, Cm(1.0), SLIDE_HEIGHT, theme["accent_color"])
    _add_shape(slide, Cm(1.0), 0, Cm(0.08), SLIDE_HEIGHT, dc[1], alpha=0.5)

    # 右侧淡色区域
    _add_shape(slide, Cm(5.0), Cm(1.0), SLIDE_WIDTH - Cm(5.5), Cm(9.0),
               theme["bg_color2"], alpha=0.4)

    # CHAPTER标签
    txBox_label = _add_textbox(slide, Cm(2.5), Cm(1.5), Cm(3.0), Cm(1.0))
    p_label = txBox_label.text_frame.paragraphs[0]
    p_label.text = "CHAPTER"
    p_label.font.size = Pt(12)
    p_label.font.color.rgb = theme["accent2_color"]
    p_label.font.name = theme["body_font"]
    p_label.font.bold = True

    # 大标题
    txBox = _add_textbox(slide, Cm(2.5), Cm(2.8), SLIDE_WIDTH - Cm(4.0), Cm(6.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = theme["section_size"]
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 底部线
    _add_shape(slide, Cm(2.5), Cm(8.5), Cm(5.0), Cm(0.05), theme["accent2_color"])
    _add_shape(slide, Cm(7.5), Cm(8.35), Cm(0.3), Cm(0.3), theme["accent_color"])


def _draw_content_slide(slide, data: SlideContent, theme: dict):
    """内容页 — 标题占上方1/4，正文占下方3/4，填满页面"""
    dc = theme["decorative_colors"]

    # 顶部强调条
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.35), theme["accent_color"])
    _add_shape(slide, 0, Cm(0.35), SLIDE_WIDTH, Cm(0.06),
               dc[3] if len(dc) > 3 else dc[0])

    # 标题 - 占上方
    txBox = _add_textbox(slide, Cm(1.5), Cm(0.7), SLIDE_WIDTH - Cm(3.0), Cm(1.8))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 标题左侧竖条
    _add_shape(slide, Cm(1.0), Cm(0.7), Cm(0.12), Cm(1.5), theme["accent_color"])

    # 标题下分隔线
    _add_shape(slide, Cm(1.5), Cm(2.6), Cm(4.0), Cm(0.04), theme["accent_color"])

    # 正文区域 - 浅色背景卡，占满下方
    _add_shape(slide, Cm(1.2), Cm(2.9), SLIDE_WIDTH - Cm(2.4), Cm(12.8),
               theme["bg_color2"], alpha=0.35, corner_radius=Cm(0.3))

    # 正文内容 - 占满卡片区域
    body_top = Cm(3.3)
    body_height = Cm(11.5)

    if data.bullet_items:
        txBox2 = _add_textbox(slide, Cm(2.0), body_top,
                               SLIDE_WIDTH - Cm(4.0), body_height)
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, item in enumerate(data.bullet_items):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = f"  •  {item}"
            p.font.size = Pt(20)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(12)
            p.space_before = Pt(4)

    elif data.body_lines:
        txBox2 = _add_textbox(slide, Cm(2.0), body_top,
                               SLIDE_WIDTH - Cm(4.0), body_height)
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, line in enumerate(data.body_lines):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = line
            p.font.size = Pt(18)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(10)
            p.space_before = Pt(4)


def _draw_toc_slide(slide, data: SlideContent, theme: dict):
    """目录页 — 占满页面"""
    # 标题
    txBox = _add_textbox(slide, Cm(1.5), Cm(1.0), SLIDE_WIDTH - Cm(3.0), Cm(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = data.title or "目录"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    _add_shape(slide, Cm(1.5), Cm(2.5), Cm(3.0), Cm(0.04), theme["accent_color"])

    # 目录项 - 占满下方
    txBox2 = _add_textbox(slide, Cm(2.0), Cm(3.0), SLIDE_WIDTH - Cm(4.0), Cm(12.0))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    for i, item in enumerate(data.bullet_items):
        if i == 0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        run_num = p.add_run()
        run_num.text = f"  {i+1}.  "
        run_num.font.size = Pt(24)
        run_num.font.color.rgb = theme["accent_color"]
        run_num.font.name = theme["body_font"]
        run_num.font.bold = True
        run_text = p.add_run()
        run_text.text = item
        run_text.font.size = Pt(22)
        run_text.font.color.rgb = theme["body_color"]
        run_text.font.name = theme["body_font"]
        p.space_after = Pt(16)


def _draw_end_slide(slide, data: SlideContent, theme: dict):
    """结束页 — 全页居中"""
    dc = theme["decorative_colors"]

    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(5.0), SLIDE_WIDTH, Cm(5.0),
               theme["bg_color2"], alpha=0.5)
    _add_shape(slide, Cm(2.0), Cm(5.8), SLIDE_WIDTH - Cm(4.0), Cm(0.06),
               theme["accent_color"])
    _add_shape_diamond(slide, SLIDE_WIDTH // 2 - Cm(0.3), Cm(5.65),
                       Cm(0.6), Cm(0.6), theme["accent_color"])

    txBox = _add_textbox(slide, Cm(2.0), Cm(2.5), SLIDE_WIDTH - Cm(4.0), Cm(3.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(52)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]
    p.alignment = PP_ALIGN.CENTER

    txBox2 = _add_textbox(slide, Cm(2.0), Cm(6.8), SLIDE_WIDTH - Cm(4.0), Cm(2.0))
    p2 = txBox2.text_frame.paragraphs[0]
    p2.text = "THANK YOU"
    p2.font.size = Pt(20)
    p2.font.color.rgb = theme["accent_color"]
    p2.font.name = theme["body_font"]
    p2.alignment = PP_ALIGN.CENTER


# ══════════════════════════════════════════════════════════════
#  装饰系统
# ══════════════════════════════════════════════════════════════

def _add_decorations(slide, theme: dict, layout: str):
    ds = theme["deco_style"]
    dc = theme["decorative_colors"]
    deco_map = {
        "ink": _deco_ink, "organic": _deco_organic, "chinese": _deco_chinese,
        "bold": _deco_bold, "wood": _deco_wood, "corporate": _deco_corporate,
        "sakura": _deco_sakura, "neon": _deco_neon,
    }
    fn = deco_map.get(ds)
    if fn:
        fn(slide, dc, theme, layout)


def _deco_ink(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.6), SLIDE_WIDTH, Cm(0.6), dc[0], alpha=0.25)
        _add_shape_oval(slide, Cm(0.3), SLIDE_HEIGHT - Cm(1.8), Cm(1.2), Cm(1.0), dc[1], alpha=0.15)
    _add_shape(slide, SLIDE_WIDTH - Cm(2.5), 0, Cm(2.5), Cm(0.6), dc[2], alpha=0.12)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(3.0), SLIDE_HEIGHT - Cm(2.5), Cm(2.5), Cm(2.0), dc[0], alpha=0.1)

def _deco_organic(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5), dc[0], alpha=0.4)
    _add_shape(slide, 0, Cm(1.5), Cm(0.25), Cm(4.0), dc[1], alpha=0.3)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(4.0), -Cm(2.0), Cm(5.0), Cm(5.0), dc[3] if len(dc)>3 else dc[0], alpha=0.1)
    _add_shape_oval(slide, Cm(0.5), SLIDE_HEIGHT - Cm(2.5), Cm(1.5), Cm(1.5), dc[2], alpha=0.15)

def _deco_chinese(slide, dc, theme, layout):
    if layout not in ("title",):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.25), SLIDE_WIDTH, Cm(0.25), dc[0], alpha=0.7)
    _add_shape(slide, SLIDE_WIDTH - Cm(1.0), 0, Cm(1.0), Cm(1.0), dc[1], alpha=0.2)
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), Cm(1.0), Cm(1.0), dc[1], alpha=0.15)
    _add_shape(slide, SLIDE_WIDTH - Cm(0.8), SLIDE_HEIGHT - Cm(0.8), Cm(0.8), Cm(0.8), dc[0], alpha=0.5)
    if layout not in ("title",):
        _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.06), dc[1], alpha=0.4)

def _deco_bold(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), SLIDE_WIDTH, Cm(1.0), dc[0], alpha=0.85)
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), SLIDE_WIDTH, Cm(0.06), dc[1], alpha=0.7)
    _add_shape(slide, 0, 0, Cm(0.35), SLIDE_HEIGHT, dc[1], alpha=0.6)
    _add_shape(slide, SLIDE_WIDTH - Cm(2.0), 0, Cm(2.0), Cm(1.5), dc[0], alpha=0.2)

def _deco_wood(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5), dc[0], alpha=0.5)
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.04), dc[4] if len(dc)>4 else dc[1], alpha=0.3)
    _add_shape(slide, SLIDE_WIDTH - Cm(0.12), 0, Cm(0.12), SLIDE_HEIGHT, dc[1], alpha=0.25)
    _add_shape(slide, Cm(0.5), Cm(0.5), Cm(0.5), Cm(0.5), dc[1], alpha=0.15)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(2.0), SLIDE_HEIGHT - Cm(2.0), Cm(1.5), Cm(1.5), dc[2], alpha=0.12)

def _deco_corporate(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.4), SLIDE_WIDTH, Cm(0.4), dc[1], alpha=0.6)
    _add_shape(slide, SLIDE_WIDTH - Cm(0.1), 0, Cm(0.1), SLIDE_HEIGHT, dc[2], alpha=0.2)
    _add_shape(slide, 0, 0, Cm(0.8), Cm(0.8), dc[1], alpha=0.5)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(2.5), SLIDE_HEIGHT - Cm(2.5), Cm(2.0), Cm(2.0), dc[0], alpha=0.15)

def _deco_sakura(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.4), SLIDE_WIDTH, Cm(0.4), dc[0], alpha=0.4)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(3.5), -Cm(1.0), Cm(4.0), Cm(4.0), dc[0], alpha=0.12)
    _add_shape_oval(slide, Cm(0.5), SLIDE_HEIGHT - Cm(2.0), Cm(1.0), Cm(1.0), dc[2], alpha=0.2)
    _add_shape(slide, SLIDE_WIDTH - Cm(0.08), Cm(2.0), Cm(0.08), Cm(5.0), dc[1], alpha=0.15)
    _add_shape_oval(slide, Cm(20.0), Cm(3.0), Cm(0.4), Cm(0.4), dc[1], alpha=0.1)
    _add_shape_oval(slide, Cm(25.0), Cm(6.0), Cm(0.3), Cm(0.3), dc[2], alpha=0.1)

def _deco_neon(slide, dc, theme, layout):
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.3), SLIDE_WIDTH, Cm(0.3), dc[0], alpha=0.7)
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.15), dc[1], alpha=0.5)
    _add_shape(slide, 0, 0, Cm(0.15), SLIDE_HEIGHT, dc[0], alpha=0.4)
    _add_shape(slide, SLIDE_WIDTH - Cm(2.0), 0, Cm(2.0), Cm(1.5), dc[2], alpha=0.3)
    _add_shape(slide, SLIDE_WIDTH - Cm(0.4), Cm(2.0), Cm(0.4), Cm(6.0), dc[1], alpha=0.15)
    _add_shape_oval(slide, Cm(18.0), Cm(3.0), Cm(0.3), Cm(0.3), dc[0], alpha=0.2)
    _add_shape_oval(slide, Cm(28.0), Cm(8.0), Cm(0.4), Cm(0.4), dc[4] if len(dc)>4 else dc[1], alpha=0.15)


# ══════════════════════════════════════════════════════════════
#  图片插入
# ══════════════════════════════════════════════════════════════

def _insert_images(prs: Presentation, slides: List[SlideContent],
                   searcher: UnsplashSearcher, theme: dict):
    for i, slide_data in enumerate(slides):
        if slide_data.layout not in ("image_text", "content"):
            continue
        if not slide_data.image_query:
            continue
        query = translate_keywords(slide_data.keywords) or slide_data.image_query
        results = searcher.search_images(query, per_page=1, orientation="landscape")
        if not results:
            continue
        img_data = searcher.download_image(results[0]["url"])
        if not img_data:
            continue
        try:
            img_stream = BytesIO(img_data)
            slide = prs.slides[i]
            # 内容页：右下角较大图片
            img_width = int(Cm(7.0))
            img_height = int(Cm(5.0))
            img_left = prs.slide_width - int(Cm(1.5)) - img_width
            img_top = prs.slide_height - int(Cm(1.5)) - img_height
            slide.shapes.add_picture(img_stream, img_left, img_top, img_width, img_height)
        except Exception as e:
            print(f"插入图片失败: {e}")


# ══════════════════════════════════════════════════════════════
#  辅助函数
# ══════════════════════════════════════════════════════════════

def _add_textbox(slide, left, top, width, height):
    return slide.shapes.add_textbox(int(left), int(top), int(width), int(height))

def _add_shape(slide, left, top, width, height, fill_color, alpha=1.0, corner_radius=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if corner_radius else MSO_SHAPE.RECTANGLE,
        int(left), int(top), int(width), int(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if alpha < 1.0:
        _set_shape_alpha(shape, alpha)
    return shape

def _add_shape_oval(slide, left, top, width, height, fill_color, alpha=1.0):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(left), int(top), int(width), int(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if alpha < 1.0:
        _set_shape_alpha(shape, alpha)
    return shape

def _add_shape_diamond(slide, left, top, width, height, fill_color, alpha=1.0):
    shape = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, int(left), int(top), int(width), int(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if alpha < 1.0:
        _set_shape_alpha(shape, alpha)
    return shape

def _set_shape_alpha(shape, alpha: float):
    try:
        fill = shape.fill._fill
        srgb = fill.find(qn('a:solidFill')).find(qn('a:srgbClr'))
        if srgb is not None:
            alpha_elem = srgb.makeelement(qn('a:alpha'), {})
            alpha_elem.set('val', str(int(alpha * 100000)))
            srgb.append(alpha_elem)
    except Exception:
        pass

def _set_background(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def _add_page_number(slide, idx: int, total: int, theme: dict):
    txBox = _add_textbox(slide, SLIDE_WIDTH - Cm(2.5), SLIDE_HEIGHT - Cm(0.8), Cm(2.0), Cm(0.5))
    p = txBox.text_frame.paragraphs[0]
    p.text = f"{idx + 1} / {total}"
    p.font.size = Pt(9)
    p.font.color.rgb = theme["accent_color"]
    p.font.name = theme["body_font"]
    p.alignment = PP_ALIGN.RIGHT

def save_presentation(prs: Presentation, filepath: str):
    prs.save(filepath)

def export_to_bytes(prs: Presentation) -> bytes:
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()
