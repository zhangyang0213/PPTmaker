"""PPTX生成器 - 精致排版，8种风格装饰，模板上传支持"""

import os
import tempfile
from io import BytesIO
from typing import List, Optional
from copy import deepcopy

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
TITLE_TOP = Cm(1.2)
TITLE_LEFT = Cm(1.5)
TITLE_WIDTH = SLIDE_WIDTH - 2 * MARGIN
BODY_LEFT = Cm(1.5)
BODY_WIDTH = SLIDE_WIDTH - 2 * MARGIN

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

    # 如果用户上传了模板，基于模板创建
    if template_prs is not None:
        prs = _create_from_template(slides, theme, template_prs, style_id,
                                     unsplash_key, enable_images)
        return prs

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
            _draw_image_text_slide(slide, slide_data, theme)
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


def _create_from_template(
    slides: List[SlideContent],
    theme: dict,
    template_prs: Presentation,
    style_id: str,
    unsplash_key: str,
    enable_images: bool,
) -> Presentation:
    """基于用户上传的模板生成PPT"""
    # 拷贝模板的尺寸和slide layouts
    prs = Presentation()
    prs.slide_width = template_prs.slide_width
    prs.slide_height = template_prs.slide_height

    # 收集模板中的slide layouts
    template_layouts = template_prs.slide_layouts

    # 确定使用哪个布局：优先使用模板的布局
    # 通常 layouts: 0=title, 1=title+content, 2=section, 6=blank
    def _pick_layout(layout_type: str):
        """根据幻灯片类型选择模板布局"""
        layout_map = {
            "title": [0, 1],      # 封面布局
            "section": [2, 0, 1], # 章节布局
            "content": [1, 3],    # 内容布局
            "image_text": [1, 3],
            "toc": [1, 3],
            "end": [0, 6],
        }
        candidates = layout_map.get(layout_type, [6])
        for c in candidates:
            if c < len(template_layouts):
                return template_layouts[c]
        return template_layouts[-1]  # fallback

    searcher = UnsplashSearcher(unsplash_key) if unsplash_key else None

    for idx, slide_data in enumerate(slides):
        layout = _pick_layout(slide_data.layout)
        slide = prs.slides.add_slide(layout)

        # 尝试在模板布局的placeholder中填入内容
        _fill_placeholders(slide, slide_data, theme)

        # 添加页码
        if slide_data.layout not in ("title", "end"):
            _add_page_number(slide, idx, len(slides), theme)

    # 插入配图
    if enable_images and searcher:
        _insert_images(prs, slides, searcher, theme)

    return prs


def _fill_placeholders(slide, data: SlideContent, theme: dict):
    """将内容填入模板的placeholder中"""
    for shape in slide.placeholders:
        ph = shape.placeholder_format
        # idx 0 = 标题, idx 1 = 正文
        if ph.idx == 0 and shape.has_text_frame:
            p = shape.text_frame.paragraphs[0]
            p.text = data.title
            _apply_font(p, theme["title_font"], Pt(28), theme["title_color"], bold=True)
        elif ph.idx == 1 and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            items = data.bullet_items if data.bullet_items else data.body_lines
            prefix = "  •  " if data.bullet_items else ""
            for i, item in enumerate(items):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = f"{prefix}{item}"
                _apply_font(p, theme["body_font"], Pt(16), theme["body_color"])
                p.space_after = Pt(6)

        # 处理其他placeholder（副标题等）
        elif shape.has_text_frame and ph.idx >= 2:
            if data.subtitle and shape.placeholder_format.type == 3:  # SUBTITLE
                p = shape.text_frame.paragraphs[0]
                p.text = data.subtitle
                _apply_font(p, theme["body_font"], Pt(18), theme["accent_color"])


def _apply_font(paragraph, font_name, font_size, font_color, bold=False):
    """应用字体样式到段落"""
    paragraph.font.name = font_name
    paragraph.font.size = font_size
    paragraph.font.color.rgb = font_color
    paragraph.font.bold = bold


# ══════════════════════════════════════════════════════════════
#  各布局绘制函数 — 精致版
# ══════════════════════════════════════════════════════════════

def _draw_title_slide(slide, data: SlideContent, theme: dict):
    """封面页 — 大标题居中 + 装饰线 + 副标题"""
    dc = theme["decorative_colors"]

    # 背景渐变底色条（底部1/3）
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(5.0), SLIDE_WIDTH, Cm(5.0),
               theme["bg_color2"], alpha=0.6)

    # 顶部装饰细线
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.15), dc[0])

    # 主标题
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(2.5), TITLE_WIDTH, Cm(3.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]
    p.alignment = PP_ALIGN.CENTER

    # 标题下装饰横线
    _add_shape(slide, SLIDE_WIDTH // 2 - Cm(3.0), Cm(6.2), Cm(6.0), Cm(0.06),
               theme["accent_color"])

    # 装饰小菱形（线中央）
    _add_shape_diamond(slide, SLIDE_WIDTH // 2 - Cm(0.2), Cm(6.05),
                       Cm(0.4), Cm(0.4), theme["accent_color"])

    # 副标题
    if data.subtitle:
        txBox2 = _add_textbox(slide, TITLE_LEFT, Cm(6.8), TITLE_WIDTH, Cm(2.0))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = data.subtitle
        p2.font.size = Pt(20)
        p2.font.color.rgb = theme["accent_color"]
        p2.font.name = theme["body_font"]
        p2.alignment = PP_ALIGN.CENTER


def _draw_section_slide(slide, data: SlideContent, theme: dict):
    """章节页 — 左侧强调条 + 大标题 + 底部线"""
    dc = theme["decorative_colors"]

    # 左侧宽装饰条
    _add_shape(slide, 0, 0, Cm(1.0), SLIDE_HEIGHT, theme["accent_color"])
    # 装饰条内细线
    _add_shape(slide, Cm(1.0), 0, Cm(0.08), SLIDE_HEIGHT, dc[1], alpha=0.5)

    # 右侧大块淡色区域
    _add_shape(slide, Cm(5.0), Cm(1.5), SLIDE_WIDTH - Cm(5.5), Cm(8.0),
               theme["bg_color2"], alpha=0.4)

    # 章节编号或小标签
    txBox_label = _add_textbox(slide, Cm(2.5), Cm(1.5), Cm(2.0), Cm(1.0))
    tf_label = txBox_label.text_frame
    p_label = tf_label.paragraphs[0]
    p_label.text = "CHAPTER"
    p_label.font.size = Pt(12)
    p_label.font.color.rgb = theme["accent2_color"]
    p_label.font.name = theme["body_font"]
    p_label.font.bold = True

    # 章节标题
    txBox = _add_textbox(slide, Cm(2.5), Cm(2.5), SLIDE_WIDTH - Cm(4.0), Cm(4.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = theme["section_size"]
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 底部装饰线 + 小方块
    _add_shape(slide, Cm(2.5), Cm(7.5), Cm(5.0), Cm(0.05), theme["accent2_color"])
    _add_shape(slide, Cm(7.5), Cm(7.35), Cm(0.3), Cm(0.3), theme["accent_color"])


def _draw_content_slide(slide, data: SlideContent, theme: dict):
    """内容页 — 顶部条 + 标题 + 分隔线 + 正文"""
    dc = theme["decorative_colors"]

    # 顶部强调条
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.35), theme["accent_color"])
    # 条下细线
    _add_shape(slide, 0, Cm(0.35), SLIDE_WIDTH, Cm(0.06), dc[3] if len(dc) > 3 else dc[0])

    # 标题
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(0.8), TITLE_WIDTH, Cm(1.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 标题左侧小竖条
    _add_shape(slide, Cm(1.0), Cm(0.8), Cm(0.12), Cm(1.2), theme["accent_color"])

    # 标题下分隔线
    _add_shape(slide, TITLE_LEFT, Cm(2.4), Cm(4.0), Cm(0.04), theme["accent_color"])

    # 正文区域 - 带浅色背景卡
    card_top = Cm(2.7)
    card_height = Cm(12.5)
    _add_shape(slide, BODY_LEFT, card_top, BODY_WIDTH, card_height,
               theme["bg_color2"], alpha=0.35, corner_radius=Cm(0.3))

    body_top = Cm(3.0)
    if data.bullet_items:
        txBox2 = _add_textbox(slide, Cm(2.0), body_top, BODY_WIDTH - Cm(1.0), Cm(11.5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, item in enumerate(data.bullet_items):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = f"  •  {item}"
            p.font.size = Pt(17)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(10)

    elif data.body_lines:
        txBox2 = _add_textbox(slide, Cm(2.0), body_top, BODY_WIDTH - Cm(1.0), Cm(11.5))
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        for i, line in enumerate(data.body_lines):
            if i == 0:
                p = tf2.paragraphs[0]
            else:
                p = tf2.add_paragraph()
            p.text = line
            p.font.size = Pt(15)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(8)


def _draw_image_text_slide(slide, data: SlideContent, theme: dict):
    """图文页：左侧图片区，右侧文字"""
    dc = theme["decorative_colors"]

    # 顶部强调条
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.35), theme["accent_color"])

    # 标题
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(0.8), TITLE_WIDTH, Cm(1.5))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 标题左侧小竖条
    _add_shape(slide, Cm(1.0), Cm(0.8), Cm(0.12), Cm(1.2), theme["accent_color"])

    # 图片占位区域（左侧）- 圆角灰色框
    img_left = BODY_LEFT
    img_top = Cm(2.7)
    img_width = int((SLIDE_WIDTH - 3 * MARGIN) * IMAGE_FRACTION)
    img_height = Cm(12.5)
    _add_shape(slide, img_left, img_top, img_width, img_height,
               RGBColor(0xE8, 0xE8, 0xE8), corner_radius=Cm(0.3))
    # 占位文字
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
    txBox2 = _add_textbox(slide, text_left, Cm(2.7), text_width, Cm(12.5))
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
            p.space_after = Pt(10)

    if data.body_lines:
        for line in data.body_lines:
            p = tf2.add_paragraph()
            p.text = line
            p.font.size = Pt(14)
            p.font.color.rgb = theme["body_color"]
            p.font.name = theme["body_font"]
            p.space_after = Pt(8)


def _draw_toc_slide(slide, data: SlideContent, theme: dict):
    """目录页"""
    dc = theme["decorative_colors"]

    # 标题
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(1.0), TITLE_WIDTH, Cm(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = data.title or "目录"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    # 标题下线
    _add_shape(slide, TITLE_LEFT, Cm(2.5), Cm(3.0), Cm(0.04), theme["accent_color"])

    # 目录项
    txBox2 = _add_textbox(slide, BODY_LEFT, Cm(3.0), BODY_WIDTH, Cm(12.0))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    for i, item in enumerate(data.bullet_items):
        if i == 0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        # 编号用强调色
        run_num = p.add_run()
        run_num.text = f"  {i+1}.  "
        run_num.font.size = Pt(22)
        run_num.font.color.rgb = theme["accent_color"]
        run_num.font.name = theme["body_font"]
        run_num.font.bold = True
        # 内容用正文色
        run_text = p.add_run()
        run_text.text = item
        run_text.font.size = Pt(20)
        run_text.font.color.rgb = theme["body_color"]
        run_text.font.name = theme["body_font"]
        p.space_after = Pt(14)


def _draw_end_slide(slide, data: SlideContent, theme: dict):
    """结束页"""
    dc = theme["decorative_colors"]

    # 底部渐变色块
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(5.0), SLIDE_WIDTH, Cm(5.0),
               theme["bg_color2"], alpha=0.5)

    # 装饰横线
    _add_shape(slide, MARGIN, Cm(5.5), SLIDE_WIDTH - 2 * MARGIN, Cm(0.06),
               theme["accent_color"])

    # 中央菱形装饰
    _add_shape_diamond(slide, SLIDE_WIDTH // 2 - Cm(0.3), Cm(5.35),
                       Cm(0.6), Cm(0.6), theme["accent_color"])

    # 感谢文字
    txBox = _add_textbox(slide, TITLE_LEFT, Cm(2.5), TITLE_WIDTH, Cm(3.5))
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
    txBox2 = _add_textbox(slide, TITLE_LEFT, Cm(6.5), TITLE_WIDTH, Cm(2.0))
    tf2 = txBox2.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = "THANK YOU"
    p2.font.size = Pt(18)
    p2.font.color.rgb = theme["accent_color"]
    p2.font.name = theme["body_font"]
    p2.alignment = PP_ALIGN.CENTER


# ══════════════════════════════════════════════════════════════
#  装饰系统 — 每种风格独特视觉特征
# ══════════════════════════════════════════════════════════════

def _add_decorations(slide, theme: dict, layout: str):
    """根据风格添加装饰元素"""
    ds = theme["deco_style"]
    dc = theme["decorative_colors"]

    if ds == "ink":
        _deco_ink(slide, dc, theme, layout)
    elif ds == "organic":
        _deco_organic(slide, dc, theme, layout)
    elif ds == "chinese":
        _deco_chinese(slide, dc, theme, layout)
    elif ds == "bold":
        _deco_bold(slide, dc, theme, layout)
    elif ds == "wood":
        _deco_wood(slide, dc, theme, layout)
    elif ds == "corporate":
        _deco_corporate(slide, dc, theme, layout)
    elif ds == "sakura":
        _deco_sakura(slide, dc, theme, layout)
    elif ds == "neon":
        _deco_neon(slide, dc, theme, layout)


def _deco_ink(slide, dc, theme, layout):
    """水墨风装饰：墨点、晕染边缘、留白意境"""
    if layout not in ("title", "end"):
        # 底部淡墨晕染条
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.6), SLIDE_WIDTH, Cm(0.6),
                   dc[0], alpha=0.25)
        # 左下角墨点
        _add_shape_oval(slide, Cm(0.3), SLIDE_HEIGHT - Cm(1.8), Cm(1.2), Cm(1.0),
                        dc[1], alpha=0.15)
    # 右上角小墨块
    _add_shape(slide, SLIDE_WIDTH - Cm(2.5), 0, Cm(2.5), Cm(0.6),
               dc[2], alpha=0.12)
    # 右下角淡墨圆
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(3.0), SLIDE_HEIGHT - Cm(2.5),
                    Cm(2.5), Cm(2.0), dc[0], alpha=0.1)


def _deco_organic(slide, dc, theme, layout):
    """自然风装饰：叶片色块、有机弧线"""
    if layout not in ("title", "end"):
        # 底部绿色渐变条
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5),
                   dc[0], alpha=0.4)
    # 左侧竖叶条
    _add_shape(slide, 0, Cm(1.5), Cm(0.25), Cm(4.0), dc[1], alpha=0.3)
    # 右上角圆弧装饰
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(4.0), -Cm(2.0), Cm(5.0), Cm(5.0),
                    dc[3] if len(dc) > 3 else dc[0], alpha=0.1)
    # 左下角小圆
    _add_shape_oval(slide, Cm(0.5), SLIDE_HEIGHT - Cm(2.5), Cm(1.5), Cm(1.5),
                    dc[2], alpha=0.15)


def _deco_chinese(slide, dc, theme, layout):
    """国潮风装饰：红色边框线、金色角标、回纹"""
    if layout not in ("title",):
        # 底部红色细线
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.25), SLIDE_WIDTH, Cm(0.25),
                   dc[0], alpha=0.7)
    # 右上角金色方块
    _add_shape(slide, SLIDE_WIDTH - Cm(1.0), 0, Cm(1.0), Cm(1.0),
               dc[1], alpha=0.2)
    # 左下角金色方块
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), Cm(1.0), Cm(1.0),
               dc[1], alpha=0.15)
    # 右下角红色小三角
    _add_shape(slide, SLIDE_WIDTH - Cm(0.8), SLIDE_HEIGHT - Cm(0.8),
               Cm(0.8), Cm(0.8), dc[0], alpha=0.5)
    # 顶部金色细线
    if layout not in ("title",):
        _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.06), dc[1], alpha=0.4)


def _deco_bold(slide, dc, theme, layout):
    """热血风装饰：粗红条、黄色竖线、星形"""
    if layout not in ("title", "end"):
        # 底部粗红条
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), SLIDE_WIDTH, Cm(1.0),
                   dc[0], alpha=0.85)
        # 底部条上细黄线
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), SLIDE_WIDTH, Cm(0.06),
                   dc[1], alpha=0.7)
    # 左侧黄色竖条
    _add_shape(slide, 0, 0, Cm(0.35), SLIDE_HEIGHT, dc[1], alpha=0.6)
    # 右上角红色三角
    _add_shape(slide, SLIDE_WIDTH - Cm(2.0), 0, Cm(2.0), Cm(1.5),
               dc[0], alpha=0.2)
    # 右下角小黄色条
    _add_shape(slide, SLIDE_WIDTH - Cm(0.8), SLIDE_HEIGHT - Cm(3.0),
               Cm(0.15), Cm(2.0), dc[4] if len(dc) > 4 else dc[1], alpha=0.3)


def _deco_wood(slide, dc, theme, layout):
    """书香风装饰：木色横纹、竖线、角落印记"""
    if layout not in ("title", "end"):
        # 底部木色条
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5),
                   dc[0], alpha=0.5)
        # 条上深色细线
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.04),
                   dc[4] if len(dc) > 4 else dc[1], alpha=0.3)
    # 右侧细边线
    _add_shape(slide, SLIDE_WIDTH - Cm(0.12), 0, Cm(0.12), SLIDE_HEIGHT,
               dc[1], alpha=0.25)
    # 左上角小方块印记
    _add_shape(slide, Cm(0.5), Cm(0.5), Cm(0.5), Cm(0.5), dc[1], alpha=0.15)
    # 右下角圆印
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(2.0), SLIDE_HEIGHT - Cm(2.0),
                    Cm(1.5), Cm(1.5), dc[2], alpha=0.12)


def _deco_corporate(slide, dc, theme, layout):
    """商务风装饰：蓝色条块、几何线、整洁布局"""
    if layout not in ("title", "end"):
        # 底部蓝色条
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.4), SLIDE_WIDTH, Cm(0.4),
                   dc[1], alpha=0.6)
    # 右侧竖线
    _add_shape(slide, SLIDE_WIDTH - Cm(0.1), 0, Cm(0.1), SLIDE_HEIGHT,
               dc[2], alpha=0.2)
    # 左上角蓝色方块
    _add_shape(slide, 0, 0, Cm(0.8), Cm(0.8), dc[1], alpha=0.5)
    # 右下角浅蓝圆
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(2.5), SLIDE_HEIGHT - Cm(2.5),
                    Cm(2.0), Cm(2.0), dc[0], alpha=0.15)


def _deco_sakura(slide, dc, theme, layout):
    """樱花风装饰：粉色圆点、柔美曲线、花瓣印"""
    if layout not in ("title", "end"):
        # 底部粉色条
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.4), SLIDE_WIDTH, Cm(0.4),
                   dc[0], alpha=0.4)
    # 右上角粉色圆（花瓣）
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(3.5), -Cm(1.0), Cm(4.0), Cm(4.0),
                    dc[0], alpha=0.12)
    # 左下角小圆
    _add_shape_oval(slide, Cm(0.5), SLIDE_HEIGHT - Cm(2.0), Cm(1.0), Cm(1.0),
                    dc[2], alpha=0.2)
    # 右侧细线
    _add_shape(slide, SLIDE_WIDTH - Cm(0.08), Cm(2.0), Cm(0.08), Cm(5.0),
               dc[1], alpha=0.15)
    # 散落的小圆点
    _add_shape_oval(slide, Cm(20.0), Cm(3.0), Cm(0.4), Cm(0.4), dc[1], alpha=0.1)
    _add_shape_oval(slide, Cm(25.0), Cm(6.0), Cm(0.3), Cm(0.3), dc[2], alpha=0.1)
    _add_shape_oval(slide, Cm(22.0), Cm(12.0), Cm(0.5), Cm(0.5), dc[0], alpha=0.08)


def _deco_neon(slide, dc, theme, layout):
    """科技风装饰：霓虹线条、光效方块、电路线"""
    # 底部霓虹线
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.3), SLIDE_WIDTH, Cm(0.3),
               dc[0], alpha=0.7)
    # 顶部霓虹线
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.15), dc[1], alpha=0.5)
    # 左侧竖霓虹线
    _add_shape(slide, 0, 0, Cm(0.15), SLIDE_HEIGHT, dc[0], alpha=0.4)
    # 右上角光效方块
    _add_shape(slide, SLIDE_WIDTH - Cm(2.0), 0, Cm(2.0), Cm(1.5),
               dc[2], alpha=0.3)
    # 右侧光条
    _add_shape(slide, SLIDE_WIDTH - Cm(0.4), Cm(2.0), Cm(0.4), Cm(6.0),
               dc[1], alpha=0.15)
    # 左下角小方块
    _add_shape(slide, Cm(0.5), SLIDE_HEIGHT - Cm(1.5), Cm(0.6), Cm(0.6),
               dc[4] if len(dc) > 4 else dc[0], alpha=0.3)
    # 散点光斑
    _add_shape_oval(slide, Cm(18.0), Cm(3.0), Cm(0.3), Cm(0.3), dc[0], alpha=0.2)
    _add_shape_oval(slide, Cm(28.0), Cm(8.0), Cm(0.4), Cm(0.4), dc[4] if len(dc) > 4 else dc[1], alpha=0.15)


# ══════════════════════════════════════════════════════════════
#  图片插入
# ══════════════════════════════════════════════════════════════

def _insert_images(prs: Presentation, slides: List[SlideContent],
                   searcher: UnsplashSearcher, theme: dict):
    """为图文页和内容页插入Unsplash配图"""
    for i, slide_data in enumerate(slides):
        if slide_data.layout not in ("image_text", "content"):
            continue
        if not slide_data.image_query:
            continue

        query = translate_keywords(slide_data.keywords)
        if not query:
            query = slide_data.image_query

        results = searcher.search_images(query, per_page=1, orientation="landscape")
        if not results:
            continue

        img_data = searcher.download_image(results[0]["url"])
        if not img_data:
            continue

        try:
            img_stream = BytesIO(img_data)
            slide = prs.slides[i]

            if slide_data.layout == "image_text":
                img_left = int(MARGIN)
                img_top = int(Cm(2.7))
                img_width = int((SLIDE_WIDTH - 3 * MARGIN) * IMAGE_FRACTION)
                img_height = int(Cm(12.5))
                slide.shapes.add_picture(img_stream, img_left, img_top, img_width, img_height)
            else:
                img_width = int(Cm(5.0))
                img_height = int(Cm(3.5))
                img_left = SLIDE_WIDTH - MARGIN - img_width
                img_top = SLIDE_HEIGHT - Cm(1.5) - img_height
                slide.shapes.add_picture(img_stream, img_left, img_top, img_width, img_height)
        except Exception as e:
            print(f"插入图片失败: {e}")
            continue


# ══════════════════════════════════════════════════════════════
#  辅助函数
# ══════════════════════════════════════════════════════════════

def _add_textbox(slide, left, top, width, height):
    return slide.shapes.add_textbox(int(left), int(top), int(width), int(height))


def _add_shape(slide, left, top, width, height, fill_color, alpha=1.0,
               corner_radius=None):
    """添加矩形装饰，支持圆角"""
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
    """添加椭圆装饰"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        int(left), int(top), int(width), int(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()

    if alpha < 1.0:
        _set_shape_alpha(shape, alpha)
    return shape


def _add_shape_diamond(slide, left, top, width, height, fill_color, alpha=1.0):
    """添加菱形装饰"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.DIAMOND,
        int(left), int(top), int(width), int(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()

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
        pass


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
        SLIDE_HEIGHT - Cm(0.8),
        Cm(2.0),
        Cm(0.5)
    )
    tf = txBox.text_frame
    p = tf.paragraphs[0]
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
