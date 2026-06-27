"""PPTX生成器 - 全页布局，模板深度读取，预览支持"""

import os
import tempfile
from io import BytesIO
from typing import List, Optional, Tuple, Dict
from copy import deepcopy
import math

from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

from styles import (
    get_theme, get_style_name, SLIDE_WIDTH, SLIDE_HEIGHT,
    STYLE_CATEGORIES, LAYOUT_NAMES,
)
from parser import SlideContent
from image_search import UnsplashSearcher, translate_keywords

MARGIN = Cm(1.5)
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
            _draw_content_slide(slide, slide_data, theme)
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
#  模板上传生成 - 深度读取模板实际slide
# ══════════════════════════════════════════════════════════════

def _create_from_template(
    slides: List[SlideContent],
    theme: dict,
    template_prs: Presentation,
    style_id: str,
    unsplash_key: str,
    enable_images: bool,
) -> Presentation:
    """基于用户上传的模板生成PPT
    
    核心原则：保留模板的一切视觉元素（背景、图片、装饰、校徽等），
    只修改文本框中的文字内容。
    
    逻辑：
    1. 直接复制整个模板文件
    2. 分析每个slide的类型
    3. 只替换文本框中的文字，不动其他任何东西
    4. 内容过多时增加正文页（复制正文页slide）
    5. 多余的模板slide保留（不做删除）
    """
    from copy import deepcopy
    import tempfile
    
    # 直接基于模板文件创建新演示文稿（保留一切）
    buf = BytesIO()
    template_prs.save(buf)
    buf.seek(0)
    prs = Presentation(buf)

    sw = prs.slide_width
    sh = prs.slide_height

    # ── 步骤1: 分析模板中每个slide ──
    template_slides_info = []
    for i, tslide in enumerate(prs.slides):
        info = _analyze_template_slide(tslide, sw, sh)
        info["index"] = i
        template_slides_info.append(info)

    # ── 步骤2: 按类型分组 ──
    cover_candidates = [s for s in template_slides_info if s["type"] == "cover"]
    toc_candidates = [s for s in template_slides_info if s["type"] == "toc"]
    section_candidates = [s for s in template_slides_info if s["type"] == "section"]
    content_candidates = [s for s in template_slides_info if s["type"] == "content"]
    end_candidates = [s for s in template_slides_info if s["type"] == "end"]

    # fallback
    if not cover_candidates:
        cover_candidates = template_slides_info[:1]
    if not section_candidates:
        section_candidates = cover_candidates
    if not content_candidates:
        content_candidates = template_slides_info[1:2] if len(template_slides_info) > 1 else template_slides_info[:1]
    if not end_candidates:
        end_candidates = cover_candidates

    # ── 步骤3: 收集所有章节名称（用于侧边目录） ──
    all_section_names = []
    for sd in slides:
        if sd.layout == "section" or (sd.layout == "content" and sd.level <= 1):
            if sd.title and sd.title not in all_section_names:
                all_section_names.append(sd.title)

    # ── 步骤4: 构建内容→模板slide映射 ──
    # 将我们的内容slides分配到模板slides
    mapping = []  # [(content_slide, template_slide_index)]
    content_idx = 0
    section_idx = 0
    
    for sd in slides:
        layout = sd.layout
        if layout == "title":
            tmpl = cover_candidates[0]
        elif layout == "toc":
            tmpl = toc_candidates[0] if toc_candidates else content_candidates[0]
        elif layout == "section":
            tmpl = section_candidates[section_idx % len(section_candidates)]
            section_idx += 1
        elif layout == "end":
            tmpl = end_candidates[0]
        else:
            tmpl = content_candidates[content_idx % len(content_candidates)]
            content_idx += 1
        mapping.append((sd, tmpl["index"]))

    # ── 步骤4: 替换模板slide中的文本 ──
    for content_sd, tmpl_idx in mapping:
        slide = prs.slides[tmpl_idx]
        _replace_text_in_slide(slide, content_sd, theme, all_section_names)

    # ── 步骤5: 如果内容页比模板正文页多，添加额外的slide ──
    # (通过复制最后一个正文页模板slide)
    extra_content = content_idx - len(content_candidates)
    if extra_content > 0 and content_candidates:
        # 需要增加额外页
        last_content_idx = content_candidates[-1]["index"]
        for extra_i in range(extra_content):
            # 找到对应的内容slide
            target_content_sd = None
            ci = len(content_candidates) + extra_i
            for sd, ti in mapping:
                pass  # mapping里已经分配完了
            
    # ── 步骤6: 清空未分配使用的模板slide中的文本 ──
    used_indices = set(ti for _, ti in mapping)
    for i, tslide_info in enumerate(template_slides_info):
        if i not in used_indices:
            slide = prs.slides[i]
            # 清空文本但保留所有形状和图片
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            run.text = ""

    # ── 步骤7: 插入配图 ──
    if enable_images and unsplash_key:
        searcher = UnsplashSearcher(unsplash_key)
        _insert_images(prs, slides, searcher, theme)

    return prs


def _replace_text_in_slide(slide, data: SlideContent, theme: dict,
                           all_section_names: list = None):
    """只替换slide中的文本内容，保留所有格式、图片、装饰不变
    
    关键理解：
    - 封面页：最大文本→标题，其次→副标题/姓名等
    - 目录页：小号***是目录标题(22pt)，大号01/02/03/04是编号(66pt)
    - 章节分隔页：大号编号保留，次大文本→章节标题
    - 正文页：左侧0-1.7cm的小文本是侧边目录，2.2cm后是正文标题
    - 结尾页：最大文本→结束语
    
    只修改run.text属性，不改font/color/size。
    """
    if all_section_names is None:
        all_section_names = []
    # 收集所有有文本的shape
    text_shapes = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue
        
        max_font = 0
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    max_font = max(max_font, run.font.size.pt)
        
        left_cm = shape.left / 914400
        
        text_shapes.append({
            "shape": shape,
            "text": text,
            "font_size": max_font,
            "top": shape.top,
            "left": shape.left,
            "left_cm": left_cm,
        })

    layout = data.layout
    items = data.bullet_items if data.bullet_items else data.body_lines

    if layout == "title":
        # 封面页：按字号排序，最大→标题，其次→副标题
        sorted_shapes = sorted(text_shapes, key=lambda x: -x["font_size"])
        for i, ts in enumerate(sorted_shapes):
            if i == 0:
                _set_shape_text(ts["shape"], data.title)
            elif i == 1 and data.subtitle:
                _set_shape_text(ts["shape"], data.subtitle)
            else:
                _set_shape_text(ts["shape"], "")

    elif layout == "section":
        # 章节分隔页：最大→编号(保留)，次大→章节标题
        sorted_shapes = sorted(text_shapes, key=lambda x: -x["font_size"])
        for i, ts in enumerate(sorted_shapes):
            if i == 0:
                # 大号编号，保留模板原始内容
                pass
            elif i == 1:
                # 章节标题
                _set_shape_text(ts["shape"], data.title)
            else:
                _set_shape_text(ts["shape"], "")

    elif layout == "toc":
        # 目录页：按位置(从上到下、从左到右)排序
        # 小号文本(22pt的***)→目录标题，大号编号(66pt)→保留
        sorted_shapes = sorted(text_shapes, key=lambda x: (x["font_size"] > 40, x["top"], x["left"]))
        
        toc_items = data.bullet_items if data.bullet_items else []
        toc_idx = 0
        for ts in sorted_shapes:
            if ts["font_size"] > 40:
                # 大号编号，保留
                continue
            else:
                # 小号文本→目录标题
                if toc_idx < len(toc_items):
                    _set_shape_text(ts["shape"], toc_items[toc_idx])
                    toc_idx += 1
                else:
                    _set_shape_text(ts["shape"], "")

    elif layout == "end":
        # 结尾页：最大文本→结束语
        sorted_shapes = sorted(text_shapes, key=lambda x: -x["font_size"])
        for i, ts in enumerate(sorted_shapes):
            if i == 0:
                _set_shape_text(ts["shape"], data.title)
            else:
                _set_shape_text(ts["shape"], "")

    else:
        # 正文内容页
        # 分成两组：左侧(0-1.7cm)是侧边目录，右侧(2.2cm+)是正文区
        sidebar_shapes = [ts for ts in text_shapes if ts["left_cm"] < 2.0]
        content_shapes = [ts for ts in text_shapes if ts["left_cm"] >= 2.0]
        
        # 侧边目录：填入各章节名
        sidebar_sorted = sorted(sidebar_shapes, key=lambda x: x["top"])
        for i, ts in enumerate(sidebar_sorted):
            if i < len(all_section_names):
                _set_shape_text(ts["shape"], all_section_names[i])
            else:
                _set_shape_text(ts["shape"], "")
        
        # 正文区：按字号排序，最大→标题，其余→正文要点
        content_sorted = sorted(content_shapes, key=lambda x: -x["font_size"])
        content_idx = 0
        for i, ts in enumerate(content_sorted):
            if i == 0:
                # 标题
                _set_shape_text(ts["shape"], data.title)
            else:
                # 正文要点
                if content_idx < len(items):
                    _set_shape_text(ts["shape"], items[content_idx])
                    content_idx += 1
                else:
                    _set_shape_text(ts["shape"], "")


def _set_shape_text(shape, new_text: str):
    """替换shape中的文本，保留第一个run的格式
    
    只修改run的text属性，不改动font、color等任何格式。
    """
    if not shape.has_text_frame:
        return
    
    tf = shape.text_frame
    
    # 找到第一个有文本的paragraph
    for para in tf.paragraphs:
        if para.runs:
            # 保留第一个run的格式，设置新文本
            para.runs[0].text = new_text
            # 清空其余runs
            for run in para.runs[1:]:
                run.text = ""
        else:
            # 没有run的段落，直接设置text
            para.text = new_text


def _analyze_template_slide(slide, sw: int, sh: int) -> dict:
    """分析模板中一个实际slide的结构
    
    识别规则：
    - cover: 有大标题(40pt+)、有姓名/学号等字段、有图片
    - section: 有"PART"字样或大号数字(96pt+)编号(01/02/03/04)
    - toc: 有"目录/CONTENTS"字样
    - content: 有标题(28pt左右) + 多个分布的小标题(20pt左右)
    - end: 有"谢谢/感谢/批评指正"等关键词
    """
    shapes_info = []
    has_big_title = False
    has_subtitle = False
    has_body = False
    has_section_marker = False
    has_toc_marker = False
    has_end_marker = False
    title_font_size = 0
    max_font_size = 0
    has_image = False
    small_title_count = 0

    for shape in slide.shapes:
        # 检测图片
        if hasattr(shape, 'image'):
            has_image = True
            continue

        if not shape.has_text_frame:
            continue

        left_cm = shape.left / 914400
        top_cm = shape.top / 914400
        width_cm = shape.width / 914400
        height_cm = shape.height / 914400
        sw_cm = sw / 914400
        sh_cm = sh / 914400

        text = shape.text_frame.text.strip()
        if not text:
            continue

        # 分析字号
        shape_max_font = 0
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    font_pt = run.font.size.pt
                    shape_max_font = max(shape_max_font, font_pt)
                    max_font_size = max(max_font_size, font_pt)

        is_top = top_cm < sh_cm * 0.4

        # 检测章节标记: 大号数字(96pt) 或 "PART" 字样
        text_lower = text.lower().strip()
        if shape_max_font >= 60 and (text.isdigit() or len(text) <= 3):
            has_section_marker = True
        if "part" in text_lower:
            has_section_marker = True

        # 检测目录标记
        if any(kw in text_lower for kw in ["目录", "contents", "目 录"]):
            has_toc_marker = True

        # 检测结尾标记
        if any(kw in text_lower for kw in ["谢谢", "感谢", "thank", "批评指正", "结束", "聆听"]):
            has_end_marker = True

        # 检测大标题(40pt+)
        if shape_max_font >= 40 and is_top:
            has_big_title = True
            title_font_size = max(title_font_size, shape_max_font)

        # 检测小标题(16-24pt，在页面中下部分)
        if 16 <= shape_max_font <= 24 and not is_top:
            small_title_count += 1

        # 检测正文标题(28pt左右)
        if 24 <= shape_max_font <= 32:
            has_body = True

        # 检测姓名/学号等字段
        if any(kw in text for kw in ["姓名", "学号", "老师", "姓名："]):
            has_subtitle = True

        shapes_info.append({
            "shape": shape,
            "text": text,
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
            "font_size": shape_max_font,
            "is_top": is_top,
            "is_bottom": not is_top,
            "is_center": left_cm > sw_cm * 0.2 and (left_cm + width_cm) < sw_cm * 0.8,
            "para_count": len(shape.text_frame.paragraphs),
        })

    # ── 判断slide类型 ──
    if has_end_marker:
        slide_type = "end"
    elif has_toc_marker:
        slide_type = "toc"
    elif has_section_marker:
        slide_type = "section"
    elif has_big_title and has_image:
        # 有大标题+图片 → 封面页
        slide_type = "cover"
    elif has_big_title and has_subtitle:
        slide_type = "cover"
    elif has_big_title and not has_body and not has_section_marker:
        # 只有大标题没有正文 → 封面页
        slide_type = "cover"
    elif has_body or small_title_count >= 2:
        slide_type = "content"
    else:
        slide_type = "content"  # 默认

    # 收集可填充的文本区域
    text_areas = []
    for si in shapes_info:
        if si["text"]:
            text_areas.append({
                "left": si["left"],
                "top": si["top"],
                "width": si["width"],
                "height": si["height"],
                "font_size": si["font_size"],
                "is_top": si["is_top"],
                "is_bottom": si["is_bottom"],
                "para_count": si["para_count"],
            })

    return {
        "type": slide_type,
        "shapes": shapes_info,
        "text_areas": text_areas,
        "has_big_title": has_big_title,
        "has_body": has_body,
        "has_image": has_image,
        "has_section_marker": has_section_marker,
        "small_title_count": small_title_count,
        "max_font_size": max_font_size,
    }


def _clone_slide(prs: Presentation, source_slide, source_prs: Presentation):
    """复制模板slide到新演示文稿（包括背景和形状）"""
    # 使用source的layout创建slide
    blank_layout = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank_layout)

    # 复制背景
    try:
        bg = source_slide.background
        fill = bg.fill
        if fill.type is not None:
            new_slide.background.fill.solid()
            try:
                new_slide.background.fill.fore_color.rgb = fill.fore_color.rgb
            except:
                pass
    except:
        pass

    # 复制所有形状
    for shape in source_slide.shapes:
        el = deepcopy(shape._element)
        new_slide.shapes._spTree.append(el)

    return new_slide


def _clear_all_text(slide):
    """清空slide上所有文本框中的文本，保留位置和格式"""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            # 保留段落格式，只清除文本
            for run in para.runs:
                run.text = ""
            # 也清除直接在paragraph上的文本
            if para.text and not para.runs:
                # 用一个空run替换
                run = para.add_run()
                run.text = ""


def _fill_template_slide(slide, data: SlideContent, theme: dict,
                          tmpl_info: dict, sw: int, sh: int):
    """根据分析出的文本区域位置填入内容"""
    text_areas = tmpl_info["text_areas"]
    layout = data.layout

    # 按位置排序：上方的先，下方的后
    sorted_areas = sorted(text_areas, key=lambda a: (a["top"], a["left"]))

    if layout == "title":
        _fill_title_from_template(slide, data, theme, sorted_areas, sw, sh)
    elif layout == "end":
        _fill_end_from_template(slide, data, theme, sorted_areas, sw, sh)
    elif layout == "toc":
        _fill_toc_from_template(slide, data, theme, sorted_areas, sw, sh)
    else:
        _fill_content_from_template(slide, data, theme, sorted_areas, sw, sh)


def _fill_title_from_template(slide, data, theme, areas, sw, sh):
    """填充封面页"""
    title_area = None
    subtitle_area = None

    for area in areas:
        if area["is_top"] and (title_area is None or area["font_size"] > title_area["font_size"]):
            title_area = area
        elif area["is_top"] and title_area and area["font_size"] < title_area["font_size"]:
            subtitle_area = area

    # 如果没找到合适区域，用第一个作为标题
    if title_area is None and areas:
        title_area = areas[0]

    # 找到或创建文本框
    shapes = list(slide.shapes)
    
    if title_area:
        shape = _find_shape_at(slide, title_area)
        if shape and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = data.title
            p.font.size = Pt(max(36, title_area["font_size"] if title_area["font_size"] > 0 else 36))
            p.font.bold = True
            p.font.color.rgb = theme["title_color"]
            p.font.name = theme["title_font"]
            p.alignment = PP_ALIGN.CENTER

    if subtitle_area and data.subtitle:
        shape = _find_shape_at(slide, subtitle_area)
        if shape and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = data.subtitle
            p.font.size = Pt(max(18, subtitle_area["font_size"] if subtitle_area["font_size"] > 0 else 18))
            p.font.color.rgb = theme["accent_color"]
            p.font.name = theme["body_font"]
            p.alignment = PP_ALIGN.CENTER

    # 如果只有标题没有副标题区域，但有副标题内容，在标题下方添加
    if data.subtitle and not subtitle_area and title_area:
        sub_top = title_area["top"] + title_area["height"]
        _add_textbox_to_slide(slide, data.subtitle, theme["accent_color"],
                               Pt(18), theme["body_font"],
                               title_area["left"], sub_top + Cm(0.5),
                               title_area["width"], Cm(2.0),
                               align=PP_ALIGN.CENTER)


def _fill_end_from_template(slide, data, theme, areas, sw, sh):
    """填充结尾页"""
    # 用最大的文本区域放"感谢聆听"
    if areas:
        main_area = max(areas, key=lambda a: a["font_size"] if a["font_size"] > 0 else a["width"] * a["height"])
        shape = _find_shape_at(slide, main_area)
        if shape and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = data.title
            p.font.size = Pt(max(36, main_area["font_size"] if main_area["font_size"] > 0 else 36))
            p.font.bold = True
            p.font.color.rgb = theme["title_color"]
            p.font.name = theme["title_font"]
            p.alignment = PP_ALIGN.CENTER

    # 在其他区域写"THANK YOU"
    for area in areas:
        if area is not areas[0] if areas else False:
            shape = _find_shape_at(slide, area)
            if shape and shape.has_text_frame:
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = "THANK YOU"
                p.font.size = Pt(16)
                p.font.color.rgb = theme["accent_color"]
                p.font.name = theme["body_font"]
                p.alignment = PP_ALIGN.CENTER


def _fill_toc_from_template(slide, data, theme, areas, sw, sh):
    """填充目录页"""
    title_area = None
    body_area = None

    for area in areas:
        if area["is_top"] and title_area is None:
            title_area = area
        elif area["is_bottom"] and body_area is None:
            body_area = area

    if not areas:
        return

    if title_area is None:
        title_area = areas[0]
    if body_area is None and len(areas) > 1:
        body_area = areas[-1]

    # 填标题
    shape = _find_shape_at(slide, title_area)
    if shape and shape.has_text_frame:
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data.title or "目录"
        p.font.size = Pt(max(28, title_area["font_size"] if title_area["font_size"] > 0 else 28))
        p.font.bold = True
        p.font.color.rgb = theme["title_color"]
        p.font.name = theme["title_font"]

    # 填目录项
    if body_area and data.bullet_items:
        shape = _find_shape_at(slide, body_area)
        if shape and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            tf.word_wrap = True
            for i, item in enumerate(data.bullet_items):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                run_num = p.add_run()
                run_num.text = f"{i+1}. "
                run_num.font.size = Pt(20)
                run_num.font.color.rgb = theme["accent_color"]
                run_num.font.bold = True
                run_num.font.name = theme["body_font"]
                run_text = p.add_run()
                run_text.text = item
                run_text.font.size = Pt(18)
                run_text.font.color.rgb = theme["body_color"]
                run_text.font.name = theme["body_font"]
                p.space_after = Pt(8)


def _fill_content_from_template(slide, data, theme, areas, sw, sh):
    """填充内容页 - 标题+正文"""
    title_area = None
    body_area = None

    # 找最上方的大字区域作为标题
    for area in sorted(areas, key=lambda a: a["top"]):
        if title_area is None:
            title_area = area
        else:
            body_area = area
            break

    if not areas:
        # 没有检测到文本区域，创建自由文本框
        _fill_with_free_textboxes(slide, data, theme, sw, sh)
        return

    if body_area is None and len(areas) > 1:
        body_area = areas[-1]
    if body_area is None:
        # 只有一个区域，上半部分标题，下半部分正文
        title_area = areas[0]
        # 用标题区域的下半部分作为正文区域
        half_height = title_area["height"] // 2
        body_area = {
            "left": title_area["left"],
            "top": title_area["top"] + half_height,
            "width": title_area["width"],
            "height": half_height,
            "font_size": 0,
        }

    # 填标题
    shape = _find_shape_at(slide, title_area)
    if shape and shape.has_text_frame:
        tf = shape.text_frame
        tf.clear()
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = data.title
        p.font.size = Pt(max(26, title_area["font_size"] if title_area["font_size"] > 0 else 26))
        p.font.bold = True
        p.font.color.rgb = theme["title_color"]
        p.font.name = theme["title_font"]

    # 填正文
    items = data.bullet_items if data.bullet_items else data.body_lines
    prefix = "  •  " if data.bullet_items else ""

    if body_area and items:
        shape = _find_shape_at(slide, body_area)
        if shape and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            tf.word_wrap = True
            # 计算可容纳的行数（根据区域高度和字号）
            font_size = Pt(max(16, body_area["font_size"] if body_area["font_size"] > 0 else 16))
            line_height_emu = int(font_size * 2.0)  # 行高约2倍字号
            area_height_emu = body_area["height"]
            max_lines = max(1, area_height_emu // line_height_emu)

            for i, item in enumerate(items[:max_lines]):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = f"{prefix}{item}"
                p.font.size = font_size
                p.font.color.rgb = theme["body_color"]
                p.font.name = theme["body_font"]
                p.space_after = Pt(6)
        else:
            # 找不到shape，用文本框
            _add_textbox_to_slide(slide, "", theme["body_color"], Pt(16),
                                   theme["body_font"],
                                   body_area["left"], body_area["top"],
                                   body_area["width"], body_area["height"])
            # 重新找
            shape = _find_shape_at(slide, body_area)
            if shape and shape.has_text_frame:
                tf = shape.text_frame
                tf.word_wrap = True
                font_size = Pt(max(16, body_area["font_size"] if body_area["font_size"] > 0 else 16))
                for i, item in enumerate(items):
                    if i == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.text = f"{prefix}{item}"
                    p.font.size = font_size
                    p.font.color.rgb = theme["body_color"]
                    p.font.name = theme["body_font"]
                    p.space_after = Pt(6)


def _fill_with_free_textboxes(slide, data, theme, sw, sh):
    """无可用文本区域时，用自由文本框填充"""
    margin = int(Cm(1.5))
    layout = data.layout

    if layout == "title":
        _add_textbox_to_slide(slide, data.title, theme["title_color"],
                               Pt(44), theme["title_font"],
                               margin, int(Cm(2.5)), sw - 2*margin, int(Cm(3.5)),
                               bold=True, align=PP_ALIGN.CENTER)
        if data.subtitle:
            _add_textbox_to_slide(slide, data.subtitle, theme["accent_color"],
                                   Pt(22), theme["body_font"],
                                   margin, int(Cm(6.5)), sw - 2*margin, int(Cm(2.0)),
                                   align=PP_ALIGN.CENTER)
    elif layout == "end":
        _add_textbox_to_slide(slide, data.title, theme["title_color"],
                               Pt(48), theme["title_font"],
                               margin, int(Cm(3.0)), sw - 2*margin, int(Cm(4.0)),
                               bold=True, align=PP_ALIGN.CENTER)
    else:
        # 标题
        _add_textbox_to_slide(slide, data.title, theme["title_color"],
                               Pt(28), theme["title_font"],
                               margin, int(Cm(0.8)), sw - 2*margin, int(Cm(1.5)),
                               bold=True)
        # 正文
        items = data.bullet_items if data.bullet_items else data.body_lines
        prefix = "  •  " if data.bullet_items else ""
        body_top = int(Cm(2.5))
        body_height = sh - body_top - int(Cm(1.5))
        text = "\n".join(f"{prefix}{item}" for item in items)
        _add_textbox_to_slide(slide, text, theme["body_color"],
                               Pt(18), theme["body_font"],
                               margin, body_top, sw - 2*margin, body_height)


def _find_shape_at(slide, area: dict):
    """在slide上找到位于指定位置的shape"""
    best_match = None
    best_overlap = 0

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        # 计算位置重叠度
        overlap_left = max(shape.left, area["left"])
        overlap_top = max(shape.top, area["top"])
        overlap_right = min(shape.left + shape.width, area["left"] + area["width"])
        overlap_bottom = min(shape.top + shape.height, area["top"] + area["height"])

        if overlap_right > overlap_left and overlap_bottom > overlap_top:
            overlap = (overlap_right - overlap_left) * (overlap_bottom - overlap_top)
            if overlap > best_overlap:
                best_overlap = overlap
                best_match = shape

    return best_match


def _add_textbox_to_slide(slide, text, color, size, font_name,
                            left, top, width, height,
                            bold=False, align=PP_ALIGN.LEFT):
    """在slide上添加文本框"""
    txBox = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = size
    p.font.color.rgb = color
    p.font.name = font_name
    p.font.bold = bold
    p.alignment = align
    return txBox


# ══════════════════════════════════════════════════════════════
#  预览功能 - 生成幻灯片缩略图
# ══════════════════════════════════════════════════════════════

def generate_preview_html(prs: Presentation) -> str:
    """生成PPT预览的HTML，直接在浏览器中渲染（解决中文乱码问题）"""
    slides_html = []
    sw_cm = prs.slide_width / 914400
    sh_cm = prs.slide_height / 914400

    for page_idx, slide in enumerate(prs.slides):
        # 获取背景色
        bg_color = "#FFFFFF"
        try:
            bg = slide.background
            fill = bg.fill
            if fill.type is not None:
                bg_rgb = fill.fore_color.rgb
                bg_color = f"#{str(bg_rgb)[0:2]}{str(bg_rgb)[2:4]}{str(bg_rgb)[4:6]}"
        except:
            pass

        shapes_html = []
        for shape in slide.shapes:
            left_pct = shape.left / prs.slide_width * 100
            top_pct = shape.top / prs.slide_height * 100
            width_pct = shape.width / prs.slide_width * 100
            height_pct = shape.height / prs.slide_height * 100

            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if not text:
                    continue

                # 获取字号
                font_size = 14
                font_color = "#333333"
                font_bold = False
                try:
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.size:
                                font_size = max(font_size, int(run.font.size.pt * 0.55))
                            if run.font.color and run.font.color.rgb:
                                c = run.font.color.rgb
                                font_color = f"#{str(c)[0:2]}{str(c)[2:4]}{str(c)[4:6]}"
                            if run.font.bold:
                                font_bold = True
                except:
                    pass

                # 转义HTML特殊字符
                text_escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                # 截断过长文本
                lines = text_escaped.split('\n')
                display_lines = '<br>'.join(lines[:8])
                if len(lines) > 8:
                    display_lines += '...'

                bold_css = "font-weight:bold;" if font_bold else ""
                shapes_html.append(
                    f'<div style="position:absolute;left:{left_pct}%;top:{top_pct}%;'
                    f'width:{width_pct}%;height:{height_pct}%;'
                    f'font-size:{font_size}px;color:{font_color};{bold_css}'
                    f'overflow:hidden;word-wrap:break-word;padding:2px;">'
                    f'{display_lines}</div>'
                )

            elif hasattr(shape, 'image'):
                # 图片占位
                shapes_html.append(
                    f'<div style="position:absolute;left:{left_pct}%;top:{top_pct}%;'
                    f'width:{width_pct}%;height:{height_pct}%;'
                    f'background:#e0e0e0;border-radius:4px;display:flex;'
                    f'align-items:center;justify-content:center;color:#999;font-size:10px;">'
                    f'图片</div>'
                )

        slides_html.append(
            f'<div style="position:relative;width:100%;padding-bottom:{sh_cm/sw_cm*100}%;'
            f'background:{bg_color};border:1px solid #ddd;border-radius:4px;margin-bottom:8px;">'
            f'{"".join(shapes_html)}</div>'
        )

    return "".join(slides_html)


def generate_preview_images(prs: Presentation) -> List[bytes]:
    """生成PPT每页的预览图 - 已废弃，改用HTML预览"""
    return []


# ══════════════════════════════════════════════════════════════
#  内置风格绘制 — 全页填充布局
# ══════════════════════════════════════════════════════════════

def _draw_title_slide(slide, data, theme):
    dc = theme["decorative_colors"]
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(5.0), SLIDE_WIDTH, Cm(5.0), theme["bg_color2"], alpha=0.6)
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.15), dc[0])

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

    _add_shape(slide, SLIDE_WIDTH // 2 - Cm(3.0), Cm(6.8), Cm(6.0), Cm(0.06), theme["accent_color"])
    _add_shape_diamond(slide, SLIDE_WIDTH // 2 - Cm(0.2), Cm(6.65), Cm(0.4), Cm(0.4), theme["accent_color"])

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


def _draw_section_slide(slide, data, theme):
    dc = theme["decorative_colors"]
    _add_shape(slide, 0, 0, Cm(1.0), SLIDE_HEIGHT, theme["accent_color"])
    _add_shape(slide, Cm(1.0), 0, Cm(0.08), SLIDE_HEIGHT, dc[1], alpha=0.5)
    _add_shape(slide, Cm(5.0), Cm(1.0), SLIDE_WIDTH - Cm(5.5), Cm(9.0), theme["bg_color2"], alpha=0.4)

    txBox_label = _add_textbox(slide, Cm(2.5), Cm(1.5), Cm(3.0), Cm(1.0))
    p_label = txBox_label.text_frame.paragraphs[0]
    p_label.text = "CHAPTER"
    p_label.font.size = Pt(12)
    p_label.font.color.rgb = theme["accent2_color"]
    p_label.font.name = theme["body_font"]
    p_label.font.bold = True

    txBox = _add_textbox(slide, Cm(2.5), Cm(2.8), SLIDE_WIDTH - Cm(4.0), Cm(6.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = theme["section_size"]
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    _add_shape(slide, Cm(2.5), Cm(8.5), Cm(5.0), Cm(0.05), theme["accent2_color"])
    _add_shape(slide, Cm(7.5), Cm(8.35), Cm(0.3), Cm(0.3), theme["accent_color"])


def _draw_content_slide(slide, data, theme):
    dc = theme["decorative_colors"]
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.35), theme["accent_color"])
    _add_shape(slide, 0, Cm(0.35), SLIDE_WIDTH, Cm(0.06), dc[3] if len(dc)>3 else dc[0])

    txBox = _add_textbox(slide, Cm(1.5), Cm(0.7), SLIDE_WIDTH - Cm(3.0), Cm(1.8))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    _add_shape(slide, Cm(1.0), Cm(0.7), Cm(0.12), Cm(1.5), theme["accent_color"])
    _add_shape(slide, Cm(1.5), Cm(2.6), Cm(4.0), Cm(0.04), theme["accent_color"])

    _add_shape(slide, Cm(1.2), Cm(2.9), SLIDE_WIDTH - Cm(2.4), Cm(12.8),
               theme["bg_color2"], alpha=0.35, corner_radius=Cm(0.3))

    body_top = Cm(3.3)
    body_height = Cm(11.5)

    if data.bullet_items:
        txBox2 = _add_textbox(slide, Cm(2.0), body_top, SLIDE_WIDTH - Cm(4.0), body_height)
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
        txBox2 = _add_textbox(slide, Cm(2.0), body_top, SLIDE_WIDTH - Cm(4.0), body_height)
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


def _draw_toc_slide(slide, data, theme):
    txBox = _add_textbox(slide, Cm(1.5), Cm(1.0), SLIDE_WIDTH - Cm(3.0), Cm(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = data.title or "目录"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = theme["title_color"]
    p.font.name = theme["title_font"]

    _add_shape(slide, Cm(1.5), Cm(2.5), Cm(3.0), Cm(0.04), theme["accent_color"])

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


def _draw_end_slide(slide, data, theme):
    dc = theme["decorative_colors"]
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(5.0), SLIDE_WIDTH, Cm(5.0), theme["bg_color2"], alpha=0.5)
    _add_shape(slide, Cm(2.0), Cm(5.8), SLIDE_WIDTH - Cm(4.0), Cm(0.06), theme["accent_color"])
    _add_shape_diamond(slide, SLIDE_WIDTH // 2 - Cm(0.3), Cm(5.65), Cm(0.6), Cm(0.6), theme["accent_color"])

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

def _add_decorations(slide, theme, layout):
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
    _add_shape(slide, SLIDE_WIDTH - Cm(2.5), 0, Cm(2.5), Cm(0.6), dc[2], alpha=0.12)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(3.0), SLIDE_HEIGHT - Cm(2.5), Cm(2.5), Cm(2.0), dc[0], alpha=0.1)

def _deco_organic(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5), dc[0], alpha=0.4)
    _add_shape(slide, 0, Cm(1.5), Cm(0.25), Cm(4.0), dc[1], alpha=0.3)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(4.0), -Cm(2.0), Cm(5.0), Cm(5.0), dc[3] if len(dc)>3 else dc[0], alpha=0.1)

def _deco_chinese(slide, dc, theme, layout):
    if layout not in ("title",):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.25), SLIDE_WIDTH, Cm(0.25), dc[0], alpha=0.7)
    _add_shape(slide, SLIDE_WIDTH - Cm(1.0), 0, Cm(1.0), Cm(1.0), dc[1], alpha=0.2)
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), Cm(1.0), Cm(1.0), dc[1], alpha=0.15)

def _deco_bold(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(1.0), SLIDE_WIDTH, Cm(1.0), dc[0], alpha=0.85)
    _add_shape(slide, 0, 0, Cm(0.35), SLIDE_HEIGHT, dc[1], alpha=0.6)
    _add_shape(slide, SLIDE_WIDTH - Cm(2.0), 0, Cm(2.0), Cm(1.5), dc[0], alpha=0.2)

def _deco_wood(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.5), SLIDE_WIDTH, Cm(0.5), dc[0], alpha=0.5)
    _add_shape(slide, SLIDE_WIDTH - Cm(0.12), 0, Cm(0.12), SLIDE_HEIGHT, dc[1], alpha=0.25)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(2.0), SLIDE_HEIGHT - Cm(2.0), Cm(1.5), Cm(1.5), dc[2], alpha=0.12)

def _deco_corporate(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.4), SLIDE_WIDTH, Cm(0.4), dc[1], alpha=0.6)
    _add_shape(slide, 0, 0, Cm(0.8), Cm(0.8), dc[1], alpha=0.5)

def _deco_sakura(slide, dc, theme, layout):
    if layout not in ("title", "end"):
        _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.4), SLIDE_WIDTH, Cm(0.4), dc[0], alpha=0.4)
    _add_shape_oval(slide, SLIDE_WIDTH - Cm(3.5), -Cm(1.0), Cm(4.0), Cm(4.0), dc[0], alpha=0.12)
    _add_shape_oval(slide, Cm(20.0), Cm(3.0), Cm(0.4), Cm(0.4), dc[1], alpha=0.1)

def _deco_neon(slide, dc, theme, layout):
    _add_shape(slide, 0, SLIDE_HEIGHT - Cm(0.3), SLIDE_WIDTH, Cm(0.3), dc[0], alpha=0.7)
    _add_shape(slide, 0, 0, SLIDE_WIDTH, Cm(0.15), dc[1], alpha=0.5)
    _add_shape(slide, 0, 0, Cm(0.15), SLIDE_HEIGHT, dc[0], alpha=0.4)


# ══════════════════════════════════════════════════════════════
#  图片插入
# ══════════════════════════════════════════════════════════════

def _insert_images(prs, slides, searcher, theme):
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

def _set_shape_alpha(shape, alpha):
    try:
        fill = shape.fill._fill
        srgb = fill.find(qn('a:solidFill')).find(qn('a:srgbClr'))
        if srgb is not None:
            alpha_elem = srgb.makeelement(qn('a:alpha'), {})
            alpha_elem.set('val', str(int(alpha * 100000)))
            srgb.append(alpha_elem)
    except:
        pass

def _set_background(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def _add_page_number(slide, idx, total, theme):
    txBox = _add_textbox(slide, SLIDE_WIDTH - Cm(2.5), SLIDE_HEIGHT - Cm(0.8), Cm(2.0), Cm(0.5))
    p = txBox.text_frame.paragraphs[0]
    p.text = f"{idx + 1} / {total}"
    p.font.size = Pt(9)
    p.font.color.rgb = theme["accent_color"]
    p.font.name = theme["body_font"]
    p.alignment = PP_ALIGN.RIGHT

def save_presentation(prs, filepath):
    prs.save(filepath)

def export_to_bytes(prs):
    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()
