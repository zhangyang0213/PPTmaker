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
    1. 分析模板中每个slide的类型，找到cover/toc/section/content/end模板
    2. 根据解析内容构建slide计划（自动为H2内容插入section分隔页）
    3. 创建新Presentation，按计划从模板克隆slide
    4. 对每个克隆的slide替换文本
    5. 添加页码和可选配图
    """
    import copy

    sw = template_prs.slide_width
    sh = template_prs.slide_height

    # ── 步骤1: 分析模板中每个slide ──
    template_slides_info = []
    for i, tslide in enumerate(template_prs.slides):
        info = _analyze_template_slide(tslide, sw, sh)
        info["index"] = i
        template_slides_info.append(info)

    # ── 步骤2: 按类型分组 ──
    cover_list = [s for s in template_slides_info if s["type"] == "cover"]
    toc_list = [s for s in template_slides_info if s["type"] == "toc"]
    section_list = [s for s in template_slides_info if s["type"] == "section"]
    content_list = [s for s in template_slides_info if s["type"] == "content"]
    end_list = [s for s in template_slides_info if s["type"] == "end"]

    # 筛选section模板：优先使用含有"PART"字样的正规章节分隔页，
    # 排除像"评分细则"这样的特殊页面
    part_sections = []
    for s in section_list:
        tslide = template_prs.slides[s["index"]]
        has_part = any(
            shape.has_text_frame and "part" in shape.text_frame.text.lower()
            for shape in tslide.shapes
        )
        if has_part:
            part_sections.append(s)
    if part_sections:
        section_list = part_sections

    # fallback
    if not cover_list:
        cover_list = template_slides_info[:1]
    if not section_list:
        section_list = cover_list
    if not content_list:
        content_list = template_slides_info[1:2] if len(template_slides_info) > 1 else template_slides_info[:1]
    if not end_list:
        end_list = cover_list

    # ── 步骤3: 收集所有章节名称（用于侧边目录和目录页） ──
    all_section_names = []
    section_num = 0
    for sd in slides:
        if sd.layout == "section":
            section_num += 1
            all_section_names.append(sd.title)
        elif sd.layout == "content" and sd.level <= 1:
            if not all_section_names or all_section_names[-1] != sd.title:
                section_num += 1
                all_section_names.append(sd.title)

    # ── 步骤4: 构建slide计划（自动为H2内容添加section分隔页） ──
    slide_plan = []  # [(type, data, tmpl_idx, section_num)]
    section_num = 0
    last_was_section = False

    for sd in slides:
        layout = sd.layout
        if layout == "title":
            slide_plan.append(("cover", sd, cover_list[0]["index"], 0))
        elif layout == "toc":
            slide_plan.append(("toc", sd, toc_list[0]["index"] if toc_list else content_list[0]["index"], 0))
        elif layout == "section":
            section_num += 1
            tmpl = section_list[(section_num - 1) % len(section_list)]
            slide_plan.append(("section", sd, tmpl["index"], section_num))
            last_was_section = True
        elif layout == "end":
            slide_plan.append(("end", sd, end_list[0]["index"], section_num))
        else:
            # Content: 自动添加section分隔页（H2内容前面没有section时）
            if sd.level <= 1 and not last_was_section:
                section_num += 1
                section_sd = SlideContent(layout="section", title=sd.title, level=1)
                tmpl = section_list[(section_num - 1) % len(section_list)]
                slide_plan.append(("section", section_sd, tmpl["index"], section_num))
            slide_plan.append(("content", sd, content_list[0]["index"], section_num))
            last_was_section = False

    # ── 步骤5: 创建新Presentation，按计划克隆模板slide ──
    new_prs = Presentation()
    new_prs.slide_width = sw
    new_prs.slide_height = sh

    for slide_type, data, tmpl_idx, sec_num in slide_plan:
        source_slide = template_prs.slides[tmpl_idx]
        new_slide = _clone_slide(new_prs, source_slide, template_prs)
        # 根据类型替换文本
        if slide_type == "cover":
            _replace_cover_text(new_slide, data, theme)
        elif slide_type == "toc":
            _replace_toc_text(new_slide, data, theme, all_section_names)
        elif slide_type == "section":
            _replace_section_text(new_slide, data, theme, sec_num)
        elif slide_type == "content":
            _replace_content_text(new_slide, data, theme, all_section_names, sec_num)
        elif slide_type == "end":
            _replace_end_text(new_slide, data, theme)

    # ── 步骤6: 添加页码 ──
    for i, (slide_type, _, _, _) in enumerate(slide_plan):
        if slide_type in ("content", "section"):
            slide = new_prs.slides[i]
            page_num_box = slide.shapes.add_textbox(
                int(Cm(10.2)), int(Cm(7.0)), int(Cm(3.0)), int(Cm(0.4))
            )
            p = page_num_box.text_frame.paragraphs[0]
            p.text = str(i + 1)
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            p.font.name = "微软雅黑"

    # ── 步骤7: 插入配图 ──
    if enable_images and unsplash_key:
        searcher = UnsplashSearcher(unsplash_key)
        _insert_images(new_prs, slides, searcher, theme)

    return new_prs


def _collect_text_shapes(slide):
    """收集slide中所有有文本的shape信息"""
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
    return text_shapes


def _replace_cover_text(slide, data: SlideContent, theme: dict):
    """封面页：最大文本→标题，其次→副标题，清空其余"""
    text_shapes = _collect_text_shapes(slide)
    sorted_shapes = sorted(text_shapes, key=lambda x: -x["font_size"])
    for i, ts in enumerate(sorted_shapes):
        if i == 0:
            _set_shape_text(ts["shape"], data.title)
        elif i == 1:
            _set_shape_text(ts["shape"], data.subtitle or "")
        else:
            _set_shape_text(ts["shape"], "")


def _replace_toc_text(slide, data: SlideContent, theme: dict,
                       all_section_names: list):
    """目录页：小号文本(22pt)→章节名称，大号编号和"目录"保留"""
    text_shapes = _collect_text_shapes(slide)
    # 按位置排序：小号文本在前（用于填目录项），大号编号保留
    sorted_shapes = sorted(text_shapes, key=lambda x: (x["font_size"] > 40, x["top"], x["left"]))

    toc_items = all_section_names if all_section_names else (data.bullet_items if data.bullet_items else [])
    toc_idx = 0
    for ts in sorted_shapes:
        if ts["font_size"] > 40:
            # 大号编号(01-04)和"目录"/"CONTENTS"，保留
            continue
        else:
            # 小号文本→目录标题
            if toc_idx < len(toc_items):
                _set_shape_text(ts["shape"], toc_items[toc_idx])
                toc_idx += 1
            else:
                _set_shape_text(ts["shape"], "")


def _replace_section_text(slide, data: SlideContent, theme: dict,
                           section_num: int):
    """章节分隔页：最大→章节编号(01)，次大→'01 章节名'，保留'PART ONE'等"""
    text_shapes = _collect_text_shapes(slide)
    sorted_shapes = sorted(text_shapes, key=lambda x: -x["font_size"])

    num_str = f"{section_num:02d}"
    section_title = f"{num_str} {data.title}"

    for i, ts in enumerate(sorted_shapes):
        if i == 0:
            # 最大字号→章节编号
            _set_shape_text(ts["shape"], num_str)
        elif i == 1:
            # 次大字号→章节标题（编号+名称）
            _set_shape_text(ts["shape"], section_title)
        else:
            # 保留含"PART"的文本，清空其余多余文本
            if "part" in ts["text"].lower():
                pass
            else:
                _set_shape_text(ts["shape"], "")


def _replace_content_text(slide, data: SlideContent, theme: dict,
                           all_section_names: list, section_num: int):
    """正文内容页：侧边→章节名，标题→内容标题，添加正文文本框"""
    text_shapes = _collect_text_shapes(slide)
    # 分成两组：左侧(0-2cm)是侧边目录，右侧(2.2cm+)是正文区
    sidebar_shapes = [ts for ts in text_shapes if ts["left_cm"] < 2.0]
    content_shapes = [ts for ts in text_shapes if ts["left_cm"] >= 2.0]

    # 侧边目录：填入各章节名
    sidebar_sorted = sorted(sidebar_shapes, key=lambda x: x["top"])
    for i, ts in enumerate(sidebar_sorted):
        if i < len(all_section_names):
            _set_shape_text(ts["shape"], all_section_names[i])
        else:
            _set_shape_text(ts["shape"], "")

    # 正文区：最大→标题，其余清空（正文用新增文本框）
    content_sorted = sorted(content_shapes, key=lambda x: -x["font_size"])
    for i, ts in enumerate(content_sorted):
        if i == 0:
            _set_shape_text(ts["shape"], data.title)
        else:
            _set_shape_text(ts["shape"], "")

    # 添加正文文本框（18pt，10.5x5.6cm，位于left 2.2cm, top 1.2cm）
    items = data.body_lines if data.body_lines else data.bullet_items
    if items:
        txBox = slide.shapes.add_textbox(
            int(Cm(2.2)), int(Cm(1.2)),
            int(Cm(10.5)), int(Cm(5.6))
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        for i, item in enumerate(items):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = item
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            p.font.name = "微软雅黑"
            p.space_after = Pt(6)


def _replace_end_text(slide, data: SlideContent, theme: dict):
    """结尾页：最大文本→结束语，保留其余"""
    text_shapes = _collect_text_shapes(slide)
    sorted_shapes = sorted(text_shapes, key=lambda x: -x["font_size"])
    for i, ts in enumerate(sorted_shapes):
        if i == 0:
            _set_shape_text(ts["shape"], data.title)
        else:
            # 保留其余文本（如"THANK YOU"、图片等）
            pass


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
    """复制模板slide到新演示文稿（包括背景、图片和形状）

    跨Presentation复制时，图片需要通过blob重新添加，
    非图片形状则深度复制XML元素。同时复制背景XML。
    """
    import copy

    blank_layout = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank_layout)

    # 复制背景XML元素
    src_cSld = source_slide._element.find(qn('p:cSld'))
    new_cSld = new_slide._element.find(qn('p:cSld'))
    if src_cSld is not None and new_cSld is not None:
        src_bg = src_cSld.find(qn('p:bg'))
        if src_bg is not None:
            new_bg = new_cSld.find(qn('p:bg'))
            if new_bg is not None:
                new_cSld.remove(new_bg)
            new_cSld.insert(0, copy.deepcopy(src_bg))

    # 复制所有形状（图片用add_picture，其他深拷贝XML）
    for shape in source_slide.shapes:
        try:
            img_blob = shape.image.blob
            img_stream = BytesIO(img_blob)
            new_slide.shapes.add_picture(
                img_stream, shape.left, shape.top, shape.width, shape.height
            )
            continue
        except:
            pass
        el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(el)

    return new_slide


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
