"""风格主题定义 - 配色、字体、装饰方案（8种风格）"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 幻灯片尺寸 (16:9, 单位: EMU)
SLIDE_WIDTH = 12192000
SLIDE_HEIGHT = 6858000

# 布局模板类型
LAYOUT_TITLE = "title"
LAYOUT_TOC = "toc"
LAYOUT_SECTION = "section"
LAYOUT_CONTENT = "content"
LAYOUT_IMAGE_TEXT = "image_text"
LAYOUT_FULL_IMAGE = "full_image"
LAYOUT_TWO_COL = "two_col"
LAYOUT_END = "end"

LAYOUT_NAMES = {
    LAYOUT_TITLE: "封面页",
    LAYOUT_TOC: "目录页",
    LAYOUT_SECTION: "章节页",
    LAYOUT_CONTENT: "文字内容页",
    LAYOUT_IMAGE_TEXT: "图文页",
    LAYOUT_FULL_IMAGE: "全图页",
    LAYOUT_TWO_COL: "双栏页",
    LAYOUT_END: "结束页",
}

# 风格大类定义
STYLE_CATEGORIES = {
    "1": {
        "name": "🎨 水墨儒雅风",
        "desc": "黑白灰为主，点缀墨蓝，淡雅如水墨画",
    },
    "2": {
        "name": "🌿 自然清新风",
        "desc": "清新绿意，明亮自然，舒适宜人",
    },
    "3": {
        "name": "🏮 华丽国潮风",
        "desc": "中国红与金色交织，传统纹样点缀",
    },
    "4": {
        "name": "🔥 热血五四风",
        "desc": "红黑撞色，刚劲有力，热血激昂",
    },
    "5": {
        "name": "📚 书香木纹风",
        "desc": "木色书香，温润厚重，儒雅内敛",
    },
    "6": {
        "name": "💼 商务简约风",
        "desc": "深蓝灰调，沉稳大气，专业干练",
    },
    "7": {
        "name": "🌸 樱花浪漫风",
        "desc": "粉白渐变，柔美浪漫，温暖治愈",
    },
    "8": {
        "name": "🚀 科技未来风",
        "desc": "深色基底，霓虹光效，酷炫未来感",
    },
}


def _theme(
    bg_color, bg_color2, title_color, body_color, accent_color, accent2_color,
    title_font, body_font, title_size, body_size, section_size,
    decorative_colors, border_style, deco_style
):
    """构造主题字典"""
    return {
        "bg_color": bg_color,
        "bg_color2": bg_color2,          # 渐变终止色
        "title_color": title_color,
        "body_color": body_color,
        "accent_color": accent_color,
        "accent2_color": accent2_color,
        "title_font": title_font,
        "body_font": body_font,
        "title_size": title_size,
        "body_size": body_size,
        "section_size": section_size,
        "decorative_colors": decorative_colors,
        "border_style": border_style,
        "deco_style": deco_style,         # 装饰方案标识
    }


THEMES = {
    # ── 1. 水墨儒雅风 ─────────────────────────────────────
    "1": _theme(
        bg_color=RGBColor(0xF5, 0xF2, 0xEB),
        bg_color2=RGBColor(0xE8, 0xE2, 0xD6),
        title_color=RGBColor(0x1A, 0x1A, 0x1A),
        body_color=RGBColor(0x3D, 0x3D, 0x3D),
        accent_color=RGBColor(0x3B, 0x5E, 0x7A),
        accent2_color=RGBColor(0x8B, 0x73, 0x55),
        title_font="华文中宋",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xC8, 0xC0, 0xAD),
            RGBColor(0x9E, 0x94, 0x84),
            RGBColor(0x3B, 0x5E, 0x7A),
            RGBColor(0xD4, 0xCE, 0xC0),
            RGBColor(0x6B, 0x5E, 0x4E),
        ],
        border_style="ink",
        deco_style="ink",
    ),
    # ── 2. 自然清新风 ─────────────────────────────────────
    "2": _theme(
        bg_color=RGBColor(0xF2, 0xFA, 0xEF),
        bg_color2=RGBColor(0xE0, 0xF0, 0xD8),
        title_color=RGBColor(0x1B, 0x5E, 0x20),
        body_color=RGBColor(0x2E, 0x4E, 0x2E),
        accent_color=RGBColor(0x43, 0xA0, 0x47),
        accent2_color=RGBColor(0x7C, 0xB3, 0x42),
        title_font="等线",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xA5, 0xD6, 0xA7),
            RGBColor(0x66, 0xBB, 0x6A),
            RGBColor(0x7C, 0xB3, 0x42),
            RGBColor(0xDC, 0xED, 0xC8),
            RGBColor(0x33, 0x69, 0x1E),
        ],
        border_style="organic",
        deco_style="organic",
    ),
    # ── 3. 华丽国潮风 ─────────────────────────────────────
    "3": _theme(
        bg_color=RGBColor(0xFE, 0xF3, 0xE0),
        bg_color2=RGBColor(0xFD, 0xE8, 0xC8),
        title_color=RGBColor(0x8B, 0x00, 0x00),
        body_color=RGBColor(0x4A, 0x2C, 0x0A),
        accent_color=RGBColor(0xC6, 0x28, 0x28),
        accent2_color=RGBColor(0xD4, 0xA0, 0x17),
        title_font="华文行楷",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xC6, 0x28, 0x28),
            RGBColor(0xD4, 0xA0, 0x17),
            RGBColor(0xFF, 0xF8, 0xE1),
            RGBColor(0xE5, 0x73, 0x73),
            RGBColor(0xFF, 0xD5, 0x4F),
        ],
        border_style="chinese",
        deco_style="chinese",
    ),
    # ── 4. 热血五四风 ─────────────────────────────────────
    "4": _theme(
        bg_color=RGBColor(0x14, 0x14, 0x14),
        bg_color2=RGBColor(0x1F, 0x1F, 0x1F),
        title_color=RGBColor(0xFF, 0xFF, 0xFF),
        body_color=RGBColor(0xD0, 0xD0, 0xD0),
        accent_color=RGBColor(0xE5, 0x39, 0x35),
        accent2_color=RGBColor(0xFF, 0xD6, 0x00),
        title_font="黑体",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(48),
        decorative_colors=[
            RGBColor(0xE5, 0x39, 0x35),
            RGBColor(0xFF, 0xD6, 0x00),
            RGBColor(0x42, 0x42, 0x42),
            RGBColor(0xFF, 0x8A, 0x80),
            RGBColor(0xFF, 0xE0, 0x82),
        ],
        border_style="bold",
        deco_style="bold",
    ),
    # ── 5. 书香木纹风 ─────────────────────────────────────
    "5": _theme(
        bg_color=RGBColor(0xF5, 0xEB, 0xD6),
        bg_color2=RGBColor(0xE8, 0xDA, 0xC0),
        title_color=RGBColor(0x33, 0x1A, 0x00),
        body_color=RGBColor(0x4E, 0x34, 0x2E),
        accent_color=RGBColor(0x6D, 0x4C, 0x41),
        accent2_color=RGBColor(0xA1, 0x88, 0x7F),
        title_font="华文中宋",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xBC, 0xAA, 0x94),
            RGBColor(0x8D, 0x6E, 0x63),
            RGBColor(0xA1, 0x88, 0x7F),
            RGBColor(0xD7, 0xCC, 0xC8),
            RGBColor(0x5D, 0x40, 0x37),
        ],
        border_style="wood",
        deco_style="wood",
    ),
    # ── 6. 商务简约风 ─────────────────────────────────────
    "6": _theme(
        bg_color=RGBColor(0xF5, 0xF7, 0xFA),
        bg_color2=RGBColor(0xE8, 0xEC, 0xF1),
        title_color=RGBColor(0x1A, 0x23, 0x7E),
        body_color=RGBColor(0x37, 0x47, 0x4F),
        accent_color=RGBColor(0x1E, 0x88, 0xE5),
        accent2_color=RGBColor(0x42, 0xA5, 0xF5),
        title_font="微软雅黑",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xBB, 0xDE, 0xFB),
            RGBColor(0x1E, 0x88, 0xE5),
            RGBColor(0x42, 0xA5, 0xF5),
            RGBColor(0x90, 0xCA, 0xF9),
            RGBColor(0x0D, 0x47, 0xA1),
        ],
        border_style="corporate",
        deco_style="corporate",
    ),
    # ── 7. 樱花浪漫风 ─────────────────────────────────────
    "7": _theme(
        bg_color=RGBColor(0xFF, 0xF5, 0xF8),
        bg_color2=RGBColor(0xFC, 0xE4, 0xEC),
        title_color=RGBColor(0x88, 0x0E, 0x4F),
        body_color=RGBColor(0x5C, 0x2D, 0x42),
        accent_color=RGBColor(0xEC, 0x40, 0x7A),
        accent2_color=RGBColor(0xF4, 0x8F, 0xB1),
        title_font="等线",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xF8, 0xBB, 0xD0),
            RGBColor(0xEC, 0x40, 0x7A),
            RGBColor(0xF4, 0x8F, 0xB1),
            RGBColor(0xFC, 0xE4, 0xEC),
            RGBColor(0xAD, 0x14, 0x57),
        ],
        border_style="sakura",
        deco_style="sakura",
    ),
    # ── 8. 科技未来风 ─────────────────────────────────────
    "8": _theme(
        bg_color=RGBColor(0x0A, 0x0E, 0x27),
        bg_color2=RGBColor(0x0D, 0x14, 0x3A),
        title_color=RGBColor(0x00, 0xE5, 0xFF),
        body_color=RGBColor(0xB0, 0xBE, 0xC5),
        accent_color=RGBColor(0x00, 0xE5, 0xFF),
        accent2_color=RGBColor(0x7C, 0x4D, 0xFF),
        title_font="等线",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(48),
        decorative_colors=[
            RGBColor(0x00, 0xE5, 0xFF),
            RGBColor(0x7C, 0x4D, 0xFF),
            RGBColor(0x1A, 0x23, 0x7E),
            RGBColor(0x00, 0x96, 0xC7),
            RGBColor(0xB3, 0x88, 0xFF),
        ],
        border_style="neon",
        deco_style="neon",
    ),
}


def get_theme(style_id: str) -> dict:
    return THEMES.get(str(style_id), THEMES["1"])


def get_style_name(style_id: str) -> str:
    cat = STYLE_CATEGORIES.get(str(style_id), STYLE_CATEGORIES["1"])
    return cat["name"]
