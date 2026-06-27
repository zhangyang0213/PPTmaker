"""5种风格主题定义 - 配色、字体、布局参数"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 幻灯片尺寸 (16:9, 单位: EMU)
SLIDE_WIDTH = 12192000   # ~33.02cm
SLIDE_HEIGHT = 6858000   # ~19.05cm

# 布局模板类型
LAYOUT_TITLE = "title"           # 封面页
LAYOUT_TOC = "toc"               # 目录页
LAYOUT_SECTION = "section"       # 章节页
LAYOUT_CONTENT = "content"       # 文字内容页
LAYOUT_IMAGE_TEXT = "image_text" # 图文页
LAYOUT_FULL_IMAGE = "full_image" # 全图页
LAYOUT_TWO_COL = "two_col"      # 双栏页
LAYOUT_END = "end"              # 结束页

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
}

def _theme(
    bg_color, title_color, body_color, accent_color, accent2_color,
    title_font, body_font, title_size, body_size, section_size,
    decorative_colors, border_style
):
    """构造主题字典"""
    return {
        "bg_color": bg_color,
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
    }

# 5种风格主题详细配置
THEMES = {
    "1": _theme(
        bg_color=RGBColor(0xF5, 0xF2, 0xEB),          # 宣纸色
        title_color=RGBColor(0x2C, 0x2C, 0x2C),        # 墨黑
        body_color=RGBColor(0x4A, 0x4A, 0x4A),         # 深灰
        accent_color=RGBColor(0x3B, 0x5E, 0x7A),       # 墨蓝
        accent2_color=RGBColor(0x8B, 0x73, 0x55),      # 赭石
        title_font="华文中宋",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xC8, 0xC0, 0xAD),  # 淡墨
            RGBColor(0x9E, 0x94, 0x84),  # 中墨
            RGBColor(0x3B, 0x5E, 0x7A),  # 墨蓝
        ],
        border_style="ink",  # 水墨边框
    ),
    "2": _theme(
        bg_color=RGBColor(0xF8, 0xFC, 0xF5),          # 薄荷白
        title_color=RGBColor(0x2D, 0x5A, 0x27),        # 深绿
        body_color=RGBColor(0x3D, 0x5A, 0x3D),         # 暗绿
        accent_color=RGBColor(0x5B, 0xA5, 0x5B),       # 草绿
        accent2_color=RGBColor(0x8B, 0xC3, 0x4A),      # 嫩绿
        title_font="等线",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xB8, 0xDB, 0xA8),  # 浅绿
            RGBColor(0x5B, 0xA5, 0x5B),  # 草绿
            RGBColor(0x8B, 0xC3, 0x4A),  # 嫩绿
        ],
        border_style="organic",  # 有机曲线
    ),
    "3": _theme(
        bg_color=RGBColor(0xFD, 0xF5, 0xE6),          # 米黄
        title_color=RGBColor(0x8B, 0x00, 0x00),        # 暗红
        body_color=RGBColor(0x4A, 0x2C, 0x0A),         # 深棕
        accent_color=RGBColor(0xC4, 0x1E, 0x1E),       # 中国红
        accent2_color=RGBColor(0xD4, 0xA0, 0x17),      # 金色
        title_font="华文行楷",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xC4, 0x1E, 0x1E),  # 中国红
            RGBColor(0xD4, 0xA0, 0x17),  # 金色
            RGBColor(0xF5, 0xE6, 0xC8),  # 浅金
        ],
        border_style="chinese",  # 中式边框
    ),
    "4": _theme(
        bg_color=RGBColor(0x1A, 0x1A, 0x1A),           # 纯黑
        title_color=RGBColor(0xFF, 0xFF, 0xFF),         # 白色
        body_color=RGBColor(0xE0, 0xE0, 0xE0),         # 浅灰
        accent_color=RGBColor(0xE6, 0x2C, 0x2C),       # 革命红
        accent2_color=RGBColor(0xFF, 0xD7, 0x00),      # 金黄
        title_font="黑体",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(48),
        decorative_colors=[
            RGBColor(0xE6, 0x2C, 0x2C),  # 革命红
            RGBColor(0xFF, 0xD7, 0x00),  # 金黄
            RGBColor(0x4A, 0x4A, 0x4A),  # 深灰
        ],
        border_style="bold",  # 粗犷边框
    ),
    "5": _theme(
        bg_color=RGBColor(0xF5, 0xEB, 0xD6),          # 木色
        title_color=RGBColor(0x3E, 0x27, 0x23),        # 深棕
        body_color=RGBColor(0x5C, 0x3A, 0x21),         # 赭棕
        accent_color=RGBColor(0x8B, 0x5A, 0x2B),       # 木棕
        accent2_color=RGBColor(0xA0, 0x7E, 0x4F),      # 浅木色
        title_font="华文中宋",
        body_font="微软雅黑",
        title_size=Pt(36),
        body_size=Pt(18),
        section_size=Pt(44),
        decorative_colors=[
            RGBColor(0xD2, 0xB4, 0x8C),  # 沙棕
            RGBColor(0x8B, 0x5A, 0x2B),  # 木棕
            RGBColor(0xA0, 0x7E, 0x4F),  # 浅木色
        ],
        border_style="wood",  # 木纹边框
    ),
}


def get_theme(style_id: str) -> dict:
    """获取指定风格的主题配置"""
    return THEMES.get(str(style_id), THEMES["1"])


def get_style_name(style_id: str) -> str:
    """获取风格名称"""
    cat = STYLE_CATEGORIES.get(str(style_id), STYLE_CATEGORIES["1"])
    return cat["name"]
