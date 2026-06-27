"""智能PPT生成助手 - Streamlit Web应用"""

import streamlit as st
from io import BytesIO
import time

from styles import STYLE_CATEGORIES, LAYOUT_NAMES, get_style_name, get_theme
from parser import parse_markdown, parse_plain_text, extract_keywords, SlideContent
from generator import create_presentation, export_to_bytes
from image_search import UnsplashSearcher

# ── 页面配置 ─────────────────────────────────────────────────
st.set_page_config(
    page_title="智能PPT生成助手",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 自定义CSS ────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2.5rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        font-size: 1.1rem;
        text-align: center;
        color: #888;
        margin-bottom: 2rem;
    }
    .slide-preview {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 6px 0;
        border-left: 4px solid #4a90d9;
    }
    .slide-title {
        font-weight: 600;
        font-size: 1.05rem;
        color: #333;
    }
    .slide-body {
        color: #666;
        font-size: 0.9rem;
        margin-top: 4px;
    }
    .style-card {
        padding: 8px 12px;
        border-radius: 8px;
        border: 2px solid #e0e0e0;
        margin: 4px 0;
        cursor: pointer;
        transition: all 0.2s;
    }
    .style-card:hover {
        border-color: #4a90d9;
    }
    .style-card.selected {
        border-color: #4a90d9;
        background: #f0f7ff;
    }
</style>
""", unsafe_allow_html=True)

# ── 示例大纲 ─────────────────────────────────────────────────
EXAMPLE_MARKDOWN = """# 人工智能技术发展报告

探索AI的无限可能

## 技术概述

人工智能（Artificial Intelligence）是计算机科学的一个重要分支，致力于研究和开发用于模拟、延伸和扩展人类智能的理论、方法、技术及应用系统。

- 机器学习：从数据中自动学习规律
- 深度学习：多层神经网络的特征学习
- 自然语言处理：理解和生成人类语言
- 计算机视觉：让机器看懂世界

## 发展历程

### 早期探索（1950-1980）

- 1950年 图灵测试提出
- 1956年 达特茅斯会议，AI正式诞生
- 1960年代 专家系统初步发展
- 1970-80年代 AI寒冬与反思

### 快速发展期（2000-至今）

- 2012年 深度学习在ImageNet取得突破
- 2016年 AlphaGo战胜李世石
- 2022年 ChatGPT引发大模型浪潮
- 2024年 多模态大模型全面爆发

## 应用领域

### 医疗健康

- 辅助诊断与影像识别
- 药物研发加速
- 个性化治疗方案

### 智能制造

- 质量检测自动化
- 预测性维护
- 供应链优化

## 未来展望

人工智能将继续深刻改变人类社会：

- 通用人工智能的探索
- 人机协作新模式
- 伦理与安全的平衡
- 可持续发展的AI
"""

EXAMPLE_TEXT = """人工智能技术发展报告
探索AI的无限可能

技术概述
人工智能是计算机科学的重要分支，致力于模拟和扩展人类智能。

- 机器学习：从数据中自动学习规律
- 深度学习：多层神经网络的特征学习
- 自然语言处理：理解和生成人类语言
- 计算机视觉：让机器看懂世界

发展历程
从1956年达特茅斯会议AI正式诞生，到2022年ChatGPT引发大模型浪潮，AI经历了跌宕起伏的发展之路。

应用领域
AI已深入医疗、制造、教育、金融等多个领域，正在重塑产业格局。

未来展望
通用人工智能的探索、人机协作新模式、伦理与安全的平衡，是AI未来发展的核心议题。"""


# ── 主界面 ───────────────────────────────────────────────────

# 标题
st.markdown('<div class="main-title">🎬 智能PPT生成助手</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">输入大纲，一键生成精美幻灯片</div>', unsafe_allow_html=True)

# ── 侧边栏：设置 ────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 生成设置")

    # 风格选择
    st.subheader("选择风格")
    style_id = st.radio(
        "模板风格",
        options=list(STYLE_CATEGORIES.keys()),
        format_func=lambda x: f"{STYLE_CATEGORIES[x]['name']}  —  {STYLE_CATEGORIES[x]['desc']}",
        index=0,
        horizontal=False,
    )

    # 显示风格预览色
    theme = get_theme(style_id)
    cols = st.columns(5)
    color_names = ["背景", "标题", "正文", "强调", "辅助"]
    color_values = [
        theme["bg_color"], theme["title_color"], theme["body_color"],
        theme["accent_color"], theme["accent2_color"]
    ]
    for i, (name, color) in enumerate(zip(color_names, color_values)):
        with cols[i]:
            hex_color = f"#{color.red:02x}{color.green:02x}{color.blue:02x}"
            st.markdown(
                f'<div style="width:36px;height:36px;border-radius:6px;'
                f'background:{hex_color};margin:0 auto;border:1px solid #ddd;"></div>'
                f'<div style="text-align:center;font-size:10px;margin-top:2px;">{name}</div>',
                unsafe_allow_html=True
            )

    st.divider()

    # 图片设置
    st.subheader("配图设置")
    enable_images = st.checkbox("启用Unsplash配图", value=False,
                                 help="需要Unsplash API Key")
    unsplash_key = st.text_input(
        "Unsplash Access Key",
        type="password",
        value="",
        help="在 unsplash.com/developers 免费申请",
        disabled=not enable_images,
    )

    st.divider()

    # 主题色自定义
    st.subheader("自定义主题色")
    custom_accent = st.color_picker("自定义强调色", value=None,
                                     help="留空则使用风格默认色")

    st.divider()
    st.caption("智能PPT生成助手 v1.0")
    st.caption("基于 python-pptx + Streamlit 构建")

# ── 主区域：输入与生成 ───────────────────────────────────────

# 输入模式选择
input_mode = st.radio(
    "输入方式",
    options=["markdown", "text"],
    format_func=lambda x: "📝 Markdown大纲" if x == "markdown" else "📄 粘贴文本",
    horizontal=True,
)

# 文本输入区
if input_mode == "markdown":
    col1, col2 = st.columns([3, 1])
    with col1:
        user_input = st.text_area(
            "请输入Markdown大纲",
            value=EXAMPLE_MARKDOWN,
            height=400,
            placeholder="支持 # ## ### 标题层级\n- 列表项\n普通段落",
        )
    with col2:
        st.markdown("**格式说明**")
        st.caption("# 一级标题 → 封面页")
        st.caption("## 二级标题 → 章节/内容页")
        st.caption("### 三级标题 → 内容页")
        st.caption("- 列表项 → 幻灯片要点")
        st.caption("普通文本 → 正文段落")
else:
    user_input = st.text_area(
        "请粘贴文本内容",
        value=EXAMPLE_TEXT,
        height=400,
        placeholder="输入文字内容，空行分隔不同页面...",
    )

# 解析按钮
if st.button("🔍 解析大纲", type="primary", use_container_width=True):
    if not user_input.strip():
        st.warning("请先输入内容！")
    else:
        with st.spinner("正在解析..."):
            if input_mode == "markdown":
                slides = parse_markdown(user_input)
            else:
                slides = parse_plain_text(user_input)

            # 提取关键词
            slides = extract_keywords(slides)

            # 存入session
            st.session_state["slides"] = slides
            st.session_state["parsed"] = True

# ── 解析结果预览 ─────────────────────────────────────────────
if st.session_state.get("parsed") and st.session_state.get("slides"):
    slides = st.session_state["slides"]

    st.subheader(f"📑 解析结果：共 {len(slides)} 页")

    # 显示幻灯片预览
    for i, slide in enumerate(slides):
        layout_name = LAYOUT_NAMES.get(slide.layout, slide.layout)
        with st.expander(f"第 {i+1} 页 — {layout_name}：{slide.title}", expanded=(i < 3)):
            col_a, col_b = st.columns([2, 1])
            with col_a:
                st.markdown(f"**标题**：{slide.title}")
                if slide.subtitle:
                    st.markdown(f"**副标题**：{slide.subtitle}")
                if slide.bullet_items:
                    st.markdown("**要点**：")
                    for item in slide.bullet_items:
                        st.markdown(f"  • {item}")
                if slide.body_lines:
                    st.markdown("**正文**：")
                    for line in slide.body_lines:
                        st.markdown(f"  {line}")
            with col_b:
                st.markdown(f"**布局**：{layout_name}")
                st.markdown(f"**层级**：H{slide.level + 1}")
                if slide.keywords:
                    st.markdown(f"**关键词**：{'、'.join(slide.keywords[:3])}")
                if slide.image_query:
                    st.markdown(f"**配图词**：{slide.image_query}")

    # ── 生成PPT ──────────────────────────────────────────────
    st.divider()

    if st.button("🚀 生成PPTX文件", type="primary", use_container_width=True):
        with st.spinner("正在生成PPT，请稍候..."):
            try:
                prs = create_presentation(
                    slides=slides,
                    style_id=style_id,
                    unsplash_key=unsplash_key if enable_images else "",
                    enable_images=enable_images,
                )

                # 导出为字节流
                pptx_bytes = export_to_bytes(prs)

                # 生成文件名
                style_name = get_style_name(style_id).replace(" ", "_")
                filename = f"智能PPT_{style_name}_{len(slides)}页.pptx"

                st.session_state["pptx_bytes"] = pptx_bytes
                st.session_state["filename"] = filename
                st.session_state["generated"] = True

                st.success(f"PPT生成成功！共 {len(slides)} 页幻灯片")

            except Exception as e:
                st.error(f"生成失败：{e}")
                import traceback
                st.code(traceback.format_exc())

    # 下载按钮
    if st.session_state.get("generated"):
        pptx_bytes = st.session_state["pptx_bytes"]
        filename = st.session_state["filename"]

        st.download_button(
            label="📥 下载PPTX文件",
            data=pptx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

        st.info("下载后可用 PowerPoint / WPS / LibreOffice 打开编辑")

# ── 初始化session_state ──────────────────────────────────────
if "parsed" not in st.session_state:
    st.session_state["parsed"] = False
if "slides" not in st.session_state:
    st.session_state["slides"] = []
if "generated" not in st.session_state:
    st.session_state["generated"] = False
