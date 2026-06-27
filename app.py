"""智能PPT生成助手 - Streamlit Web应用"""

import streamlit as st
from io import BytesIO
import time

from pptx import Presentation

from styles import STYLE_CATEGORIES, LAYOUT_NAMES, get_style_name, get_theme
from parser import parse_markdown, parse_plain_text, extract_keywords, SlideContent
from generator import create_presentation, export_to_bytes, generate_preview_images
from image_search import UnsplashSearcher

st.set_page_config(
    page_title="智能PPT生成助手",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-title { font-size: 2.5rem; font-weight: 700; text-align: center; margin-bottom: 0.5rem; }
    .sub-title { font-size: 1.1rem; text-align: center; color: #888; margin-bottom: 2rem; }
    .success-box { background: #d4edda; border-radius: 8px; padding: 16px; margin: 12px 0; border-left: 4px solid #28a745; }
    .info-box { background: #e7f3ff; border-radius: 8px; padding: 16px; margin: 12px 0; border-left: 4px solid #007bff; }
</style>
""", unsafe_allow_html=True)

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

# 初始化session_state
if "parsed" not in st.session_state:
    st.session_state["parsed"] = False
if "slides" not in st.session_state:
    st.session_state["slides"] = []
if "generated" not in st.session_state:
    st.session_state["generated"] = False
if "pptx_bytes" not in st.session_state:
    st.session_state["pptx_bytes"] = None
if "filename" not in st.session_state:
    st.session_state["filename"] = ""
if "template_prs" not in st.session_state:
    st.session_state["template_prs"] = None
if "use_template" not in st.session_state:
    st.session_state["use_template"] = False

# ── 主界面 ──
st.markdown('<div class="main-title">🎬 智能PPT生成助手</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">输入大纲，一键生成精美幻灯片</div>', unsafe_allow_html=True)

# ── 侧边栏 ──
with st.sidebar:
    st.header("⚙️ 生成设置")

    # 风格选择（含"导入模板"选项）
    st.subheader("选择风格")
    style_options = list(STYLE_CATEGORIES.keys()) + ["template"]
    style_labels = {k: f"{STYLE_CATEGORIES[k]['name']}  —  {STYLE_CATEGORIES[k]['desc']}" for k in STYLE_CATEGORIES}
    style_labels["template"] = "📎 导入模板  —  上传自己的PPT模板"

    style_id = st.radio(
        "风格",
        options=style_options,
        format_func=lambda x: style_labels[x],
        index=0,
        horizontal=False,
        label_visibility="collapsed",
    )

    use_template = (style_id == "template")
    st.session_state["use_template"] = use_template

    # 显示风格预览色
    if not use_template:
        theme = get_theme(style_id)
        cols = st.columns(6)
        color_names = ["背景", "渐变", "标题", "正文", "强调", "辅助"]
        color_values = [
            theme["bg_color"], theme["bg_color2"], theme["title_color"],
            theme["body_color"], theme["accent_color"], theme["accent2_color"]
        ]
        for i, (name, color) in enumerate(zip(color_names, color_values)):
            with cols[i]:
                hex_color = f"#{str(color)[0:2]}{str(color)[2:4]}{str(color)[4:6]}"
                st.markdown(
                    f'<div style="width:30px;height:30px;border-radius:6px;'
                    f'background:{hex_color};margin:0 auto;border:1px solid #ddd;"></div>'
                    f'<div style="text-align:center;font-size:9px;margin-top:2px;">{name}</div>',
                    unsafe_allow_html=True
                )

    # 模板上传
    if use_template:
        st.divider()
        st.subheader("上传PPT模板")
        template_file = st.file_uploader(
            "上传.pptx模板文件",
            type=["pptx"],
            help="上传你的PPT模板，将基于模板的布局生成新PPT",
            key="template_uploader",
        )
        if template_file is not None:
            try:
                template_bytes = template_file.read()
                st.session_state["template_prs"] = Presentation(BytesIO(template_bytes))
                tprs = st.session_state["template_prs"]
                st.success(f"模板已加载：{template_file.name}")
                st.info(f"模板包含 {len(tprs.slides)} 页幻灯片，将自动识别封面/目录/正文/结尾页")
            except Exception as e:
                st.error(f"模板加载失败：{e}")
                st.session_state["template_prs"] = None
        else:
            st.session_state["template_prs"] = None
            st.caption("请上传.pptx格式的PPT模板文件")

    st.divider()

    # 配图设置
    st.subheader("配图设置")
    enable_images = st.checkbox("启用Unsplash配图", value=True,
                                 help="根据内容关键词自动搜索配图")
    unsplash_key = st.text_input(
        "Unsplash Access Key",
        type="password",
        value="AerhhHr9KNc0kxSEzVj2q3mEdWBTh2bRmDfU9McNneI",
        help="已预置默认Key",
        disabled=not enable_images,
    )

    st.divider()
    st.caption("智能PPT生成助手 v3.0")
    st.caption("8种风格 | 模板导入 | 在线预览")

# ── 主区域：输入 ──
input_mode = st.radio(
    "输入方式",
    options=["markdown", "text"],
    format_func=lambda x: "📝 Markdown大纲" if x == "markdown" else "📄 粘贴文本",
    horizontal=True,
)

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
            slides = extract_keywords(slides)
            st.session_state["slides"] = slides
            st.session_state["parsed"] = True
            st.session_state["generated"] = False

# ── 解析结果预览 ──
if st.session_state.get("parsed") and st.session_state.get("slides"):
    slides = st.session_state["slides"]
    st.subheader(f"📑 解析结果：共 {len(slides)} 页")

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

    # ── 生成PPT ──
    st.divider()

    col_gen1, col_gen2 = st.columns([1, 1])
    with col_gen1:
        generate_btn = st.button("🚀 生成PPTX文件", type="primary", use_container_width=True)
    with col_gen2:
        if st.session_state.get("generated"):
            switch_btn = st.button("🔄 切换风格重新生成", use_container_width=True)
        else:
            switch_btn = False

    # 检查模板模式是否上传了模板
    can_generate = True
    if use_template and st.session_state.get("template_prs") is None:
        can_generate = False
        if generate_btn:
            st.warning("请先上传PPT模板文件！")

    if (generate_btn or switch_btn) and can_generate:
        with st.spinner("正在生成PPT，请稍候..."):
            try:
                # 确定模板
                template = st.session_state.get("template_prs") if use_template else None
                # 确定风格ID（模板模式用默认风格1的配色）
                sid = style_id if not use_template else "1"

                prs = create_presentation(
                    slides=slides,
                    style_id=sid,
                    unsplash_key=unsplash_key if enable_images else "",
                    enable_images=enable_images,
                    template_prs=template,
                )

                pptx_bytes = export_to_bytes(prs)
                style_name = "导入模板" if use_template else get_style_name(style_id)
                filename = f"智能PPT_{style_name}_{len(slides)}页.pptx"

                # 生成预览图
                preview_imgs = generate_preview_images(prs)

                st.session_state["pptx_bytes"] = pptx_bytes
                st.session_state["filename"] = filename
                st.session_state["generated"] = True
                st.session_state["preview_imgs"] = preview_imgs

                st.markdown(
                    f'<div class="success-box">'
                    f'✅ <b>生成成功！</b>共 {len(slides)} 页 | 风格：{style_name} | {len(pptx_bytes)//1024}KB'
                    f'<br>💡 不满意？切换风格后点击 <b>🔄 切换风格重新生成</b>'
                    f'</div>',
                    unsafe_allow_html=True
                )

            except Exception as e:
                st.error(f"生成失败：{e}")
                import traceback
                st.code(traceback.format_exc())

    # ── 预览与下载区 ──
    if st.session_state.get("generated") and st.session_state.get("pptx_bytes"):
        st.divider()
        st.subheader("👀 在线预览")

        pptx_bytes = st.session_state["pptx_bytes"]
        filename = st.session_state["filename"]
        preview_imgs = st.session_state.get("preview_imgs", [])

        # 预览幻灯片
        if preview_imgs:
            # 使用tabs展示每页
            tab_names = [f"第{i+1}页" for i in range(len(preview_imgs))]
            tabs = st.tabs(tab_names)
            for i, (tab, img_data) in enumerate(zip(tabs, preview_imgs)):
                with tab:
                    st.image(img_data, caption=f"第 {i+1} 页", use_column_width=True)
        else:
            st.info("预览图生成失败，请下载后查看")

        st.divider()
        st.subheader("📥 下载文件")

        col_dl1, col_dl2 = st.columns([2, 1])
        with col_dl1:
            st.download_button(
                label="📥 下载PPTX文件",
                data=pptx_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
            st.info("下载后可用 PowerPoint / WPS / LibreOffice 打开编辑")

        with col_dl2:
            current_style = "导入模板" if use_template else get_style_name(style_id)
            st.markdown(
                f'<div class="info-box">'
                f'<b>当前风格</b><br>{current_style}<br><br>'
                f'<b>幻灯片数</b><br>{len(slides)} 页<br><br>'
                f'<b>文件大小</b><br>{len(pptx_bytes)//1024} KB'
                f'</div>',
                unsafe_allow_html=True
            )

        # 风格切换提示
        st.markdown("---")
        st.markdown("#### 💡 快速切换风格")
        st.markdown("在左侧边栏选择新风格或导入模板，然后点击 **🔄 切换风格重新生成** 即可，无需重新输入大纲。")

        preview_cols = st.columns(4)
        for i, (sid, scat) in enumerate(STYLE_CATEGORIES.items()):
            with preview_cols[i % 4]:
                t = get_theme(sid)
                bg_hex = f"#{str(t['bg_color'])[0:2]}{str(t['bg_color'])[2:4]}{str(t['bg_color'])[4:6]}"
                accent_hex = f"#{str(t['accent_color'])[0:2]}{str(t['accent_color'])[2:4]}{str(t['accent_color'])[4:6]}"
                is_current = "✅ " if sid == style_id and not use_template else ""
                st.markdown(
                    f'<div style="padding:8px;border-radius:8px;border:2px solid '
                    f'{"#28a745" if sid == style_id and not use_template else "#e0e0e0"};'
                    f'background:{bg_hex};margin:4px 0;">'
                    f'<div style="font-size:12px;color:{accent_hex};font-weight:bold;">'
                    f'{is_current}{scat["name"]}</div>'
                    f'<div style="font-size:9px;color:#666;">{scat["desc"]}</div>'
                    f'</div>',
                    unsafe_allow_html=True
                )
