# -*- coding: utf-8 -*-
"""
PPT模板制作工具 - Streamlit Web应用
一键生成精美的PPT模板文件
"""

import streamlit as st
import json
from datetime import datetime

from config_presets import (
    THEME_PRESETS, 
    AVAILABLE_FONTS, 
    LAYOUT_TYPES, 
    DEFAULT_CONFIG
)
from ppt_generator import build_presentation


# ==================== 页面配置 ====================
st.set_page_config(
    page_title="PPT模板制作工具",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义CSS样式 ====================
st.markdown("""
<style>
    /* 全局字体优化 */
    html, body, [class*="css"] {
        font-family: 'Inter', 'Microsoft YaHei', sans-serif;
    }
    
    /* 隐藏默认的汉堡菜单和Footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 顶部 Header 优化 */
    .main-header {
        background: linear-gradient(135deg, #1a365d 0%, #2563eb 100%);
        color: white;
        padding: 2rem;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 {
        color: white !important;
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
    }
    .main-header p {
        font-size: 1.1rem;
        opacity: 0.9;
    }

    /* 卡片样式通用类 */
    .stCard {
        background-color: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
        border: 1px solid #e2e8f0;
        margin-bottom: 1rem;
        transition: all 0.2s ease;
    }
    .stCard:hover {
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        transform: translateY(-2px);
    }
    
    /* 颜色预览卡片 */
    .color-card {
        padding: 1rem;
        border-radius: 12px;
        text-align: center;
        color: white;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .color-card span {
        display: block;
    }
    .color-name {
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 4px;
    }
    .color-hex {
        font-family: monospace;
        font-size: 0.9rem;
        opacity: 0.9;
        background: rgba(0,0,0,0.1);
        padding: 2px 6px;
        border-radius: 4px;
    }
    
    /* 版式卡片 */
    .layout-card-container {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem;
        height: 100%;
        transition: all 0.2s;
    }
    .layout-card-container:hover {
        border-color: #3182ce;
        box-shadow: 0 0 0 3px rgba(49, 130, 206, 0.1);
    }
    .layout-title {
        color: #1a365d;
        font-weight: 700;
        font-size: 1.1rem;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .layout-desc {
        color: #64748b;
        font-size: 0.9rem;
        line-height: 1.5;
        margin-bottom: 1rem;
        height: 40px; /* 固定高度保持对齐 */
    }
    
    /* 预览幻灯片 */
    .slide-preview {
        aspect-ratio: 16/9;
        border-radius: 8px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
        background-color: white;
    }
    
    /* Tabs 样式优化 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
        background-color: transparent;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 0;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border: none;
        color: #64748b;
        font-weight: 600;
        padding: 0 4px;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #1a365d;
    }
    .stTabs [aria-selected="true"] {
        color: #1a365d !important;
        border-bottom: 3px solid #1a365d !important;
    }
    
    /* 侧边栏优化 */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    section[data-testid="stSidebar"] h2 {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1e293b;
    }
</style>
""", unsafe_allow_html=True)


# ==================== 初始化会话状态 ====================
def init_session_state():
    """初始化Streamlit会话状态"""
    if 'config' not in st.session_state:
        st.session_state.config = DEFAULT_CONFIG.copy()
    if 'generated' not in st.session_state:
        st.session_state.generated = False
    if 'ppt_buffer' not in st.session_state:
        st.session_state.ppt_buffer = None
    if 'logo_bytes' not in st.session_state:
        st.session_state.logo_bytes = None
    if 'uploaded_images' not in st.session_state:
        st.session_state.uploaded_images = []


init_session_state()


# ==================== 侧边栏 - 全局设置 ====================
def render_sidebar():
    """渲染侧边栏的全局设置"""
    with st.sidebar:
        st.markdown("## 🛠️ 全局配置")
        
        # 1. 基础信息
        with st.expander("📝 基础信息", expanded=True):
            st.session_state.config['template_name'] = st.text_input(
                "模板名称",
                value=st.session_state.config.get('template_name', '我的PPT模板')
            )
            st.session_state.config['ratio'] = st.radio(
                "画布比例",
                options=['16:9', '4:3'],
                index=0 if st.session_state.config.get('ratio', '16:9') == '16:9' else 1,
                horizontal=True
            )

        # 2. 主题风格
        with st.expander("🎨 主题风格", expanded=True):
            theme_names = list(THEME_PRESETS.keys())
            selected_theme = st.selectbox(
                "选择预设主题",
                options=theme_names,
                index=theme_names.index(st.session_state.config.get('theme', '商务简约'))
            )
            
            if st.button("应用主题预设", use_container_width=True, type="secondary"):
                theme = THEME_PRESETS[selected_theme]
                st.session_state.config.update({
                    'theme': selected_theme,
                    'primary': theme['primary'],
                    'secondary': theme['secondary'],
                    'accent': theme['accent'],
                    'background': theme['background'],
                    'title_font': theme['title_font'],
                    'body_font': theme['body_font']
                })
                st.rerun()
            
            if selected_theme in THEME_PRESETS:
                st.caption(f"💡 {THEME_PRESETS[selected_theme]['description']}")

        # 3. 自定义配色
        with st.expander("🖌️ 自定义配色", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                st.session_state.config['primary'] = st.color_picker("主色", value=st.session_state.config.get('primary', '#1a365d'))
                st.session_state.config['accent'] = st.color_picker("强调色", value=st.session_state.config.get('accent', '#3182ce'))
            with c2:
                st.session_state.config['secondary'] = st.color_picker("辅色", value=st.session_state.config.get('secondary', '#4a5568'))
                st.session_state.config['background'] = st.color_picker("背景色", value=st.session_state.config.get('background', '#ffffff'))

        # 4. 字体设置
        with st.expander("Aa 字体设置", expanded=False):
            st.session_state.config['title_font'] = st.selectbox(
                "标题字体",
                options=AVAILABLE_FONTS['title'],
                index=AVAILABLE_FONTS['title'].index(st.session_state.config.get('title_font', 'Microsoft YaHei')) if st.session_state.config.get('title_font') in AVAILABLE_FONTS['title'] else 0
            )
            st.session_state.config['body_font'] = st.selectbox(
                "正文字体",
                options=AVAILABLE_FONTS['body'],
                index=AVAILABLE_FONTS['body'].index(st.session_state.config.get('body_font', 'Microsoft YaHei')) if st.session_state.config.get('body_font') in AVAILABLE_FONTS['body'] else 0
            )

        # 5. 资源库 (Logo & 图片)
        with st.expander("📂 资源库", expanded=False):
            st.markdown("**Logo 上传**")
            uploaded_logo = st.file_uploader("上传Logo (PNG/JPG)", type=['png', 'jpg', 'jpeg'], key="logo_uploader")
            if uploaded_logo:
                st.session_state.logo_bytes = uploaded_logo.read()
                st.image(uploaded_logo, width=80, caption="Logo预览")
            
            if st.session_state.logo_bytes:
                if st.button("🗑️ 清除Logo", use_container_width=True):
                    st.session_state.logo_bytes = None
                    st.rerun()
            
            st.divider()
            
            st.markdown("**图文页图片**")
            uploaded_images = st.file_uploader("上传图片 (多选)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True, key="img_uploader")
            if uploaded_images:
                st.session_state.uploaded_images = [{'name': img.name, 'bytes': img.read()} for img in uploaded_images]
                st.success(f"已加载 {len(uploaded_images)} 张图片")
            
            if st.session_state.uploaded_images:
                if st.button("🗑️ 清除图片库", use_container_width=True):
                    st.session_state.uploaded_images = []
                    st.rerun()

        # 6. 页脚与水印
        with st.expander("📑 页脚与水印", expanded=False):
            st.markdown("**水印**")
            watermark_on = st.toggle("启用水印", value=st.session_state.config.get('watermark_enabled', False))
            st.session_state.config['watermark_enabled'] = watermark_on
            if watermark_on:
                st.session_state.config['watermark_text'] = st.text_input("水印内容", value=st.session_state.config.get('watermark_text', '内部资料'))
                st.session_state.config['watermark_opacity'] = st.slider("透明度", 5, 50, st.session_state.config.get('watermark_opacity', 15))
            
            st.markdown("**页脚**")
            page_num_on = st.toggle("显示页码", value=st.session_state.config.get('show_page_number', True))
            st.session_state.config['show_page_number'] = page_num_on
            st.session_state.config['footer_text'] = st.text_input("页脚文字", value=st.session_state.config.get('footer_text', '公司名称'))

        st.divider()
        
        # 导出配置
        with st.expander("💾 配置管理", expanded=False):
            config_json = json.dumps(st.session_state.config, ensure_ascii=False, indent=2)
            st.download_button("📥 下载配置", data=config_json, file_name="config.json", mime="application/json", use_container_width=True)
            uploaded_config = st.file_uploader("📤 导入配置", type=['json'])
            if uploaded_config:
                try:
                    st.session_state.config.update(json.load(uploaded_config))
                    st.success("导入成功")
                    st.rerun()
                except:
                    st.error("导入失败")


# ==================== 主区域 - Tab1: 主题预览 ====================
def render_theme_preview():
    """渲染主题预览页面"""
    st.markdown("### 🎨 主题预览")
    
    config = st.session_state.config
    
    # 色彩卡片行
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="color-card" style="background:{config['primary']};">
            <span class="color-name">主色 Primary</span><span class="color-hex">{config['primary']}</span></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="color-card" style="background:{config['secondary']};">
            <span class="color-name">辅色 Secondary</span><span class="color-hex">{config['secondary']}</span></div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="color-card" style="background:{config['accent']};">
            <span class="color-name">强调色 Accent</span><span class="color-hex">{config['accent']}</span></div>""", unsafe_allow_html=True)
    with c4:
        bg_text = "#1a202c" if config['background'].lower() in ['#ffffff', '#fff', '#f8fafc'] else "#ffffff"
        st.markdown(f"""<div class="color-card" style="background:{config['background']};color:{bg_text};border:1px solid #e2e8f0;">
            <span class="color-name">背景 Background</span><span class="color-hex" style="background:rgba(0,0,0,0.05)">{config['background']}</span></div>""", unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)

    # 预览幻灯片
    st.markdown("---")
    st.markdown("**📊 幻灯片预览**")
    
    # 顶部色条
    st.markdown(f'<div style="background:{config["primary"]}; height:8px; border-radius:8px 8px 0 0;"></div>', unsafe_allow_html=True)
    
    # 预览卡片
    st.markdown(f'''
    <div style="border:1px solid #e2e8f0; border-top:none; border-radius:0 0 8px 8px; padding:1.5rem; background:{config["background"]};">
        <h2 style="color:{config["primary"]}; margin:0 0 0.5rem 0;">{config.get("template_name", "演示文稿标题")}</h2>
        <p style="color:{config["secondary"]}; opacity:0.8; margin:0;">在此输入副标题或简短描述内容</p>
        <div style="width:60px; height:4px; background:{config["accent"]}; margin:1rem 0;"></div>
        <p style="color:{config["secondary"]}; opacity:0.6; font-size:0.85rem;">汇报人姓名 | {datetime.now().year}年度汇报</p>
    </div>
    ''', unsafe_allow_html=True)
    
    # Logo 显示
    if st.session_state.logo_bytes:
        st.image(st.session_state.logo_bytes, width=80, caption="已上传Logo")


# ==================== 主区域 - Tab2: 版式设置 ====================
def render_layout_settings():
    """渲染版式设置页面"""
    st.markdown("### 📐 版式配置")
    
    if 'layouts' not in st.session_state.config:
        st.session_state.config['layouts'] = DEFAULT_CONFIG['layouts'].copy()
    
    layouts = st.session_state.config['layouts']
    
    # 使用 Grid 布局
    cols = st.columns(3)
    layout_items = list(LAYOUT_TYPES.items())
    
    for i, (layout_key, layout_info) in enumerate(layout_items):
        col = cols[i % 3]
        with col:
            # 卡片容器开始
            st.markdown(f"""
            <div class="layout-card-container">
                <div class="layout-title">
                    <span style="background:#eff6ff; padding:4px 8px; border-radius:6px; font-size:0.8rem; color:#3b82f6;">#{i+1}</span>
                    {layout_info['name']}
                </div>
                <div class="layout-desc">{layout_info['description']}</div>
            </div>
            """, unsafe_allow_html=True)
            
            # 控件区域 (放在markdown下方，利用Streamlit布局自动对齐)
            c1, c2 = st.columns([1, 1.5])
            with c1:
                enabled = st.toggle("启用", value=layouts.get(layout_key, {}).get('enabled', True), key=f"en_{layout_key}")
            with c2:
                count = st.number_input("页数", min_value=0, max_value=20, value=layouts.get(layout_key, {}).get('count', 1), key=f"cnt_{layout_key}", disabled=not enabled, label_visibility="collapsed")
            
            # 更新状态
            layouts[layout_key] = {'enabled': enabled, 'count': count}
            st.markdown("<div style='margin-bottom:12px'></div>", unsafe_allow_html=True) # Spacer



# ==================== 主区域 - Tab3: 预览与导出 ====================
def render_export():
    """渲染预览与导出页面"""
    
    config = st.session_state.config
    layouts = config.get('layouts', DEFAULT_CONFIG['layouts'])
    
    st.markdown("""
    <div style="text-align:center; padding: 40px 0;">
        <h2 style="color:#1a365d; margin-bottom:10px;">🚀 准备就绪</h2>
        <p style="color:#64748b;">确认配置无误后，点击下方按钮生成您的专属PPT模板</p>
    </div>
    """, unsafe_allow_html=True)

    # 统计幻灯片总数
    total_slides = sum(
        layouts.get(k, {}).get('count', 0) 
        for k in LAYOUT_TYPES.keys() 
        if layouts.get(k, {}).get('enabled', False)
    )
    
    # 居中布局生成按钮
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("✨ 立即生成 PPT 模板", use_container_width=True, type="primary"):
            if total_slides == 0:
                st.error("请至少启用一种版式并设置页数大于0！")
                return
            
            with st.spinner("🎨 正在绘制幻灯片..."):
                try:
                    logo_bytes = st.session_state.get('logo_bytes', None)
                    uploaded_images = st.session_state.get('uploaded_images', [])
                    ppt_buffer = build_presentation(config, layouts, logo_bytes, uploaded_images)
                    st.session_state.ppt_buffer = ppt_buffer
                    st.session_state.generated = True
                    st.balloons() # 成功动画
                except Exception as e:
                    st.error(f"生成失败: {e}")
                    return
    
    # 下载区域
    if st.session_state.generated and st.session_state.ppt_buffer:
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 成功卡片
        st.markdown(f"""
        <div class="stCard" style="background:#f0fdf4; border-color:#bbf7d0; text-align:center;">
            <h3 style="color:#166534; margin:0;">🎉 生成成功！</h3>
            <p style="color:#15803d; margin:8px 0;">共计 {total_slides} 页幻灯片，文件大小约 {len(st.session_state.ppt_buffer.getvalue())/1024:.1f} KB</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            template_name = config.get('template_name', '我的PPT模板')
            file_name = f"{template_name}_模板.pptx"
            
            st.download_button(
                label="📥 点击下载文件",
                data=st.session_state.ppt_buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )


# ==================== 主函数 ====================
def main():
    """主函数 - 应用入口"""
    
    # 渲染侧边栏
    render_sidebar()
    
    # 顶部 Banner
    st.markdown('<div class="main-header">'
                '<h1>🎨 PPT 模板大师</h1>'
                '<p>一键生成专业级演示文稿模板，支持自定义配色与多种商务版式</p>'
                '</div>', unsafe_allow_html=True)
    
    # Tab导航
    tab1, tab2, tab3 = st.tabs(["🎨 主题预览", "📐 版式配置", "📥 导出文件"])
    
    with tab1:
        render_theme_preview()
    
    with tab2:
        render_layout_settings()
    
    with tab3:
        render_export()
    
    # 页脚
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    st.markdown(
        "<div style='text-align:center;color:#cbd5e1;font-size:12px;'>"
        "Powered by Streamlit & python-pptx | Design by Cascade"
        "</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
