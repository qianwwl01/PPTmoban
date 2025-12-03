# -*- coding: utf-8 -*-
"""
PPT模板预设配置
包含：预设主题风格、默认颜色、字体配置等
"""

# 预设主题风格
THEME_PRESETS = {
    "商务简约": {
        "name": "商务简约",
        "primary": "#1a365d",      # 深蓝
        "secondary": "#4a5568",    # 灰色
        "accent": "#3182ce",       # 亮蓝
        "background": "#ffffff",   # 白色背景
        "title_font": "Microsoft YaHei",
        "body_font": "Microsoft YaHei",
        "description": "深蓝配灰，专业稳重，适合商务汇报"
    },
    "科技风格": {
        "name": "科技风格",
        "primary": "#0d1117",      # 深色背景
        "secondary": "#21262d",    # 次级深色
        "accent": "#58a6ff",       # 高亮蓝
        "background": "#0d1117",   # 深色背景
        "title_font": "Microsoft YaHei",
        "body_font": "Microsoft YaHei",
        "description": "深色背景配亮色高亮，科技感十足"
    },
    "教育培训": {
        "name": "教育培训",
        "primary": "#065f46",      # 深绿
        "secondary": "#0891b2",    # 青色
        "accent": "#10b981",       # 亮绿
        "background": "#f0fdfa",   # 浅青背景
        "title_font": "Microsoft YaHei",
        "body_font": "SimSun",
        "description": "蓝绿清爽，适合教育培训场景"
    },
    "极简白色": {
        "name": "极简白色",
        "primary": "#1f2937",      # 深灰
        "secondary": "#6b7280",    # 中灰
        "accent": "#111827",       # 黑色
        "background": "#ffffff",   # 纯白
        "title_font": "Microsoft YaHei",
        "body_font": "Microsoft YaHei",
        "description": "黑白灰极简风格，干净大方"
    },
    "活力橙色": {
        "name": "活力橙色",
        "primary": "#ea580c",      # 橙色
        "secondary": "#f97316",    # 亮橙
        "accent": "#fed7aa",       # 浅橙
        "background": "#fffbeb",   # 米白
        "title_font": "Microsoft YaHei",
        "body_font": "Microsoft YaHei",
        "description": "充满活力的橙色系，适合创意展示"
    },
    "优雅紫色": {
        "name": "优雅紫色",
        "primary": "#6b21a8",      # 深紫
        "secondary": "#a855f7",    # 亮紫
        "accent": "#e9d5ff",       # 浅紫
        "background": "#faf5ff",   # 淡紫背景
        "title_font": "Microsoft YaHei",
        "body_font": "Microsoft YaHei",
        "description": "优雅大气的紫色系，适合高端展示"
    }
}

# 可选字体列表
AVAILABLE_FONTS = {
    "title": [
        "Microsoft YaHei",
        "SimHei",
        "Arial",
        "Calibri",
        "Times New Roman"
    ],
    "body": [
        "Microsoft YaHei",
        "SimSun",
        "SimHei",
        "Calibri",
        "Arial"
    ]
}

# 画布比例配置（单位：英寸）
SLIDE_RATIOS = {
    "16:9": {
        "width": 13.333,
        "height": 7.5
    },
    "4:3": {
        "width": 10.0,
        "height": 7.5
    }
}

# 版式类型配置
LAYOUT_TYPES = {
    "title": {
        "name": "标题页",
        "description": "大标题 + 副标题 + 底部信息",
        "default_count": 1
    },
    "agenda": {
        "name": "目录页",
        "description": "标题 + 多条目录条目",
        "default_count": 1
    },
    "content": {
        "name": "内容页",
        "description": "标题 + 正文（多级项目符号）",
        "default_count": 2
    },
    "image_text": {
        "name": "图文页",
        "description": "图片 + 文字混合布局",
        "default_count": 2
    },
    "comparison": {
        "name": "对比页",
        "description": "左右两列对比展示",
        "default_count": 1
    },
    "timeline": {
        "name": "时间轴页",
        "description": "水平时间线 + 节点里程碑",
        "default_count": 1
    },
    "kpi": {
        "name": "数据概览页",
        "description": "多个KPI数字 + 描述说明",
        "default_count": 1
    },
    "quote": {
        "name": "引用页",
        "description": "大号引言文字 + 作者",
        "default_count": 1
    },
    "thankyou": {
        "name": "致谢页",
        "description": "中央大标题 + 简短文字",
        "default_count": 1
    }
}

# 默认配置
DEFAULT_CONFIG = {
    "template_name": "我的PPT模板",
    "ratio": "16:9",
    "theme": "商务简约",
    "primary": "#1a365d",
    "secondary": "#4a5568",
    "accent": "#3182ce",
    "background": "#ffffff",
    "title_font": "Microsoft YaHei",
    "body_font": "Microsoft YaHei",
    "title_size": 32,
    "body_size": 18,
    "show_page_number": True,
    "footer_text": "公司名称 | 保密",
    "watermark_enabled": False,
    "watermark_text": "内部资料",
    "watermark_opacity": 15,
    "layouts": {
        "title": {"enabled": True, "count": 1},
        "agenda": {"enabled": True, "count": 1},
        "content": {"enabled": True, "count": 2},
        "image_text": {"enabled": True, "count": 2},
        "comparison": {"enabled": True, "count": 1},
        "timeline": {"enabled": True, "count": 1},
        "kpi": {"enabled": True, "count": 1},
        "quote": {"enabled": True, "count": 1},
        "thankyou": {"enabled": True, "count": 1}
    }
}
