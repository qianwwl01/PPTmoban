# 🎨 PPT 模板大师

一键生成专业级 PPT 模板的 Web 应用，支持自定义配色、多种商务版式、Logo 水印等功能。

![Python](https://img.shields.io/badge/Python-3.8+-blue)
![Streamlit](https://img.shields.io/badge/Streamlit-1.28+-red)

## ✨ 功能特性

- **6种预设主题**：商务简约、科技风格、教育培训、极简白色、活力橙色、优雅紫色
- **9种版式类型**：标题页、目录页、内容页、图文页、对比页、时间轴页、数据概览页、引用页、致谢页
- **自定义配色**：主色、辅色、强调色、背景色自由调整
- **Logo 上传**：自动添加到所有页面右下角
- **图片库**：上传图片自动填充到图文页
- **水印功能**：支持自定义水印文字和透明度
- **页脚设置**：自定义页脚文字和页码显示
- **配置导入导出**：JSON 格式保存/加载配置

## 🚀 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行应用

```bash
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`

## 📁 项目结构

```
PPTmoban/
├── app.py              # Streamlit 主应用
├── ppt_generator.py    # PPT 生成逻辑
├── config_presets.py   # 预设配置
├── requirements.txt    # 依赖库
└── README.md           # 说明文档
```

## 🛠️ 技术栈

- **前端框架**：Streamlit
- **PPT 生成**：python-pptx
- **图片处理**：Pillow

## 📝 使用说明

1. 左侧边栏选择主题或自定义颜色/字体
2. 「版式配置」选择需要的页面类型和数量
3. 「导出文件」点击生成并下载

## 📄 License

MIT License
