# Markdown to Word Converter

这是一个基于 Python 的轻量级脚本，专门用于将 **Markdown** 转换为符合中文排版规范（宋体/黑体 + Times New Roman）的 **Word (.docx)** 文档。

### 🌟 核心功能
* **中英双字体**：中文使用宋体/黑体，英文/数字自动切换为 Times New Roman。
* **物理公式支持**：深度兼容 LaTeX 数学公式（基于 Pandoc 的 `--mathjax`）。
* **精美排版**：标题 1.5 倍行距，正文 1.25 倍行距。
* **批量处理**：支持单个文件转换或整个文件夹一键批量处理。

### 🛠️ 环境准备
你需要安装 Python 3.x 以及以下库：

```bash
pip install pypandoc python-docx
```
> **注意**：系统需预装 [Pandoc](https://pandoc.org/installing.html)。

### 🚀 快速开始
1. 运行脚本：
   ```bash
   python mdtoword_paperformat_upload.py
   ```
2. 根据提示输入 **1** (单个文件) 或 **2** (文件夹)。
3. 将你的 `.md` 文件或文件夹拖入终端并按回车。

### 📝 格式细节
* **一级标题**：黑体, 22pt, 居中, 加粗
* **二级标题**：宋体, 15pt, 加粗
* **三级标题**：宋体, 12pt, 加粗
* **正文**：宋体, 10.5pt (五号)


# CHANGELOG
2026.4.21：新增上标内容{{SUP_START}}内容{{SUP_END}}
2026.4.22：新增给代码块一个浅灰色框,支持代码高亮
