# Offitrans

<div align="center">

一个强大的Office文件翻译工具库，支持PDF、Excel、PPT和Word文档的批量翻译。

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/your-username/Offitrans.svg)](https://github.com/your-username/Offitrans/stargazers)

</div>

## ✨ 特性

- 🔄 **多格式支持**：支持PDF、Excel、PPT、Word文档翻译
- 🌍 **多语言翻译**：支持中文、英文、泰文、日文、韩文、法文、德文、西班牙文等
- 🎨 **格式保持**：翻译后保持原有格式、样式和布局
- 🖼️ **图片保护**：完整保护文档中的图片不变形
- ⚡ **批量处理**：高效的批量翻译，支持文本去重和API调用优化
- 🔧 **易于集成**：简洁的API设计，方便集成到其他项目中
- 📊 **富文本支持**：支持复杂的富文本格式和合并单元格
- 🛡️ **错误处理**：完善的错误处理和重试机制

## 🚀 安装

```bash
pip install -r requirements.txt
```

或者从源码安装：

```bash
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans
pip install -e .
```

## 📦 依赖

- `openpyxl` - Excel文件处理
- `python-docx` - Word文档处理  
- `PyPDF2` - PDF文件处理
- `python-pptx` - PowerPoint文件处理
- `requests` - HTTP请求
- `Pillow` - 图片处理

## 🎯 快速开始

### Excel文件翻译

```python
from excel_translate import ExcelTranslatorV2

# 创建翻译器实例
translator = ExcelTranslatorV2()

# 翻译Excel文件
success = translator.replace_text_in_excel(
    excel_path="input.xlsx",
    output_path="output_translated.xlsx", 
    target_language="th"  # 翻译为泰文
)

if success:
    print("翻译完成！")
```

### PDF文件翻译

```python
from pdf_translate import translate_pdf

# 翻译PDF文件
translate_pdf(
    input_path="document.pdf",
    output_path="document_translated.pdf",
    target_language="en"
)
```

### Word文档翻译

```python
from word_translate import docx_translate

# 翻译Word文档
docx_translate(
    input_file="document.docx",
    output_file="document_translated.docx",
    target_language="th"
)
```

### PPT文件翻译

```python
from ppt_translate import translate_ppt

# 翻译PPT文件
translate_ppt(
    input_path="presentation.pptx", 
    output_path="presentation_translated.pptx",
    target_language="ja"
)
```

## 🔧 配置

### Google翻译API配置

项目支持Google翻译API，需要配置API密钥：

```python
from translate_tools import GoogleTranslator

translator = GoogleTranslator(
    source_lang="zh",
    target_lang="en", 
    api_key="your-google-translate-api-key"
)
```

### 支持的语言代码

| 语言 | 代码 |
|------|------|
| 中文 | zh |
| 英文 | en |
| 泰文 | th |
| 日文 | ja |
| 韩文 | ko |
| 法文 | fr |
| 德文 | de |
| 西班牙文 | es |

## 📖 详细文档

### Excel翻译器高级功能

```python
from excel_translate import ExcelTranslatorV2

# 创建带自定义配置的翻译器
translator = ExcelTranslatorV2(
    translate_api_key="your-api-key",
    font_size_adjustment=0.8  # 字体大小调整比例
)

# 分析Excel文件结构
analysis = translator.analyze_excel_structure("input.xlsx")

# 智能调整列宽
translator.smart_adjust_column_width("output.xlsx")
```

### 批量翻译工具

```python
from translate_tools import GoogleTranslator

translator = GoogleTranslator(
    source_lang="zh",
    target_lang="th",
    max_workers=5  # 并发线程数
)

# 批量翻译文本
texts = ["你好", "世界", "翻译"]
translated = translator.translate_text_batch(texts)
```

## 🤝 贡献

我们欢迎任何形式的贡献！请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解如何参与项目开发。

### 开发环境设置

```bash
# 克隆仓库
git clone https://github.com/your-username/Offitrans.git
cd Offitrans

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 运行测试
python -m pytest tests/
```

## 📝 许可证

本项目采用 [MIT许可证](LICENSE)。

## 🔗 相关链接

- [GitHub仓库](https://github.com/your-username/Offitrans)
- [问题反馈](https://github.com/your-username/Offitrans/issues)
- [版本发布](https://github.com/your-username/Offitrans/releases)

## 📊 项目状态

- ✅ Excel翻译 - 完全支持
- ✅ Word翻译 - 基础支持  
- ✅ PDF翻译 - 基础支持
- ✅ PPT翻译 - 基础支持
- 🔄 OCR支持 - 开发中
- 🔄 更多翻译引擎 - 计划中

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户！

---

<div align="center">
如果这个项目对您有帮助，请给我们一个 ⭐️
</div>