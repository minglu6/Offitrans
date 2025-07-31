# Offitrans

<div align="center">

一个强大的Office文件翻译工具库，支持PDF、Excel、PPT和Word文档的批量翻译。

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PyPI](https://img.shields.io/pypi/v/offitrans.svg)](https://pypi.org/project/offitrans/)
[![GitHub Stars](https://img.shields.io/github/stars/minglu6/Offitrans.svg)](https://github.com/minglu6/Offitrans/stargazers)

[English](README.md) | **中文**

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
- 💻 **命令行工具**：提供便捷的CLI界面，支持批量操作

## 🚀 安装

### 从 PyPI 安装（推荐）

```bash
# 基础版本
pip install offitrans

# 包含所有依赖的完整版本
pip install offitrans[full]

# 按需安装特定格式支持
pip install offitrans[excel]     # Excel支持
pip install offitrans[word]      # Word支持  
pip install offitrans[pdf]       # PDF支持
pip install offitrans[powerpoint] # PowerPoint支持
```

### 从源码安装

```bash
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans
pip install -e .
```

## 🎯 快速开始

### 命令行使用

```bash
# 翻译Excel文件
offitrans input.xlsx -t zh -o output_zh.xlsx

# 翻译PDF文件
offitrans document.pdf -t en -o document_en.pdf

# 翻译Word文档
offitrans document.docx -t th -o document_th.docx

# 翻译PowerPoint演示文稿
offitrans presentation.pptx -t ja -o presentation_ja.pptx

# 查看所有选项
offitrans --help
```

### Python API 使用

#### Excel文件翻译

```python
from offitrans.processors.excel import ExcelProcessor
from offitrans.translators.google import GoogleTranslator

# 创建翻译器和处理器
translator = GoogleTranslator()
processor = ExcelProcessor()

# 翻译Excel文件
processor.translate_file(
    input_path="input.xlsx",
    output_path="output_translated.xlsx",
    translator=translator,
    target_lang="th"  # 翻译为泰文
)
```

#### PDF文件翻译

```python
from offitrans.processors.pdf import PDFProcessor
from offitrans.translators.google import GoogleTranslator

translator = GoogleTranslator()
processor = PDFProcessor()

processor.translate_file(
    input_path="document.pdf",
    output_path="document_translated.pdf",
    translator=translator,
    target_lang="en"
)
```

#### Word文档翻译

```python
from offitrans.processors.word import WordProcessor
from offitrans.translators.google import GoogleTranslator

translator = GoogleTranslator()
processor = WordProcessor()

processor.translate_file(
    input_path="document.docx",
    output_path="document_translated.docx",
    translator=translator,
    target_lang="th"
)
```

#### PowerPoint翻译

```python
from offitrans.processors.powerpoint import PowerPointProcessor
from offitrans.translators.google import GoogleTranslator

translator = GoogleTranslator()
processor = PowerPointProcessor()

processor.translate_file(
    input_path="presentation.pptx",
    output_path="presentation_translated.pptx",
    translator=translator,
    target_lang="ja"
)
```

## 🔧 配置

### Google翻译API配置

```python
from offitrans.translators.google import GoogleTranslator

# 使用API密钥
translator = GoogleTranslator(api_key="your-google-translate-api-key")

# 或者设置环境变量
import os
os.environ['GOOGLE_TRANSLATE_API_KEY'] = 'your-api-key'
translator = GoogleTranslator()
```

### 支持的语言代码

| 语言 | 代码 | 语言 | 代码 |
|------|------|------|------|
| 中文 | zh | 英文 | en |
| 泰文 | th | 日文 | ja |
| 韩文 | ko | 法文 | fr |
| 德文 | de | 西班牙文 | es |

## 📖 高级用法

### 批量翻译

```python
from offitrans.core.utils import BatchProcessor
from offitrans.translators.google import GoogleTranslator

# 批量处理多个文件
processor = BatchProcessor()
translator = GoogleTranslator()

files = ["doc1.xlsx", "doc2.docx", "doc3.pdf"]
processor.process_files(files, translator, target_lang="en")
```

### 自定义翻译器

```python
from offitrans.translators.base_api import BaseTranslator

class CustomTranslator(BaseTranslator):
    def translate(self, text, source_lang="auto", target_lang="en"):
        # 实现自定义翻译逻辑
        return translated_text

# 使用自定义翻译器
translator = CustomTranslator()
```

### 缓存配置

```python
from offitrans.core.cache import TranslationCache

# 启用翻译缓存
cache = TranslationCache(cache_dir="./translation_cache")
translator = GoogleTranslator(cache=cache)
```

## 🏗️ 项目架构

```
offitrans/
├── cli/                    # 命令行界面
│   ├── __init__.py
│   └── main.py            # CLI入口点
├── core/                  # 核心功能
│   ├── base.py           # 基础类定义
│   ├── cache.py          # 缓存机制
│   ├── config.py         # 配置管理
│   └── utils.py          # 工具函数
├── processors/           # 文档处理器
│   ├── base.py          # 处理器基类
│   ├── excel.py         # Excel处理器
│   ├── pdf.py           # PDF处理器  
│   ├── powerpoint.py    # PowerPoint处理器
│   └── word.py          # Word处理器
├── translators/         # 翻译引擎
│   ├── base_api.py      # 翻译器基类
│   └── google.py        # Google翻译实现
└── exceptions/          # 异常定义
    └── errors.py        # 自定义异常
```

## 🧪 测试

### 运行测试

```bash
# 安装开发依赖
pip install -e .[dev]

# 运行所有测试
pytest

# 运行特定测试
pytest tests/unit/test_processors.py

# 运行测试并生成覆盖率报告
pytest --cov=offitrans --cov-report=html
```

### 测试结构

```
tests/
├── unit/                    # 单元测试
│   ├── test_core.py
│   ├── test_processors.py
│   └── test_translators.py
├── integration/             # 集成测试
│   └── test_end_to_end.py
└── fixtures/               # 测试数据
    ├── sample.xlsx
    ├── sample.docx
    └── sample.pdf
```

## 🤝 贡献

我们欢迎任何形式的贡献！请查看 [贡献指南](CONTRIBUTING_CN.md) 了解如何参与项目开发。

### 开发环境设置

```bash
# 克隆仓库
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 安装开发依赖
pip install -e .[dev]

# 安装pre-commit钩子
pre-commit install

# 运行测试
pytest
```

## 📝 许可证

本项目采用 [MIT许可证](LICENSE)。

## 🔗 相关链接

- [GitHub仓库](https://github.com/minglu6/Offitrans)
- [PyPI包](https://pypi.org/project/offitrans/)
- [问题反馈](https://github.com/minglu6/Offitrans/issues)
- [版本发布](https://github.com/minglu6/Offitrans/releases)
- [更新日志](CHANGELOG.md)

## 📊 项目状态

- ✅ Excel翻译 - 完全支持
- ✅ Word翻译 - 基础支持  
- ✅ PDF翻译 - 基础支持
- ✅ PPT翻译 - 基础支持
- ✅ CLI工具 - 完全支持
- 🔄 OCR支持 - 开发中
- 🔄 更多翻译引擎 - 计划中

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户！

### 特别感谢

- Google Translate API 提供可靠的翻译服务
- OpenPyXL 团队提供出色的Excel处理能力
- Python社区提供优秀的库和工具

---

<div align="center">
如果这个项目对您有帮助，请给我们一个 ⭐️
</div>