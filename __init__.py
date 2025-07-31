"""
Offitrans - Office文件翻译工具

一个强大的Office文件翻译工具库，支持PDF、Excel、PPT和Word文档的批量翻译。

主要特性：
- 支持多种Office文件格式 (Excel, Word, PDF, PPT)
- 保持原有格式和样式
- 批量翻译优化
- 多语言支持
- 图片保护功能
"""

__version__ = "1.0.0"
__author__ = "Offitrans Contributors"
__email__ = "offitrans@example.com"
__license__ = "MIT"
__url__ = "https://github.com/your-username/Offitrans"

# 导入主要的翻译器类
try:
    from excel_translate.translate_excel import ExcelTranslatorV2
    from translate_tools.google_translate import GoogleTranslator
    from translate_tools.base import Translator
    
    __all__ = [
        'ExcelTranslatorV2',
        'GoogleTranslator', 
        'Translator',
    ]
    
except ImportError:
    # 如果依赖未安装，只导出版本信息
    __all__ = [
        '__version__',
        '__author__',
        '__email__',
        '__license__',
        '__url__',
    ]

# 版本信息
def get_version():
    """获取版本信息"""
    return __version__

def get_info():
    """获取项目信息"""
    return {
        'name': 'Offitrans',
        'version': __version__,
        'author': __author__,
        'email': __email__,
        'license': __license__,
        'url': __url__,
        'description': '一个强大的Office文件翻译工具库'
    }