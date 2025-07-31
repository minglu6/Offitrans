import re
from typing import List


def should_translate(text: str) -> bool:
    """
    判断文本是否需要翻译
    
    :param text: 待判断的文本
    :return: True表示需要翻译，False表示跳过翻译
    """
    if not text or not text.strip():
        return False
    
    text = text.strip()
    
    # 纯数字不翻译
    if text.isdigit():
        return False
    
    # 纯符号不翻译 (如 ?, !, -, etc.)
    if re.fullmatch(r'[\W_]+', text):
        return False
    
    # 纯英文字母不翻译
    if re.fullmatch(r'[a-zA-Z]+', text):
        return False
    
    # 数字和英文字母组合不翻译 (如 ABC123, 123ABC, A1B2C3)
    if re.fullmatch(r'[a-zA-Z0-9]+', text):
        return False
    
    # 数字和标点符号组合不翻译 (如 12.5, 3.14%, $100, etc.)
    if re.fullmatch(r'[\d\W_]+', text):
        return False
    
    # URL和邮箱不翻译
    if re.search(r'https?://|www\.|@.*\.|\.com|\.org|\.net', text.lower()):
        return False
    
    # 文件路径不翻译 (如 C:\folder\file.txt)
    if re.search(r'[A-Za-z]:\\|/[a-zA-Z]|\.exe|\.dll|\.pdf|\.docx?|\.xlsx?|\.pptx?', text):
        return False
    
    # 变量名和代码片段不翻译 (包含下划线或驼峰命名)
    if re.search(r'[a-zA-Z]+_[a-zA-Z]+|[a-z]+[A-Z][a-z]*', text):
        return False
    
    # 常见单位和度量不翻译
    if re.fullmatch(r'\d+\s*(mm|cm|m|km|kg|g|ml|l|°C|°F|%|px|pt|em|rem)', text, re.IGNORECASE):
        return False
    
    # 版本号和序列号不翻译 (如 v1.2.3, Ver.2.0)
    if re.search(r'v\d+\.\d+|ver\.\d+|version\s*\d+', text.lower()):
        return False
    
    # 日期格式不翻译 (如 2024-01-01, 01/01/2024)
    if re.search(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4}', text):
        return False
    
    # 时间格式不翻译 (如 12:30, 9:00 AM)
    if re.search(r'\d{1,2}:\d{2}(\s*(AM|PM))?', text.upper()):
        return False
    
    # 检查是否包含中文字符，如果包含则需要翻译
    if re.search(r'[\u4e00-\u9fff]', text):
        return True
    
    # 检查是否包含其他需要翻译的非拉丁字符
    if re.search(r'[^\x00-\x7F]', text):
        return True
    
    # 纯英文但可能是短语或句子，包含空格的英文文本
    if ' ' in text and re.search(r'[a-zA-Z]', text):
        # 如果是简单的标识符组合 (如 "Item 1", "Page 2")，不翻译
        if re.fullmatch(r'[A-Za-z]+\s*\d+|\d+\s*[A-Za-z]+', text):
            return False
        # 其他包含空格的英文可能是需要翻译的短语
        return False  # 暂时不翻译英文短语，避免误翻
    
    # 默认情况：如果都不匹配上述规则，不翻译
    return False


def normalize_text(text: str) -> str:
    """
    标准化文本，用于更好的去重
    
    :param text: 原始文本
    :return: 标准化后的文本
    """
    if not text:
        return text
    # 去除首尾空格，标准化内部空格
    normalized = re.sub(r'\s+', ' ', text.strip())
    return normalized


def filter_translatable_texts(texts: List[str]) -> tuple[List[str], List[str]]:
    """
    过滤出需要翻译和不需要翻译的文本
    
    :param texts: 文本列表
    :return: (需要翻译的文本列表, 不需要翻译的文本列表)
    """
    needs_translation = []
    skip_translation = []
    
    for text in texts:
        if should_translate(text):
            needs_translation.append(text)
        else:
            skip_translation.append(text)
    
    return needs_translation, skip_translation