#!/usr/bin/env python3
"""
Offitrans 使用示例

这个文件展示了如何使用 Offitrans 进行各种Office文件的翻译。
"""

import os
from excel_translate.translate_excel import ExcelTranslatorV2
from translate_tools.google_translate import GoogleTranslator

def example_excel_translation():
    """Excel文件翻译示例"""
    print("=" * 50)
    print("Excel文件翻译示例")
    print("=" * 50)
    
    # 创建Excel翻译器
    translator = ExcelTranslatorV2(
        font_size_adjustment=0.8  # 字体大小调整比例
    )
    
    # 示例文件路径
    input_file = "example_input.xlsx"
    output_file = "example_output_translated.xlsx"
    
    # 检查输入文件是否存在
    if os.path.exists(input_file):
        print(f"正在翻译文件: {input_file}")
        
        # 分析文件结构
        print("分析Excel文件结构...")
        analysis = translator.analyze_excel_structure(input_file)
        
        # 执行翻译
        print("开始翻译...")
        success = translator.replace_text_in_excel(
            excel_path=input_file,
            output_path=output_file,
            target_language='en'  # 翻译为英文
        )
        
        if success:
            print(f"翻译成功！输出文件: {output_file}")
            
            # 智能调整列宽
            print("调整列宽...")
            translator.smart_adjust_column_width(output_file)
            print("列宽调整完成！")
        else:
            print("翻译失败")
    else:
        print(f"警告: 输入文件不存在: {input_file}")
        print("请准备一个Excel文件进行测试")

def example_text_translation():
    """文本翻译示例"""
    print("\n" + "=" * 50)
    print("文本翻译示例")
    print("=" * 50)
    
    # 创建翻译器
    translator = GoogleTranslator(
        source_lang='zh',
        target_lang='en',
        max_workers=3
    )
    
    # 单个文本翻译
    text = "你好，世界！"
    print(f"原文: {text}")
    
    translated = translator.translate_text(text)
    print(f"译文: {translated}")
    
    # 批量文本翻译
    texts = [
        "欢迎使用Offitrans",
        "这是一个强大的翻译工具",
        "支持多种Office文件格式",
        "保持原有格式和样式"
    ]
    
    print(f"\n批量翻译 {len(texts)} 个文本:")
    for i, text in enumerate(texts):
        print(f"{i+1}. {text}")
    
    print("\n翻译结果:")
    translated_texts = translator.translate_text_batch(texts)
    for i, (original, translated) in enumerate(zip(texts, translated_texts)):
        print(f"{i+1}. {original} -> {translated}")

def example_supported_languages():
    """支持的语言示例"""
    print("\n" + "=" * 50)
    print("支持的语言")
    print("=" * 50)
    
    from translate_tools.google_translate import get_supported_languages
    
    languages = get_supported_languages()
    print("当前支持的语言:")
    for code, name in languages.items():
        print(f"  {code}: {name}")

def main():
    """主函数"""
    print("Offitrans 使用示例")
    print("这个示例展示了如何使用 Offitrans 进行文件翻译")
    
    try:
        # Excel翻译示例
        example_excel_translation()
        
        # 文本翻译示例
        example_text_translation()
        
        # 支持的语言
        example_supported_languages()
        
        print("\n" + "=" * 50)
        print("示例运行完成！")
        print("=" * 50)
        print("更多功能请参考:")
        print("- README.md: 详细的使用文档")
        print("- CONTRIBUTING.md: 贡献指南")
        print("- GitHub: https://github.com/your-username/Offitrans")
        
    except ImportError as e:
        print(f"导入错误: {e}")
        print("请确保已正确安装所有依赖:")
        print("pip install -r requirements.txt")
        
    except Exception as e:
        print(f"运行错误: {e}")
        print("请检查配置和输入文件")

if __name__ == "__main__":
    main()