"""
翻译缓存使用示例和测试
"""

from translate_tools.google_translate import GoogleTranslator
from translate_tools.sungrow_translate import SunTranslator
from translate_tools.cache import get_global_cache


def test_translation_cache():
    """测试翻译缓存功能"""
    
    # 创建翻译器实例
    google_translator = GoogleTranslator(source_lang="zh", target_lang="en")
    sun_translator = SunTranslator(source_lang="zh", target_lang="en")
    
    # 测试文本
    test_texts = [
        "你好世界",
        "这是一个测试",
        "翻译缓存机制",
        "节省调试时间"
    ]
    
    print("=" * 50)
    print("🚀 开始测试翻译缓存功能")
    print("=" * 50)
    
    # 第一次翻译（会调用API）
    print("\n📡 第一次翻译（会调用API）:")
    for text in test_texts:
        result = google_translator.translate_text(text)
        print(f"  {text} -> {result}")
    
    # 第二次翻译（会使用缓存）
    print("\n💨 第二次翻译（使用缓存）:")
    for text in test_texts:
        result = google_translator.translate_text(text)
        print(f"  {text} -> {result}")
    
    # 显示缓存统计
    cache = get_global_cache()
    stats = cache.get_stats()
    print(f"\n📊 缓存统计: {stats}")
    
    # 手动保存缓存
    cache.save()
    
    print("\n✅ 测试完成！")


if __name__ == "__main__":
    test_translation_cache()