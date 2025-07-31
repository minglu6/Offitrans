import json
import os
import hashlib
import atexit
from typing import Dict, Optional, Any, List
from functools import wraps
import threading


class TranslationCache:
    """翻译缓存管理器"""
    
    def __init__(self, cache_file: str = "translation_cache.json"):
        """
        初始化翻译缓存
        
        :param cache_file: 缓存文件路径
        """
        self.cache_file = cache_file
        self._cache: Dict[str, str] = {}
        self._lock = threading.Lock()
        self._load_cache()
        
        # 注册程序退出时的保存函数，确保缓存不丢失
        atexit.register(self._save_cache_on_exit)
    
    def _generate_key(self, text: str, source_lang: str, target_lang: str) -> str:
        """
        生成缓存键
        
        :param text: 原文
        :param source_lang: 源语言
        :param target_lang: 目标语言
        :return: 缓存键
        """
        key_string = f"{source_lang}:{target_lang}:{text}"
        return hashlib.md5(key_string.encode('utf-8')).hexdigest()
    
    def _load_cache(self):
        """从文件加载缓存"""
        try:
            if os.path.exists(self.cache_file):
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
                print(f"✅ 翻译缓存已加载，共 {len(self._cache)} 条记录")
            else:
                self._cache = {}
                print("📝 创建新的翻译缓存文件")
        except Exception as e:
            print(f"⚠️ 加载缓存文件失败: {e}")
            self._cache = {}
    
    def _save_cache(self):
        """保存缓存到文件"""
        try:
            with self._lock:
                with open(self.cache_file, 'w', encoding='utf-8') as f:
                    json.dump(self._cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ 保存缓存文件失败: {e}")
    
    def _save_cache_on_exit(self):
        """程序退出时保存缓存"""
        try:
            self._save_cache()
            print(f"🔒 程序退出时自动保存翻译缓存: {len(self._cache)} 条记录")
        except Exception as e:
            print(f"⚠️ 程序退出时保存缓存失败: {e}")
    
    def get(self, text: str, source_lang: str, target_lang: str) -> Optional[str]:
        """
        从缓存获取翻译结果
        
        :param text: 原文
        :param source_lang: 源语言
        :param target_lang: 目标语言
        :return: 翻译结果，如果不存在则返回None
        """
        if not text or not text.strip():
            return text
            
        key = self._generate_key(text, source_lang, target_lang)
        return self._cache.get(key)
    
    def set(self, text: str, translation: str, source_lang: str, target_lang: str, force_save: bool = False):
        """
        设置缓存
        
        :param text: 原文
        :param translation: 翻译结果
        :param source_lang: 源语言
        :param target_lang: 目标语言
        :param force_save: 是否强制立即保存
        """
        if not text or not text.strip() or not translation:
            return
            
        key = self._generate_key(text, source_lang, target_lang)
        with self._lock:
            self._cache[key] = translation
        
        # 强制保存或每添加10条记录自动保存一次
        if force_save or len(self._cache) % 10 == 0:
            self._save_cache()
    
    def clear(self):
        """清空缓存"""
        with self._lock:
            self._cache.clear()
        self._save_cache()
        print("🗑️ 翻译缓存已清空")
    
    def get_batch(self, texts: List[str], source_lang: str, target_lang: str) -> Dict[str, Optional[str]]:
        """
        批量从缓存获取翻译结果
        
        :param texts: 原文列表
        :param source_lang: 源语言
        :param target_lang: 目标语言
        :return: 原文到翻译结果的字典，未找到的为None
        """
        result = {}
        for text in texts:
            if text and text.strip():
                cached_result = self.get(text, source_lang, target_lang)
                result[text] = cached_result
            else:
                result[text] = text  # 空文本直接返回
        return result
    
    def set_batch(self, text_translation_pairs: Dict[str, str], source_lang: str, target_lang: str):
        """
        批量设置缓存，并立即保存到文件
        
        :param text_translation_pairs: 原文到翻译结果的字典
        :param source_lang: 源语言
        :param target_lang: 目标语言
        """
        count = 0
        for text, translation in text_translation_pairs.items():
            if text and text.strip() and translation and translation != text:
                key = self._generate_key(text, source_lang, target_lang)
                with self._lock:
                    self._cache[key] = translation
                count += 1
        
        # 批量操作后立即保存，确保数据持久化
        if count > 0:
            self._save_cache()
            print(f"📁 批量缓存已保存: {count} 条记录")

    def save(self):
        """手动保存缓存"""
        self._save_cache()
        print(f"💾 翻译缓存已保存，共 {len(self._cache)} 条记录")
    
    def get_stats(self) -> Dict[str, Any]:
        """获取缓存统计信息"""
        return {
            "total_entries": len(self._cache),
            "cache_file": self.cache_file,
            "file_exists": os.path.exists(self.cache_file)
        }


# 全局缓存实例
_global_cache = TranslationCache()


def cached_translation(cache_instance: Optional[TranslationCache] = None):
    """
    翻译缓存装饰器
    
    :param cache_instance: 缓存实例，如果为None则使用全局缓存
    """
    def decorator(translate_func):
        @wraps(translate_func)
        def wrapper(self, text: str) -> str:
            # 使用指定的缓存实例或全局缓存
            cache = cache_instance or _global_cache
            
            # 尝试从缓存获取
            cached_result = cache.get(text, self.source_lang, self.target_lang)
            if cached_result is not None:
                print(f"🎯 缓存命中: {text[:50]}...")
                return cached_result
            
            # 缓存未命中，调用原始翻译函数
            print(f"🌐 API调用: {text[:50]}...")
            result = translate_func(self, text)
            
            # 将结果存入缓存（单个翻译强制立即保存）
            if result and result != text:  # 只缓存成功的翻译结果
                cache.set(text, result, self.source_lang, self.target_lang, force_save=True)
            
            return result
        return wrapper
    return decorator


def get_global_cache() -> TranslationCache:
    """获取全局缓存实例"""
    return _global_cache


def set_global_cache_file(cache_file: str):
    """设置全局缓存文件路径"""
    global _global_cache
    _global_cache = TranslationCache(cache_file)