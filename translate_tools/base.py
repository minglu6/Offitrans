from abc import ABC, abstractmethod
from typing import Optional, List
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed



class Translator(ABC):

    def __init__(self, source_lang: str,
                 target_lang: str,
                 max_workers: int = 5, **kwargs):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.max_workers = max_workers
        self.timeout = 120  # 增加请求超时时间
        self.retry_count = 5  # 增加重试次数
        self.retry_delay = 2  # 增加重试间隔(秒)
        self.batch_size = 20  # 批处理大小
        # 支持的语言映射
        self.supported_languages = {
            'zh': 'zh',  # 中文
            'en': 'en',  # 英文
            'th': 'th',  # 泰文
        }
        self._lock = threading.Lock()

    @abstractmethod
    def translate_text(self, text: str) -> str:
        """
        Translate a single text string from source language to target language.

        :param text: The text to translate.
        :return: Translated text.
        """
        pass

    def translate_text_batch(self, texts: List[str]) -> List[str]:
        """
        Translate a batch of text strings from source language to target language.
        Uses multithreading for improved performance.

        :param texts: List of text strings to translate.
        :return: List of translated text strings.
        """
        if not texts:
            return []

        # 使用多线程执行翻译
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # 提交所有翻译任务
            future_to_index = {executor.submit(self.translate_text, text): i for i, text in enumerate(texts)}

            # 初始化结果列表，明确类型
            results: List[str] = [""] * len(texts)

            # 收集结果
            for future in as_completed(future_to_index):
                index = future_to_index[future]
                try:
                    result = future.result()
                    results[index] = result if result is not None else ""
                except Exception as exc:
                    print(f'Translation at index {index} generated an exception: {exc}')
                    results[index] = texts[index]  # 出错时返回原文

            return results

    def translate_text_batch_simple(self, texts: List[str]) -> List[str]:
        """
        Simple multithreaded version using map.

        :param texts: List of text strings to translate.
        :return: List of translated text strings.
        """
        if not texts:
            return []

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            return list(executor.map(self.translate_text, texts))
