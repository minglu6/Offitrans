import html
import time

import requests

from .base import Translator
from .cache import cached_translation

api_key = "AIzaSyCgpcTlrNyhNAGZWXEPD8LuexBNNFlpTlM"
def get_supported_languages() -> dict:
    """获取支持的语言列表"""
    return {
        'zh': '中文',
        'en': 'English',
        'th': 'ไทย',
        'ja': '日本語',
        'ko': '한국어',
        'fr': 'Français',
        'de': 'Deutsch',
        'es': 'Español'
    }


class GoogleTranslator(Translator):
    """Google翻译API实现类"""

    def __init__(self, source_lang: str = "zh",
                 target_lang: str = "en",
                 max_workers: int = 5,
                 **kwargs):
        super().__init__(source_lang, target_lang, max_workers, **kwargs)
        self.api_url = "https://translation.googleapis.com/language/translate/v2"

    @cached_translation()
    def translate_text(self, text: str) -> str:
        """
        翻译单个文本（带重试机制）

        :param text: 要翻译的文本
        :return: 翻译后的文本
        """
        if not text or not text.strip():
            return text

        retry_count = 0
        while retry_count < self.retry_count:
            try:
                # 使用Google翻译的官方API
                base_url = self.api_url

                # 语言映射
                lang_map = {
                    'en': 'en',
                    'zh': 'zh',
                    'th': 'th',
                    'ja': 'ja',
                    'ko': 'ko',
                    'fr': 'fr',
                    'de': 'de',
                    'es': 'es'
                }

                target_lang_code = lang_map.get(self.target_lang, self.target_lang)
                source_lang_code = lang_map.get(self.source_lang, self.source_lang)

                # 构造请求参数
                params = {
                    'key': api_key,  # 使用API密钥
                    'source': source_lang_code,  # 源语言
                    'target': target_lang_code,  # 目标语言
                    'format': 'text',
                    'q': text
                }

                # 使用requests发送POST请求，增加超时时间和连接超时
                response = requests.post(
                    base_url, 
                    data=params, 
                    timeout=(30, 120)  # (连接超时, 读取超时)
                )
                response.raise_for_status()

                # 解析结果
                result_json = response.json()
                if ('data' in result_json and
                        'translations' in result_json['data'] and
                        len(result_json['data']['translations']) > 0):
                    translated_text = result_json['data']['translations'][0]['translatedText']
                    # 解码HTML实体
                    translated_text = html.unescape(translated_text)
                    return translated_text
                else:
                    raise Exception("Invalid response format")
                    
            except Exception as e:
                retry_count += 1
                if retry_count < self.retry_count:
                    print(f"Google翻译API调用失败，{self.retry_delay}秒后重试 ({retry_count}/{self.retry_count}): {e}")
                    time.sleep(self.retry_delay * retry_count)  # 逐渐增加延迟
                else:
                    print(f"Google翻译API调用失败，已达到最大重试次数: {e}")
                    # 返回原文作为备用方案
                    return text

        return text

    @cached_translation()
    def translate_batch_with_retry(self, texts: list) -> list:
        """
        带重试机制的批量翻译

        :param texts: 文本列表
        :return: 翻译结果列表
        """
        results = []

        for text in texts:
            retry_count = 0
            while retry_count < self.retry_count:
                try:
                    result = self.translate_text(text)
                    results.append(result)
                    break
                except Exception as e:
                    retry_count += 1
                    if retry_count < self.retry_count:
                        print(f"翻译失败，{self.retry_delay}秒后重试 ({retry_count}/{self.retry_count}): {e}")
                        time.sleep(self.retry_delay)
                    else:
                        print(f"翻译失败，已达到最大重试次数: {e}")
                        results.append(text)  # 返回原文

        return results

    def validate_api_key(self) -> bool:
        try:
            # 测试翻译一个简单的词汇
            test_result = self.translate_text("测试")
            return test_result != "测试"  # 如果翻译结果不等于原文，说明API有效
        except Exception as e:
            print(f"API密钥验证失败: {e}")
            return False


if __name__ == "__main__":
    translator = GoogleTranslator(source_lang="zh", target_lang="en")
    flag = translator.validate_api_key()
    print("flag:", flag)
