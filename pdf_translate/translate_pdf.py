import os
import time
from functools import lru_cache
from typing import List, Dict

import fitz  # pymupdf
import pdfplumber


class PDFTableTranslator:
    def __init__(self, translate_api_key: str = None):
        """
        初始化PDF表格翻译器

        Args:
            translate_api_key: 翻译API密钥
        """
        self.translate_api_key = translate_api_key
        self.thai_font_path = None
        self.embedded_fonts = {}  # 缓存已嵌入的字体
        self._setup_fonts()

    def _setup_fonts(self):
        """设置字体支持"""
        # 查找泰文字体
        possible_paths = [
            os.path.join(os.path.dirname(__file__), "font", "NotoSansThai-Regular.ttf"),
            "font/NotoSansThai-Regular.ttf",
            "NotoSansThai-Regular.ttf"
        ]

        for path in possible_paths:
            if os.path.exists(path):
                self.thai_font_path = path
                print(f"✅ 找到泰文字体: {path}")
                break

        if not self.thai_font_path:
            print("⚠️  未找到泰文字体，将使用默认字体")

    def _detect_language(self, text: str) -> str:
        """
        检测文本语言

        Args:
            text: 输入文本

        Returns:
            语言代码 ('th' for Thai, 'zh' for Chinese, 'en' for English)
        """
        if not text:
            return 'en'

        # 检测泰文字符 (U+0E00-U+0E7F)
        thai_chars = sum(1 for c in text if '\u0e00' <= c <= '\u0e7f')

        # 检测中文字符 (U+4E00-U+9FFF)
        chinese_chars = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')

        total_chars = len(text)

        if thai_chars > 0:
            return 'th'
        elif chinese_chars > 0:
            return 'zh'
        else:
            return 'en'

    def _get_font_for_language(self, text: str, page: fitz.Page = None) -> dict:
        """
        根据文本语言选择合适的字体，并在页面级别嵌入泰文字体

        Args:
            text: 文本内容
            page: PDF页面对象（用于字体嵌入）

        Returns:
            字体信息字典 {'fontname': str, 'fontfile': str or None}
        """
        lang = self._detect_language(text)

        if lang == 'th' and self.thai_font_path:
            # 泰文字体处理 - 页面级别嵌入
            if page is not None:
                try:
                    # 检查是否已在该页面嵌入过泰文字体
                    page_id = id(page)
                    if page_id not in self.embedded_fonts:
                        # 读取字体文件并在页面级别嵌入
                        with open(self.thai_font_path, 'rb') as f:
                            font_buffer = f.read()

                        # 页面级别字体嵌入
                        font_result = page.insert_font(fontbuffer=font_buffer, fontname="NotoSansThai")

                        # 处理返回值格式
                        if isinstance(font_result, int):
                            font_name = "NotoSansThai"  # 使用我们指定的名称
                        else:
                            font_name = str(font_result)

                        # 缓存字体名称
                        self.embedded_fonts[page_id] = font_name
                        print(f"✅ 页面级别泰文字体嵌入成功: {font_name}")

                    # 返回已嵌入的字体名称
                    return {'fontname': self.embedded_fonts[page_id], 'fontfile': None}
                except Exception as e:
                    print(f"❌ 页面级别字体嵌入失败: {e}")
                    # 降级到fontfile方式
                    return {'fontname': None, 'fontfile': self.thai_font_path}
            else:
                # 没有页面对象时，使用字体文件方式
                return {'fontname': None, 'fontfile': self.thai_font_path}
        elif lang == 'zh':
            # 中文字体
            return {'fontname': 'china-s', 'fontfile': None}
        else:
            # 英文或其他语言
            return {'fontname': 'helv', 'fontfile': None}

    def extract_tables_pdfplumber(self, pdf_path: str) -> List[Dict]:
        """
        使用pdfplumber提取表格数据

        Args:
            pdf_path: PDF文件路径

        Returns:
            表格数据列表
        """
        tables_data = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # 提取表格
                tables = page.extract_tables()

                for table_index, table in enumerate(tables):
                    if table:
                        # 获取表格位置信息
                        table_bbox = None
                        try:
                            # 尝试获取表格边界框
                            table_bbox = page.bbox
                        except:
                            pass

                        tables_data.append({
                            'page': page_num,
                            'table_index': table_index,
                            'data': table,
                            'bbox': table_bbox
                        })

        return tables_data

    def translate_text_batch(self, texts: List[str], target_lang: str = 'th') -> List[str]:
        """
        批量翻译文本（示例使用百度翻译API）

        Args:
            texts: 待翻译文本列表
            target_lang: 目标语言代码

        Returns:
            翻译结果列表
        """
        if not self.translate_api_key:
            # 如果没有API密钥，返回模拟翻译结果（保持原文，不添加前缀）
            return texts  # 直接返回原文，不添加[TRANSLATED]前缀

        translated_texts = []

        for text in texts:
            if not text or text.strip() == "":
                translated_texts.append("")
                continue

            try:
                translated_text = self._call_translate_api(text, target_lang)
                translated_texts.append(translated_text)

                time.sleep(0.1)

            except Exception as e:
                print(f"翻译失败: {text} - {str(e)}")
                translated_texts.append(text)  # 翻译失败时保留原文

        return translated_texts

    def _call_translate_api(self, text: str, target_lang: str) -> str:
        """
        调用翻译API

        Args:
            text: 待翻译文本
            target_lang: 目标语言

        Returns:
            翻译结果
        """
        import requests
        import json

        # API配置
        api_url = "https://aigw.sungrow.cn/llm/v1/chat/completions"

        # 语言映射
        lang_map = {
            'en': '英文',
            'zh': '中文',
            'th': '泰文',
        }

        target_lang_name = lang_map.get(target_lang, target_lang)

        # 构建请求参数
        payload = {
            "max_tokens": 8092,
            "stream": False,
            "messages": [
                {
                    "content": f"你是精通中文和{target_lang_name}翻译的专家，擅长将中文翻译为{target_lang_name}，请将下面文字翻译为{target_lang_name}并仅返回翻译结果即可：{text}",
                    "role": "user"
                }
            ],
            "model": "qwen3_1",
            "temperature": 0.5
        }

        # 设置请求头，包含Bearer Token认证
        headers = {
            'Content-Type': 'application/json'
        }

        # 如果有API密钥，添加Bearer Token认证
        if self.translate_api_key:
            headers['Authorization'] = f'Bearer {self.translate_api_key}'

        try:
            # 发送POST请求
            response = requests.post(api_url, json=payload, headers=headers, timeout=30)
            response.raise_for_status()

            # 解析响应
            result = response.json()

            # 提取翻译结果
            if 'choices' in result and len(result['choices']) > 0:
                translated_text = result['choices'][0]['message']['content'].strip()
                return translated_text
            else:
                print(f"API响应格式错误: {result}")
                return text

        except requests.exceptions.RequestException as e:
            print(f"API请求失败: {e}")
            return text
        except json.JSONDecodeError as e:
            print(f"JSON解析失败: {e}")
            return text
        except Exception as e:
            print(f"翻译API调用失败: {e}")
            return text

    def replace_text_in_pdf(self, pdf_path: str, text_replacements: Dict[str, str], output_path: str,
                            font_size: float = None, font_name: str = "china-s",
                            auto_fit_font: bool = True):
        """
        在PDF中替换文本，并将译文在单元格内垂直居中并向下微调

        Args:
            pdf_path: 原PDF路径
            text_replacements: 文本替换字典 {原文: 译文}
            output_path: 输出PDF路径
            font_size: 指定字体大小，None时自动检测原文字体大小
            font_name: 字体名称，支持 "helv", "times", "cour", "china-s" 等
            auto_fit_font: 是否自动调整字体大小以适应原文本区域
        """
        import fitz  # pymupdf
        doc = fitz.open(pdf_path)

        # 智能字体嵌入将在循环中根据文本内容决定

        for page_num in range(len(doc)):
            page = doc[page_num]
            text_dict = page.get_text("dict")

            for original_text, translated_text in text_replacements.items():
                if not original_text.strip() or not translated_text.strip():
                    continue

                # 查找所有匹配区域
                text_areas = page.search_for(original_text)
                original_font_size = self._get_text_font_size(text_dict, original_text)

                for area in text_areas:
                    # 覆盖原文本
                    page.draw_rect(area, color=(1, 1, 1), fill=(1, 1, 1))

                    # 智能选择字体（根据翻译后的文本语言）
                    font_info = self._get_font_for_language(translated_text, page)

                    # 选择字体大小
                    if font_size is not None:
                        use_font_size = font_size
                    elif original_font_size:
                        use_font_size = original_font_size
                    else:
                        use_font_size = 10

                    # 自动适应
                    if auto_fit_font:
                        # 使用字体名称或默认值进行自适应计算
                        font_for_calc = font_info.get('fontname', 'helv')
                        use_font_size = self._fit_text_to_area(
                            translated_text, area, use_font_size, font_for_calc
                        )

                    # —— 计算垂直居中并微调向下 —— #
                    cell_h = area.height
                    text_h = use_font_size

                    # 1. 基础垂直居中计算
                    # PDF坐标系Y轴向上，所以area.y0是底部，area.y1是顶部
                    center_y = area.y0 + cell_h / 2

                    # 2. 获取字体基线偏移
                    font_for_baseline = font_info.get('fontname', 'helv')
                    baseline = self._get_font_baseline_offset(font_for_baseline, use_font_size)

                    # 3. 应用基线偏移（通常基线在字符高度的80%左右）
                    text_baseline_y = center_y + baseline

                    # 4. 额外微调：向下偏移（在PDF坐标系中是减小Y值）
                    fine_tune_offset = text_h * 0.15  # 向下微调15%字符高度
                    final_y = text_baseline_y - fine_tune_offset

                    # 确保不超出区域边界
                    final_y = max(area.y0 + 2, min(final_y, area.y1 - 2))

                    # 最终插入点
                    insert_x = area.x0 + 1
                    insert_y = final_y
                    pt = fitz.Point(insert_x, insert_y)

                    # 调试信息
                    print(f"文本位置调整: 区域高度={cell_h:.1f}, 字体大小={use_font_size:.1f}")
                    print(f"  中心Y={center_y:.1f}, 基线偏移={baseline:.1f}, 微调偏移={fine_tune_offset:.1f}")
                    print(f"  最终Y={final_y:.1f} (原区域: {area.y0:.1f}-{area.y1:.1f})")

                    # 插入新文本
                    try:
                        # 根据字体信息选择插入方式
                        if font_info.get('fontfile'):
                            # 使用字体文件
                            page.insert_text(
                                pt,
                                translated_text,
                                fontsize=use_font_size,
                                fontfile=font_info['fontfile'],
                                color=(0, 0, 0)
                            )
                            print(f"✅ 文本插入成功: '{translated_text}' 使用字体文件: {font_info['fontfile']}")
                        else:
                            # 使用字体名称
                            page.insert_text(
                                pt,
                                translated_text,
                                fontsize=use_font_size,
                                fontname=font_info['fontname'],
                                color=(0, 0, 0)
                            )
                            print(f"✅ 文本插入成功: '{translated_text}' 使用字体: {font_info['fontname']}")
                    except Exception as e:
                        print(f"❌ 文本插入失败: {e}，降级使用默认字体")
                        page.insert_text(
                            pt,
                            translated_text,
                            fontsize=max(6, use_font_size),
                            fontname="helv",
                            color=(0, 0, 0)
                        )

        doc.save(output_path)
        doc.close()

    def _get_text_font_size(self, text_dict: dict, target_text: str) -> float:
        """
        从文本字典中获取指定文本的字体大小

        Args:
            text_dict: 页面文本字典
            target_text: 目标文本

        Returns:
            字体大小，找不到则返回None
        """
        for block in text_dict.get("blocks", []):
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if target_text in span.get("text", ""):
                            return span.get("size", None)
        return None

    def _fit_text_to_area(self, text: str, area: fitz.Rect, initial_font_size: float,
                          font_name: str) -> float:
        """
        自动调整字体大小以适应指定区域

        Args:
            text: 要插入的文本
            area: 文本区域
            initial_font_size: 初始字体大小
            font_name: 字体名称

        Returns:
            调整后的字体大小
        """
        # 简单的字体大小适应算法
        area_width = area.width
        area_height = area.height

        # 估算文本宽度（简化计算）
        char_width_ratio = 0.6  # 字符宽度与字体大小的比例
        estimated_width = len(text) * initial_font_size * char_width_ratio

        # 如果文本太宽，缩小字体
        if estimated_width > area_width:
            scale_factor = area_width / estimated_width
            adjusted_font_size = initial_font_size * scale_factor * 0.9  # 留10%余量
        else:
            adjusted_font_size = initial_font_size

        # 确保字体大小在合理范围内
        return max(6, min(adjusted_font_size, initial_font_size * 1.2))

    def process_pdf_advanced(self, pdf_path: str, output_path: str,
                             target_lang: str = 'en',
                             font_size: float | None = None,
                             font_name: str = "helv",
                             auto_fit_font: bool = True):

        print(f"\n>>> 开始处理: {pdf_path}")
        doc = fitz.open(pdf_path)
        self._doc = doc

        # -------- 采集文本信息（含基线） --------
        all_text_info = []
        for page_num in range(len(doc)):
            page = doc[page_num]
            for block in page.get_text("dict")["blocks"]:
                if "lines" not in block:
                    continue
                for line in block["lines"]:
                    for span in line["spans"]:
                        txt = span["text"].strip()
                        if not txt:
                            continue
                        all_text_info.append({
                            "page_num": page_num,
                            "text": txt,
                            "bbox": span["bbox"],
                            "origin": span["origin"],
                            "asc": span["ascender"],
                            "des": span["descender"],
                            "size": span["size"]
                        })

        unique_texts = list({s["text"] for s in all_text_info})
        translation_dict = dict(zip(
            unique_texts,
            self.translate_text_batch(unique_texts, target_lang)
        ))

        # -------- 逐页替换 --------
        replaced = 0
        for page_num in range(len(doc)):
            page = doc[page_num]

            spans = [s for s in all_text_info if s["page_num"] == page_num]

            for sp in spans:
                src = sp["text"]
                dst = translation_dict.get(src, src)
                if src == dst:
                    continue

                bbox = fitz.Rect(sp["bbox"])
                page.draw_rect(bbox, color=(1, 1, 1), fill=(1, 1, 1))

                # 智能选择字体
                font_info = self._get_font_for_language(dst, page)

                use_sz = font_size or sp["size"]
                if auto_fit_font:
                    font_for_calc = font_info.get('fontname', 'helv')
                    use_sz = self._fit_text_to_area(dst, bbox, use_sz, font_for_calc)

                font_for_baseline = font_info.get('fontname', 'helv')
                pt = self._baseline_point(
                    bbox, sp["origin"], sp["des"], font_for_baseline, use_sz
                )

                try:
                    # 根据字体信息选择插入方式
                    if font_info.get('fontfile'):
                        # 使用字体文件
                        page.insert_text(pt, dst,
                                         fontsize=use_sz,
                                         fontfile=font_info['fontfile'],
                                         color=(0, 0, 0))
                        print(f"✅ 替换: '{src}' → '{dst}' 使用字体文件: {font_info['fontfile']}")
                    else:
                        # 使用字体名称
                        page.insert_text(pt, dst,
                                         fontsize=use_sz,
                                         fontname=font_info['fontname'],
                                         color=(0, 0, 0))
                        print(f"✅ 替换: '{src}' → '{dst}' 使用字体: {font_info['fontname']}")
                    replaced += 1
                except Exception as e:
                    print(f"❌ 文本替换失败: {e}")
                    # 降级处理
                    page.insert_text(pt, dst,
                                     fontsize=use_sz,
                                     fontname="helv",
                                     color=(0, 0, 0))
                    replaced += 1

        doc.save(output_path)
        doc.close()
        self._doc = None
        print(f"<<< 完成！共替换 {replaced} 处，输出: {output_path}")

    ######################################################################
    # 2) 计算插入点改用基线，完全舍弃旧的 _calculate_text_position
    ######################################################################
    def _baseline_point(self, bbox, origin, orig_des, new_fontname, fontsize):
        asc_new, des_new = self._get_font_metrics(new_fontname)
        y = origin[1] + ((orig_des - des_new) / 1000) * fontsize
        return fitz.Point(bbox.x0 + 1, y)

    @lru_cache(maxsize=32)
    def _get_font_metrics(self, fontname: str) -> tuple[int, int]:
        """
        返回指定字体的 (ascender, descender)，单位：1000-em 千分比
        """
        import os

        # 1. 特殊处理自定义字体名称（如页面级别嵌入的字体）
        if fontname in ['NotoSansThai', 'noto', 'thai'] or (fontname and 'Noto' in fontname):
            # 对于泰文字体，使用合理的默认度量值
            return 1069, -293  # Noto Sans Thai的典型度量值

        # 2. 尝试把 fontname 当作"已知字体名"
        try:
            f = fitz.Font(fontname=fontname)
            return f.ascender, f.descender
        except RuntimeError:
            pass  # 说明 fontname 不是内置 / 已嵌入字体

        # 3. 如果 fontname 是文件路径 ⇒ 先嵌入再读取
        if os.path.isfile(fontname) and self._doc:
            try:
                internal_name = self._doc.insert_font(fontfile=fontname)
                f = fitz.Font(fontname=internal_name)
                return f.ascender, f.descender
            except Exception as e:
                raise RuntimeError(f"无法嵌入字体 {fontname}: {e}")

        # 4. 默认使用helv的度量
        try:
            f = fitz.Font(fontname='helv')
            return f.ascender, f.descender
        except:
            # 最后的回退值
            return 800, -200

    def batch_process_pdfs(self, input_folder: str, output_folder: str, target_lang: str = 'th'):
        """
        批量处理PDF文件

        Args:
            input_folder: 输入文件夹路径
            output_folder: 输出文件夹路径
            target_lang: 目标语言
        """
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]

        for pdf_file in pdf_files:
            print(f"处理文件: {pdf_file}")

            input_path = os.path.join(input_folder, pdf_file)
            output_path = os.path.join(output_folder, f"translated_{pdf_file}")

            try:
                # 方法1：使用表格提取方式
                tables_data = self.extract_tables_pdfplumber(input_path)

                if tables_data:
                    # 如果检测到表格，使用表格处理方式
                    self.process_tables_translation(input_path, output_path, target_lang)
                else:
                    # 如果没有检测到表格，使用通用文本处理
                    self.process_pdf_advanced(input_path, output_path, target_lang)

                print(f"完成: {pdf_file}")

            except Exception as e:
                print(f"处理失败 {pdf_file}: {str(e)}")

    def process_tables_translation(self, pdf_path: str, output_path: str, target_lang: str,
                                   font_size: float = None, font_name: str = "china-s",
                                   auto_fit_font: bool = True):
        """
        专门处理表格翻译

        Args:
            pdf_path: 输入PDF路径
            output_path: 输出PDF路径
            target_lang: 目标语言
            font_size: 指定字体大小，None时自动检测
            font_name: 字体名称
            auto_fit_font: 是否自动调整字体大小
        """
        # 提取表格数据
        tables_data = self.extract_tables_pdfplumber(pdf_path)

        # 收集所有需要翻译的文本
        all_texts = []
        for table_info in tables_data:
            for row in table_info['data']:
                for cell in row:
                    if cell and cell.strip():
                        all_texts.append(cell.strip())

        # 去重
        unique_texts = list(set(all_texts))

        # 批量翻译
        translated_texts = self.translate_text_batch(unique_texts, target_lang)

        # 创建翻译字典
        translation_dict = dict(zip(unique_texts, translated_texts))

        # 在PDF中替换文本
        self.replace_text_in_pdf(pdf_path, translation_dict, output_path,
                                 font_size=font_size, font_name=font_name,
                                 auto_fit_font=auto_fit_font)

    def _get_font_baseline_offset(self, fontname: str, fontsize: float) -> float:
        """
        获取字体的基线偏移量

        Args:
            fontname: 字体名称
            fontsize: 字体大小

        Returns:
            基线偏移量
        """
        # 不同字体的基线偏移量（经验值）
        font_offset_ratios = {
            'helv': 0.2,
            'times': 0.25,
            'cour': 0.2,
            'china-s': 0.15,
            'china-ss': 0.15,
            'china-t': 0.15,
            'cjk': 0.15,
            'song': 0.15,
            'noto': 0.18,  # 泰文字体
            'thai': 0.18,  # 泰文字体
        }

        # 安全处理字体名称
        if fontname is None:
            fontname = 'helv'

        fontname_str = str(fontname).lower()

        # 检查是否包含特定字体关键词
        offset_ratio = 0.2  # 默认值
        for key, ratio in font_offset_ratios.items():
            if key in fontname_str:
                offset_ratio = ratio
                break

        return fontsize * offset_ratio


# 使用示例
def main():
    # 初始化翻译器
    translator = PDFTableTranslator(translate_api_key="sk-Jg3Rds726aZWc")

    translator.process_pdf_advanced(
        pdf_path="ASD01596_CL01_GY_OP0020_A.pdf",
        output_path="translated_output_th_v3.pdf",
        target_lang="th",
        auto_fit_font=True
    )

    print("PDF翻译完成！")


if __name__ == "__main__":
    main()
