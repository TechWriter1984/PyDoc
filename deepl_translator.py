import os
import deepl
import json
from typing import Optional, List, Dict, Any
import time

class DeepLTranslator:
    """
    DeepL翻译器类，用于处理文件翻译，支持上传、翻译和下载操作，以及术语库功能。
    """
    
    def __init__(self, auth_key: Optional[str] = None):
        """
        初始化DeepL翻译器
        
        Args:
            auth_key: DeepL API认证密钥，如果不提供则尝试从环境变量获取
        """
        # 尝试从环境变量获取认证密钥
        if not auth_key:
            auth_key = os.getenv('DEEPL_AUTH_KEY')
        
        if not auth_key:
            raise ValueError("DeepL API认证密钥未提供，请设置环境变量DEEPL_AUTH_KEY或直接提供auth_key参数")
        
        # 创建DeepL客户端
        self.translator = deepl.Translator(auth_key)
        print("DeepL翻译器初始化成功")
    
    def translate_file(self, input_path: str, output_path: str, source_lang: Optional[str] = None, 
                      target_lang: str = 'EN-US', glossary_path: Optional[str] = None, reuse_glossary: bool = True) -> None:
        """
        翻译文件
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            source_lang: 源语言代码（可选，DeepL会自动检测）
            target_lang: 目标语言代码，默认为美式英语'EN-US'
            glossary_path: 术语库JSON文件路径（可选）
            reuse_glossary: 是否复用现有的同名术语库（默认为True）
        """
        try:
            # 检查输入文件是否存在
            if not os.path.exists(input_path):
                raise FileNotFoundError(f"输入文件不存在: {input_path}")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 处理术语库
            glossary_id = None
            if glossary_path:
                glossary_id = self._get_or_create_glossary(glossary_path, source_lang, target_lang, reuse_glossary)
            
            print(f"开始翻译文件: {input_path}")
            print(f"目标语言: {target_lang}")
            if glossary_id:
                print(f"使用术语库进行翻译")
            
            # 规范化语言代码
            # DeepL API不再接受"EN"作为目标语言代码，必须使用"EN-GB"或"EN-US"
            if target_lang == 'EN':
                target_lang = 'EN-US'  # 默认为美式英语
            
            # 执行翻译
            # 当使用术语库时，必须提供源语言
            if glossary_id:
                # 确保源语言被设置
                if not source_lang:
                    source_lang = 'ZH'  # 如果未指定源语言，默认为中文
                
                result = self.translator.translate_document_from_filepath(
                    input_path,
                    output_path,
                    source_lang=source_lang,
                    target_lang=target_lang,
                    glossary=glossary_id
                )
            else:
                result = self.translator.translate_document_from_filepath(
                    input_path,
                    output_path,
                    source_lang=source_lang,
                    target_lang=target_lang
                )
            
            print(f"翻译完成！")
            # 不再尝试访问不存在的属性
            if source_lang:
                print(f"源语言: {source_lang}")
            print(f"翻译文件已保存至: {output_path}")
            
        except Exception as e:
            print(f"翻译过程中出错: {str(e)}")
            raise
    
    def _get_or_create_glossary(self, glossary_path: str, source_lang: Optional[str], target_lang: str, reuse_glossary: bool = True) -> str:
        """
        获取现有术语库或创建新的术语库
        
        Args:
            glossary_path: 术语库JSON文件路径
            source_lang: 源语言代码
            target_lang: 目标语言代码
            reuse_glossary: 是否复用现有的同名术语库
            
        Returns:
            术语库ID
        """
        try:
            # 创建术语库名称
            glossary_name = f"temp_glossary_{os.path.basename(glossary_path).split('.')[0]}"
            
            # 如果设置了复用术语库，尝试查找现有术语库
            if reuse_glossary:
                try:
                    # 获取所有术语库
                    glossaries = self.translator.list_glossaries()
                    
                    # 查找同名术语库
                    for glossary in glossaries:
                        if glossary.name == glossary_name:
                            print(f"找到现有术语库: {glossary_name}")
                            print(f"源语言: {glossary.source_lang}, 目标语言: {glossary.target_lang}")
                            return glossary.glossary_id
                    print(f"未找到现有术语库，将创建新术语库")
                except Exception as e:
                    print(f"查找现有术语库时出错: {str(e)}")
                    print("将创建新术语库")
            
            # 读取术语库JSON文件
            with open(glossary_path, 'r', encoding='utf-8') as f:
                glossary_data = json.load(f)
            
            # 验证术语库格式
            if not isinstance(glossary_data, list):
                raise ValueError("术语库文件必须包含JSON数组")
            
            # 转换术语库格式
            source_terms = []
            target_terms = []
            
            # 支持两种格式：
            # 1. ["术语1", "术语2", ...] (假设源语言为中文，目标语言为英文)
            # 2. [{"source": "术语1", "target": "term1"}, ...]
            
            for item in glossary_data:
                if isinstance(item, str):
                    # 第一种格式
                    source_terms.append(item)
                    # 尝试从项目中提取英文术语（如果有）
                    # 这里做一个简单的检查，如果是字典格式
                    if isinstance(glossary_data, list) and glossary_data.index(item) < len(glossary_data) - 1:
                        next_item = glossary_data[glossary_data.index(item) + 1]
                        if isinstance(next_item, str):
                            target_terms.append(next_item)
                    else:
                        target_terms.append(item)  # 默认使用相同术语
                elif isinstance(item, dict):
                    # 第二种格式
                    if "source" in item and "target" in item:
                        source_terms.append(item["source"])
                        target_terms.append(item["target"])
                    elif "acronym" in item and "expanded_form" in item:
                        # 支持与现有glossary格式兼容
                        source_terms.append(item["acronym"])
                        target_terms.append(item["expanded_form"])
            
            if not source_terms or not target_terms:
                raise ValueError("术语库文件不包含有效的术语对")
            
            # 如果未指定源语言，默认为中文
            if not source_lang:
                source_lang = 'ZH'
            
            print(f"创建术语库: {glossary_name}")
            print(f"源语言: {source_lang}, 目标语言: {target_lang}")
            print(f"术语数量: {len(source_terms)}")
            
            # 规范化语言代码
            if target_lang == 'EN':
                target_lang = 'EN-US'
            
            # 创建术语库
            glossary = self.translator.create_glossary(
                name=glossary_name,
                source_lang=source_lang,
                target_lang=target_lang,
                entries=dict(zip(source_terms, target_terms))
                )
            
            # 添加延迟，避免API调用过于频繁
            time.sleep(1)
            
            return glossary.glossary_id
            
        except Exception as e:
            print(f"创建术语库时出错: {str(e)}")
            
            # 如果是配额超限错误，尝试查找现有术语库
            if "QuotaExceededException" in str(type(e)) or "Too many glossaries" in str(e):
                print("术语库配额已超限，尝试使用现有术语库")
                try:
                    glossaries = self.translator.list_glossaries()
                    glossary_name = f"temp_glossary_{os.path.basename(glossary_path).split('.')[0]}"
                    
                    for glossary in glossaries:
                        if glossary.name == glossary_name:
                            print(f"找到现有术语库: {glossary_name}")
                            return glossary.glossary_id
                    
                    # 如果没有找到同名术语库，尝试使用第一个可用的术语库
                    if glossaries:
                        print(f"未找到同名术语库，使用第一个可用术语库")
                        return glossaries[0].glossary_id
                except Exception as inner_e:
                    print(f"查找现有术语库时出错: {str(inner_e)}")
            
            raise
    
    def get_supported_languages(self) -> Dict[str, List[str]]:
        """
        获取DeepL支持的所有语言
        
        Returns:
            包含源语言和目标语言的字典
        """
        try:
            source_langs = []
            target_langs = []
            
            # 获取所有支持的语言
            for lang in self.translator.get_source_languages():
                source_langs.append(f"{lang.code} - {lang.name}")
            
            for lang in self.translator.get_target_languages():
                target_langs.append(f"{lang.code} - {lang.name}")
            
            return {
                "source_languages": source_langs,
                "target_languages": target_langs
            }
            
        except Exception as e:
            print(f"获取支持的语言时出错: {str(e)}")
            return {"source_languages": [], "target_languages": []}
    
    def get_supported_formats(self) -> List[str]:
        """
        获取DeepL支持的文件格式
        
        Returns:
            支持的文件格式列表
        """
        # DeepL支持的文档格式
        return [
            ".docx", ".pdf", ".pptx", ".txt", ".xlsx",
            ".doc", ".rtf", ".odt", ".ods", ".odp"
        ]
        
    def list_glossaries(self) -> List[Any]:
        """
        列出所有可用的术语库
        
        Returns:
            术语库列表
        """
        try:
            glossaries = self.translator.list_glossaries()
            print(f"找到 {len(glossaries)} 个可用术语库")
            for i, glossary in enumerate(glossaries, 1):
                print(f"{i}. 名称: {glossary.name}")
                print(f"   ID: {glossary.glossary_id}")
                print(f"   源语言: {glossary.source_lang}")
                print(f"   目标语言: {glossary.target_lang}")
                print(f"   创建时间: {glossary.creation_time}")
            return glossaries
        except Exception as e:
            print(f"列出术语库时出错: {str(e)}")
            return []
            
    def delete_all_glossaries(self) -> None:
        """
        删除所有术语库（谨慎使用）
        """
        try:
            glossaries = self.translator.list_glossaries()
            print(f"准备删除 {len(glossaries)} 个术语库")
            
            for glossary in glossaries:
                self.translator.delete_glossary(glossary.glossary_id)
                print(f"已删除术语库: {glossary.name}")
                time.sleep(0.5)  # 添加延迟避免API调用过于频繁
            
            print("所有术语库已删除")
        except Exception as e:
            print(f"删除术语库时出错: {str(e)}")