import re
from docx import Document
from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtCore import Qt

class DocumentParser:
    def __init__(self):
        """初始化文档解析器"""
        self.doc = None # 存储当前打开的Word文档对象
        self.current_para_index = 0 # 当前遍历到的段落位置
        self.sections = { # 各部分具体解析内容
            "cover": {
                "school": None, # 武汉理工大学毕业设计（论文） -- 华文中宋 一号 居中
                "title": None, # 论文题目 -- 黑体 二号 居中
                "personal_information":None # 院（系）名称、专业班级、学生姓名、指导教师、标题 --  华文中宋 三号
            },
            
            "statement": {
                "title": [], # 学位论文原创性声明 -- 黑体 小二
                "content": [] # 声明内容 -- 宋体 小四
            },
            
            "abstract_keyword": {
                "chinese_title": None, # 中文摘要标题 -- 黑体 小二
                "chinese_content": None, # 中文摘要内容 -- 宋体 小四
                "chinese_keyword_title": None, # 中文关键词标题 -- 黑体 四号
                "chinese_keyword": None, # 中文关键词 -- 宋体 小四
                "english_title": None, # 英文摘要标题 -- Times New Roman 粗体 小二
                "english_content": None, # 英文摘要内容 Times New Roman 小四
                "english_keyword_title": None, # 英文关键词标题 -- Times New Roman 粗体 四号
                "english_keyword": None, # 英文关键词 -- Times New Roman 小四号
            },
            
            "catalogue": {
                "title": None, # 目录标题 黑体 小二 
                "content": None # 目录内容 宋体小四
            },
            
            "main_text": [], # 正文、表题、图题 -- 宋体 小四
            
            "headings": {
                "chapter": [], # 各章标题 -- 黑体 小二
                "level1": [], # 一级标题 -- 黑体 三号
                "level2": [], # 二级标题 -- 黑体 四号
                "level3": [] # 三级标题 -- 黑体 小四
            },

            "figures_or_tables_title": [], # 图|表题 -- 黑体 小四 居中
            
            "figures": [], # 图
            
            "tables": [], # 表 内容 - 宋体 小四
            
            "references": {
                "title": None, # 参考文献标题 - 黑体 小二
                "content": [] # 参考文献内容 - 宋体五号
            },
            
            "acknowledgments": {
                "title": None, # 致谢标题 黑体 小二 
                "content": None # 致谢内容 宋体小四
            }
        }

        self.part_keywords = {
            "cover": ["大学毕业设计（论文）", "大学毕业论文（设计）","大学毕业设计", "大学毕业论文"],
            "statement1": ["原创性声明"],
            "statement2": ["使用授权书"],
            "chinese_abstract": ["摘要"],
            "english_abstract": ["abstract"],
            "catalogue": ["目录"],
            "main_text": ["第1章"],
            "references": ["参考文献", "参考书目", "references", "bibliography"],
            "acknowledgments": ["致谢", "致谢词"]
        }

        self.parts_order = []

        self.message_box = None # 消息框对象

    def print_sections_content(self):
        """输出所有section的文本内容，用于检查解析准确性"""
        '''
        print("=== 文档各部分内容 ===")
        
        # 封面部分
        print("\n=== 封面 ===")
        if self.sections["cover"]["school"]:
            print("学校名称:", self.sections["cover"]["school"].text)
        if self.sections["cover"]["title"]:
            print("论文题目:", self.sections["cover"]["title"].text)
        if self.sections["cover"]["personal_information"]:
            print("个人信息:")
            for info in self.sections["cover"]["personal_information"]:
                print("  -", info.text)
        
        # 声明部分
        print("\n=== 原创性声明 ===")
        if self.sections["statement"]["title"]:
            print("标题:")
            for para in self.sections["statement"]["title"]:
                print("  -", para.text)
        if self.sections["statement"]["content"]:
            print("内容:")
            for para in self.sections["statement"]["content"]:
                print("  -", para.text)
        
        # 摘要和关键词部分
        print("\n=== 摘要与关键词 ===")
        if self.sections["abstract_keyword"]["chinese_title"]:
            print("中文摘要标题:", self.sections["abstract_keyword"]["chinese_title"].text)
        if self.sections["abstract_keyword"]["chinese_content"]:
            print("中文摘要内容:")
            for para in self.sections["abstract_keyword"]["chinese_content"]:
                print("  -", para.text)
        if self.sections["abstract_keyword"]["chinese_keyword_title"]:
            print("中文关键词标题:", self.sections["abstract_keyword"]["chinese_keyword_title"].text)
        if self.sections["abstract_keyword"]["chinese_keyword"]:
            print("中文关键词:")
            print("  -", self.sections["abstract_keyword"]["chinese_keyword"].text)
        
        # 英文部分
        if self.sections["abstract_keyword"]["english_title"]:
            print("\n英文摘要标题:", self.sections["abstract_keyword"]["english_title"].text)
        if self.sections["abstract_keyword"]["english_content"]:
            print("英文摘要内容:")
            for para in self.sections["abstract_keyword"]["english_content"]:
                print("  -", para.text)
        if self.sections["abstract_keyword"]["english_keyword_title"]:
            print("英文关键词标题:", self.sections["abstract_keyword"]["english_keyword_title"].text)
        if self.sections["abstract_keyword"]["english_keyword"]:
            print("英文关键词:")
            print("  -", self.sections["abstract_keyword"]["english_keyword"].text)
        
        # 目录部分
        print("\n=== 目录 ===")
        if self.sections["catalogue"]["title"]:
            print("标题:", self.sections["catalogue"]["title"].text)
        if self.sections["catalogue"]["content"]:
            print("内容:")
            for para in self.sections["catalogue"]["content"]:
                print("  -", para.text)
        
        # 正文部分
        print("\n=== 正文 ===")
        if self.sections["main_text"]:
            print("正文内容:")
            for para in self.sections["main_text"]:
                print("  -", para.text)
        
        # 标题部分
        print("\n=== 标题 ===")
        if self.sections["headings"]["title"]:
            print("章节标题:")
            for para in self.sections["headings"]["title"]:
                print("  -", para.text)
        if self.sections["headings"]["level1"]:
            print("一级标题:")
            for para in self.sections["headings"]["level1"]:
                print("  -", para.text)
        if self.sections["headings"]["level2"]:
            print("二级标题:")
            for para in self.sections["headings"]["level2"]:
                print("  -", para.text)
        if self.sections["headings"]["level3"]:
            print("三级标题:")
            for para in self.sections["headings"]["level3"]:
                print("  -", para.text)

        # 图表题部分
        print("\n=== 图表题 ===")
        if self.sections["figures_or_tables_title"]:
            print("图表标题:")
            for para in self.sections["figures_or_tables_title"]:
                print("  -", para.text)
        
        # 图表部分
        print("\n=== 图表 ===")
        if self.sections["figures"]:
            print("图:")
            for para in self.sections["figures"]:
                print("  -", para)
        if self.sections["tables"]:
            print("表:")
            for para in self.sections["tables"]:
                print("  -", para)
        
        # 参考文献部分
        print("\n=== 参考文献 ===")
        if self.sections["references"]["title"]:
            print("标题:", self.sections["references"]["title"].text)
        if self.sections["references"]["references"]:
            print("参考文献列表:")
            for para in self.sections["references"]["references"]:
                print("  -", para.text)
        
        # 致谢部分
        print("\n=== 致谢 ===")
        if self.sections["acknowledgments"]["title"]:
            print("标题:", self.sections["acknowledgments"]["title"].text)
        if self.sections["acknowledgments"]["content"]:
            print("内容:")
            for para in self.sections["acknowledgments"]["content"]:
                print("  -", para.text)    
        '''
    def check_next_part(self, para):
        """检查当前段落是否是下一部分的开始"""
        for part, keywords in self.part_keywords.items():
            if part == "main_text" and any(para.text.strip().replace(" ", "").lower().startswith(keyword) for keyword in keywords):
                return part

            elif any(para.text.strip().replace(" ", "").lower().endswith(keyword) for keyword in keywords):
                return part
        return None
            
    def process_next_part(self, next_part):
        """处理下一部分的解析"""
        if next_part is None:
            return
        
        self.parts_order.append(next_part)
        
        if next_part == "cover":
            self._parse_cover()
        elif next_part == "statement1":
            self._parse_statement1()
        elif next_part == "statement2":
            self._parse_statement2()
        elif next_part == "chinese_abstract":
            self._parse_chinese_abstract()
        elif next_part == "english_abstract":
            self._pase_english_abstract()
        elif next_part == "catalogue":
            self._parse_catalogue()
        elif next_part == "main_text":
            self._parse_main_text()
        elif next_part == "references":
            self._parse_references()
        elif next_part == "acknowledgments":
            self._parse_acknowledgments()

    def show_message(self, message_text):
        """显示消息框"""
        self.message_box = QMessageBox()
        self.message_box.setWindowTitle("排版顺序检查")
        self.message_box.setText(message_text)
        self.message_box.setIcon(QMessageBox.Icon.Information)
        self.message_box.setMinimumSize(800, 800)
        self.message_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        # 不阻塞程序往下执行
        self.message_box.setWindowModality(Qt.WindowModality.NonModal)  # 设置为非模态
        self.message_box.show()

    def check_order(self):
        """检查各部分顺序是否符合要求"""
        # 定义正确的几种顺序
        correct_orders = [
            "cover",
            "statement1",
            "statement2",
            "chinese_abstract",
            "english_abstract",
            # "catalogue",
            "main_text",
            "references",
            "acknowledgments"
        ]

        PART_NAME_MAPPING = {
            "cover": "封面",
            "statement1": "原创性声明",
            "statement2": "使用授权声明",
            "chinese_abstract": "中文摘要",
            "english_abstract": "英文摘要",
            # "catalogue": "目录",
            "main_text": "正文",
            "references": "参考文献",
            "acknowledgments": "致谢"
        }

        # 检查是否缺失某些部分
        missing_parts = [part for part in correct_orders if part not in self.parts_order]
        if missing_parts:
            missing_parts_str = ", ".join([PART_NAME_MAPPING.get(part, part) for part in missing_parts])
            self.show_message( f"可能缺失的部分: {missing_parts_str}")
            return
        
        # 检查顺序是否正确
        for i in range(len(self.parts_order) - 1):
            current_part = self.parts_order[i]
            next_part = self.parts_order[i + 1]

            # 检查当前部分是否在正确顺序中
            if current_part not in correct_orders:
                self.show_message(f"可能非法的部分: {PART_NAME_MAPPING.get(current_part, current_part)}")
                return
            
            # 检查下一部分是否在正确顺序中
            if next_part not in correct_orders:
                self.show_message(f"可能非法的部分: {PART_NAME_MAPPING.get(next_part, next_part)}")
                return

            # 检查当前部分和下一部分的顺序
            if correct_orders.index(current_part) > correct_orders.index(next_part):
                self.show_message(f"可能错误的排版顺序: {PART_NAME_MAPPING.get(current_part, current_part)} -> {PART_NAME_MAPPING.get(next_part, next_part)}")
                return
        
        # 如果顺序正确，显示成功消息
        self.show_message("排版顺序检查通过, 各部分排版顺序正确")
            
        

        


    def parse_document(self, file_path):
        """解析Word文档"""
        self.current_para_index = 0 # 重置遍历索引

        self.parts_order = [] # 重置部分顺序列表
        self.sections = { # 各部分具体解析内容
            "cover": {
                "school": None, # 武汉理工大学毕业设计（论文） -- 华文中宋 一号 居中
                "title": None, # 论文题目 -- 黑体 二号 居中
                "personal_information":None # 院（系）名称、专业班级、学生姓名、指导教师、标题 --  华文中宋 三号
            },
            
            "statement": {
                "title": [], # 学位论文原创性声明 -- 黑体 小二
                "content": [] # 声明内容 -- 宋体 小四
            },
            
            "abstract_keyword": {
                "chinese_title": None, # 中文摘要标题 -- 黑体 小二
                "chinese_content": None, # 中文摘要内容 -- 宋体 小四
                "chinese_keyword_title": None, # 中文关键词标题 -- 黑体 四号
                "chinese_keyword": None, # 中文关键词 -- 宋体 小四
                "english_title": None, # 英文摘要标题 -- Times New Roman 粗体 小二
                "english_content": None, # 英文摘要内容 Times New Roman 小四
                "english_keyword_title": None, # 英文关键词标题 -- Times New Roman 粗体 四号
                "english_keyword": None, # 英文关键词 -- Times New Roman 小四号
            },
            
            "catalogue": {
                "title": None, # 目录标题 黑体 小二 
                "content": None # 目录内容 宋体小四
            },
            
            "main_text": [], # 正文、表题、图题 -- 宋体 小四
            
            "headings": {
                "chapter": [], # 各章标题 -- 黑体 小二
                "level1": [], # 一级标题 -- 黑体 三号
                "level2": [], # 二级标题 -- 黑体 四号
                "level3": [] # 三级标题 -- 黑体 小四
            },

            "figures_or_tables_title": [], # 图|表题 -- 黑体 小四 居中
            
            "figures": [], # 图
            
            "tables": [], # 表 内容 - 宋体 小四
            
            "references": {
                "title": None, # 参考文献标题 - 黑体 小二
                "content": [] # 参考文献内容 - 宋体五号
            },
            
            "acknowledgments": {
                "title": None, # 致谢标题 黑体 小二 
                "content": None # 致谢内容 宋体小四
            }
        }

        # 初始化需要检测的部分
        self.part_keywords = {
            "cover": ["大学毕业设计（论文）", "大学毕业论文（设计）","大学毕业设计", "大学毕业论文"],
            "statement1": ["原创性声明"],
            "statement2": ["使用授权书"],
            "chinese_abstract": ["摘要"],
            "english_abstract": ["abstract"],
            "catalogue": ["目录"],
            "main_text": ["第1章"],
            "references": ["参考文献", "参考书目", "references", "bibliography"],
            "acknowledgments": ["致谢", "致谢词"]
        }

        next_part = None
        try:
            self.doc = Document(file_path)

            while self.current_para_index < len(self.doc.paragraphs):
                para = self.doc.paragraphs[self.current_para_index]

                next_part = self.check_next_part(para)
                if next_part:
                    self.process_next_part(next_part)
                    break

                self.current_para_index += 1

            if self.current_para_index >= len(self.doc.paragraphs) or len(self.part_keywords) == 0:
                print("已遍历完所有段落或没有更多部分可解析")

            self._parse_figures()
            self._parse_tables()
            # 解析完毕，返回各部分内容
            self.print_sections_content() # 调用输出函数
            # self.check_order()
            return self.sections
        
        except Exception as e:
            raise Exception(f"文档解析失败: {str(e)}")
    
    def _parse_cover(self):
        """解析封面部分"""
        if not self.doc.paragraphs:
            return
            
        # 从part_keywords中移除cover部分
        self.part_keywords.pop("cover", None)

        next_part = None

        # 获取封面的所有段落
        cover_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            cover_paragraphs.append(para)

        # 去掉所有空段落
        cover_paragraphs = [para for para in cover_paragraphs if para.text.strip()]
        
        if not cover_paragraphs:
            return    

        # 假设第一行是学校名称
        if cover_paragraphs:
            self.sections["cover"]["school"] = cover_paragraphs[0]
            
        # 假设第二行是论文题目
        if len(cover_paragraphs) > 1:
            self.sections["cover"]["title"] = cover_paragraphs[1]
            
        # 剩余内容为个人信息
        if len(cover_paragraphs) > 2:
            self.sections["cover"]["personal_information"] = cover_paragraphs[2:]

        # 进入下一部分
        self.process_next_part(next_part)

    def _parse_statement1(self):
        """解析原创性声明部分"""
        if not self.doc.paragraphs:
            return

        # 从part_keywords中移除statement部分
        self.part_keywords.pop("statement1", None)

        next_part = None

        # 获取声明的所有段落
        statement1_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            statement1_paragraphs.append(para)

        # 去掉所有空段落
        statement1_paragraphs = [para for para in statement1_paragraphs if para.text.strip()]
        
        if not statement1_paragraphs: 
            return
        
        self.sections["statement"]["title"].append(statement1_paragraphs[0])
        
        # 合并已有的content与剩下的段落
        if len(statement1_paragraphs) > 1:
            self.sections["statement"]["content"].extend(statement1_paragraphs[1:])
            
        # 进入下一部分
        self.process_next_part(next_part)

    def _parse_statement2(self):
        """解析使用授权书部分"""
        if not self.doc.paragraphs:
            return

        # 从part_keywords中移除statement部分
        self.part_keywords.pop("statement2", None)

        next_part = None

        # 获取声明的所有段落
        statement2_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            statement2_paragraphs.append(para)

        # 去掉所有空段落
        statement2_paragraphs = [para for para in statement2_paragraphs if para.text.strip()]
        
        if not statement2_paragraphs: 
            return
        
        self.sections["statement"]["title"].append(statement2_paragraphs[0])
        
        # 合并已有的content与剩下的段落
        if len(statement2_paragraphs) > 1:
            self.sections["statement"]["content"].extend(statement2_paragraphs[1:])
            
        # 进入下一部分
        self.process_next_part(next_part)
        
    def _parse_chinese_abstract(self):
        """解析中文摘要部分"""
        if not self.doc.paragraphs:
            return
        
        self.part_keywords.pop("chinese_abstract", None)

        next_part = None

        chinese_abstract_keyword_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            chinese_abstract_keyword_paragraphs.append(para)

        # 去掉所有空段落
        chinese_abstract_keyword_paragraphs = [para for para in chinese_abstract_keyword_paragraphs if para.text.strip()]
        if not chinese_abstract_keyword_paragraphs: 
            return
            
        
        self.sections["abstract_keyword"]["chinese_title"] = chinese_abstract_keyword_paragraphs[0]
        # 第二段到倒数第二段都是content
        if len(chinese_abstract_keyword_paragraphs) > 2:
            self.sections["abstract_keyword"]["chinese_content"] = chinese_abstract_keyword_paragraphs[1:-1]
            self.sections["abstract_keyword"]["chinese_keyword_title"] = chinese_abstract_keyword_paragraphs[-2]
        self.sections["abstract_keyword"]["chinese_keyword"] = chinese_abstract_keyword_paragraphs[-1]

        # 进入下一部分
        self.process_next_part(next_part)
        
    def _pase_english_abstract(self):
        """解析英文摘要部分"""
        if not self.doc.paragraphs:
            return
            
        self.part_keywords.pop("english_abstract", None)

        next_part = None

        english_abstract_keyword_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            english_abstract_keyword_paragraphs.append(para)

        # 去掉所有空段落
        english_abstract_keyword_paragraphs = [para for para in english_abstract_keyword_paragraphs if para.text.strip()]
        
        if not english_abstract_keyword_paragraphs: 
            return
            
        self.sections["abstract_keyword"]["english_title"] = english_abstract_keyword_paragraphs[0]
        # 第二段到倒数第二段都是content
        if len(english_abstract_keyword_paragraphs) > 2:
            self.sections["abstract_keyword"]["english_content"] = english_abstract_keyword_paragraphs[1:-2]
            self.sections["abstract_keyword"]["english_keyword_title"] = english_abstract_keyword_paragraphs[-2]
        self.sections["abstract_keyword"]["english_keyword"] = english_abstract_keyword_paragraphs[-1]

        # 进入下一部分
        self.process_next_part(next_part)
                
    def _parse_catalogue(self):
        """解析目录部分"""
        if not self.doc.paragraphs:
            return
        
        self.part_keywords.pop("catalogue", None)
        next_part = None

        catalogue_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            catalogue_paragraphs.append(para)

        # 去掉所有空段落
        catalogue_paragraphs = [para for para in catalogue_paragraphs if para.text.strip()]
                
        if not catalogue_paragraphs:
            return
            
        # 假设第一行是目录标题
        if catalogue_paragraphs:
            self.sections["catalogue"]["title"] = catalogue_paragraphs[0]
            
        # 剩余内容为目录内容
        if len(catalogue_paragraphs) > 1:
            self.sections["catalogue"]["content"] = catalogue_paragraphs[1:]

        # 进入下一部分
        self.process_next_part(next_part)
    
    def _is_chapter_title(self, text: str) -> bool:
        """检查是否是章节标题"""
        # 检查是否以"第X章"开头且后面有空格 第[任意数量数字]章[空白字符]
        has_comma = '，' not in text
        has_period = '。' not in text

        if (re.match(r'^第[一二三四五六七八九十]+章\s*', text) or re.match(r'^第\d+章\s*', text)) and has_comma and has_period:
            return True
        return False

    def _is_section_title(self, text: str) -> int:
        """检查是否是节标题，返回标题级别（1-3），如果不是则返回0"""
        # 一级标题模式：1.1 标题
        level1_patterns = [
            r'^\d+\.\d+\s*.+$',  # 匹配"1.1 标题"格式
            r'^[一二三四五六七八九十]+、\s*.+$'  # 匹配"一、标题"格式
        ]
        
        # 二级标题模式：1.1.1 标题、（一）标题
        level2_patterns = [
            r'^\d+\.\d+\.\d+\s*.+$',  # 匹配"1.1.1 标题"格式
            r'^（[一二三四五六七八九十]+）\s*.+$'  # 匹配"（一）标题"格式
        ]
        
        # 三级标题模式：1. 标题、1) 标题
        level3_patterns = [
            r'^\d+\.\s*.+$',  # 匹配"1. 标题"格式
            r'^\d+\)\s*.+$'  # 匹配"1) 标题"格式
        ]

        for pattern in level2_patterns:
            if re.match(pattern, text):
                return 2
            
        for pattern in level1_patterns:
            if re.match(pattern, text):
                return 1
                
        for pattern in level3_patterns:
            if re.match(pattern, text):
                return 3
                
        return 0

    def _is_figure_or_table_title(self, text: str) -> bool:
        """检查是否是图表题"""
        # 检查是否以"图X"或"表X"开头且后面有空格
        if re.match(r'^(图|表)\d+(\.\d+)*', text):
            return True
        return False

    def _parse_main_text(self):
        """解析正文部分"""
        if not self.doc.paragraphs:
            return
        
        self.part_keywords.pop("main_text", None)

        next_part = None

        # 获取正文的所有段落
        main_text_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            main_text_paragraphs.append(para)

        # 去掉所有空段落
        main_text_paragraphs = [para for para in main_text_paragraphs if para.text.strip()]

        if not main_text_paragraphs:
            return
        
        for para in main_text_paragraphs:
            if para.text.strip():
                text = para.text.strip()
                # 检查是否是章标题
                if self._is_chapter_title(text):
                    self.sections["headings"]["chapter"].append(para)
                    continue
                
                # 检查是否是节标题
                level = self._is_section_title(text)
                if level > 0:
                    if level == 1:
                        self.sections["headings"]["level1"].append(para)
                    elif level == 2:
                        self.sections["headings"]["level2"].append(para)
                    elif level == 3:
                        self.sections["headings"]["level3"].append(para)
                    continue
                
                # 检查是否是图表题
                if self._is_figure_or_table_title(text):
                    self.sections["figures_or_tables_title"].append(para)
                    continue
            
                self.sections["main_text"].append(para)

        # 进入下一部分
        self.process_next_part(next_part)
            

    def _parse_references(self):
        """解析参考文献部分"""
        if not self.doc.paragraphs:
            return
        
        self.part_keywords.pop("references", None)

        next_part = None

        # 获取参考文献的所有段落
        references_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            references_paragraphs.append(para)
        
        # 去掉所有空段落
        references_paragraphs = [para for para in references_paragraphs if para.text.strip()]

        if not references_paragraphs:
            return
            
        # 假设第一行是参考文献标题
        if references_paragraphs:
            self.sections["references"]["title"] = references_paragraphs[0]
            
        # 剩余内容为参考文献列表
        if len(references_paragraphs) > 1:
            self.sections["references"]["content"] = references_paragraphs[1:]

        # 进入下一部分
        self.process_next_part(next_part)
                
    def _parse_acknowledgments(self):
        """解析致谢部分"""
        if not self.doc.paragraphs:
            return
        
        self.part_keywords.pop("acknowledgments", None)

        next_part = None

        # 获取致谢的所有段落
        acknowledgments_paragraphs = []
        for i in range(self.current_para_index, len(self.doc.paragraphs)):
            para = self.doc.paragraphs[i]
            
            # 如果检测到下一部分的关键词，说明该部分结束
            next_part = self.check_next_part(para)
            if next_part:
                self.current_para_index = i
                break

            acknowledgments_paragraphs.append(para)
        
        # 去掉所有空段落
        acknowledgments_paragraphs = [para for para in acknowledgments_paragraphs if para.text.strip()]

        if not acknowledgments_paragraphs:
            return
            
        # 假设第一行是致谢标题
        if acknowledgments_paragraphs:
            self.sections["acknowledgments"]["title"] = acknowledgments_paragraphs[0]
            
        # 剩余内容为致谢内容
        if len(acknowledgments_paragraphs) > 1:
            self.sections["acknowledgments"]["content"] = acknowledgments_paragraphs[1:]

        # 进入下一部分
        self.process_next_part(next_part)
        
    def _parse_figures(self):
        """解析图片"""
        self.sections["figures"] = self.doc.inline_shapes
        
    def _parse_tables(self):
        """解析表格"""
        self.sections["tables"] = self.doc.tables