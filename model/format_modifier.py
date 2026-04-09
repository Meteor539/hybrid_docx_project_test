import os
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import Dict
import re
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING

class FormatModifier:
    def __init__(self):
        """初始化格式修改器"""

        self.doc = None
        self.sections = { # 各部分具体解析内容
            "cover": {
                "school": None, # 武汉理工大学毕业设计（论文） -- 华文中宋 一号 居中
                "title": None, # 论文题目 -- 黑体 二号 居中
                "personal_information":None # 院（系）名称、专业班级、学生姓名、指导教师、标题 --  华文中宋 三号
            },
            
            "statement": {
                "title": None, # 学位论文原创性声明 -- 黑体 小二
                "content": None # 声明内容 -- 宋体 小四
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
                "title": [], # 各章标题 -- 黑体 小二
                "level1": [], # 一级标题 -- 黑体 三号
                "level2": [], # 二级标题 -- 黑体 四号
                "level3": [] # 三级标题 -- 黑体 小四
            },

            "figures_or_tables_title": [], # 图|表题 -- 黑体 小四 居中
            
            "figures": [], # 图
            
            "tables": [], # 表 内容 - 宋体 小四
            
            "references": {
                "title": None, # 参考文献标题 - 黑体 小二
                "references": [] # 参考文献内容 - 宋体五号
            },
            
            "acknowledgments": {
                "title": None, # 致谢标题 黑体 小二 
                "content": None # 致谢内容 宋体小四
            }
        }
        # 字体映射
        self.font_mapping = {
            "华文中宋": "STZhongsong",
            "宋体": "SimSun",
            "黑体": "SimHei",
            "楷体": "KaiTi",
            "仿宋": "FangSong",
            "微软雅黑": "Microsoft YaHei",
            "等线": "DengXian",
            "Times New Roman": "Times New Roman",
            "Arial": "Arial"
        }
        
        # 字号映射 (只保留 pt 值供内部使用)
        self.size_mapping = {
            "六号 (7.5pt)": 7.5, "小五 (9pt)": 9, "五号 (10.5pt)": 10.5, 
            "小四号 (12pt)": 12, "四号 (14pt)": 14, "小三号 (15pt)": 15, 
            "三号 (16pt)": 16, "小二号 (18pt)": 18, "二号 (22pt)": 22, 
            "小一号 (24pt)": 24, "一号 (26pt)": 26, "小初号 (36pt)": 36, 
            "初号 (42pt)": 42,
            "8pt": 8, "9pt": 9, "10pt": 10, "11pt": 11, "12pt": 12, "14pt": 14, 
            "16pt": 16, "18pt": 18, "20pt": 20, "22pt": 22, "24pt": 24, "26pt": 26, 
            "28pt": 28, "36pt": 36, "48pt": 48, "72pt": 72
        }
        
        # 对齐方式映射
        self.alignment_mapping = {
            "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
            "居中": WD_ALIGN_PARAGRAPH.CENTER,
            "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
            "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "分散对齐": WD_ALIGN_PARAGRAPH.DISTRIBUTE
        }

        # 表格样式映射
        self.table_style_mapping = {
            "无样式": None,
            "表格网格": "Table Grid",
            "浅色底纹": "Light Shading",
            "浅色列表": "Light List",
            "浅色网格": "Light Grid",
            "中度底纹 - 强调文字颜色 1": "Medium Shading 1 Accent 1",
            "中度列表 - 强调文字颜色 2": "Medium List 2 Accent 2",
            "中度网格 - 强调文字颜色 3": "Medium Grid 3 Accent 3",
            "深色列表": "Dark List"
        }

        self.line_spacing_rule_mapping = {
            "0.5倍行距": WD_LINE_SPACING.MULTIPLE,
            "单倍行距": WD_LINE_SPACING.MULTIPLE,
            "1.5倍行距": WD_LINE_SPACING.MULTIPLE,
            "2倍行距": WD_LINE_SPACING.MULTIPLE,
            "3倍行距": WD_LINE_SPACING.MULTIPLE,
            "固定值12pt": WD_LINE_SPACING.EXACTLY, 
            "固定值15pt": WD_LINE_SPACING.EXACTLY, 
            "固定值20pt": WD_LINE_SPACING.EXACTLY
        }

        self.line_spacing_mapping = {
            "0.5倍行距": 0.5,
            "单倍行距": 1,
            "1.5倍行距": 1.5,
            "2倍行距": 2,
            "3倍行距": 3,
            "固定值12pt": 12, 
            "固定值15pt": 15, 
            "固定值20pt": 20
        }
        
        # 默认格式设置
        self.format_rules = {
            "cover": {
                "school": {
                    "font": "华文中宋",
                    "size": "一号 (26pt)",
                    "alignment": "居中",
                    "line_spacing": "3倍行距"
                },
                "title": {
                    "font": "黑体",
                    "size": "二号 (22pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "personal_information": {
                    "font": "华文中宋",
                    "size": "三号 (16pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                }
            },

            "statement": {
                "title": {
                    "font": "黑体",
                    "size": "小二 (18pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "content": {
                    "font": "宋体",
                    "size": "小四 (12pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                }
            },

            "abstract_keyword": {
                "chinese_title": {
                    "font": "黑体",
                    "size": "小二 (18pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "chinese_content": {
                    "font": "宋体",
                    "size": "小四 (12pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                },
                "chinese_keyword_title": {
                    "font": "黑体",
                    "size": "四号 (14pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                },
                "chinese_keyword": {
                    "font": "宋体",
                    "size": "小四 (12pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                },
                "english_title": {
                    "font": "Times New Roman",
                    "size": "小二 (18pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "english_content": {
                    "font": "Times New Roman",
                    "size": "小四 (12pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                },
                "english_keyword_title": {
                    "font": "Times New Roman",
                    "size": "四号 (14pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                },
                "english_keyword": {
                    "font": "Times New Roman",
                    "size": "小四 (12pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                }
            },

            "headings": {
                "chapter": {
                    "font": "黑体",
                    "size": "小二 (18pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "level1": {
                    "font": "黑体",
                    "size": "三号 (16pt)",
                    "alignment": "左对齐",
                    "line_spacing": "固定值20pt"
                },
                "level2": {
                    "font": "黑体",
                    "size": "四号 (14pt)",
                    "alignment": "左对齐",
                    "line_spacing": "固定值20pt"
                },
                "level3": {
                    "font": "黑体",
                    "size": "小四 (12pt)",
                    "alignment": "左对齐",
                    "line_spacing": "固定值20pt"
                }
            },

            "main_text": {
                "font": "宋体",
                "size": "小四 (12pt)",
                "alignment": "两端对齐",
                "line_spacing": "固定值20pt"
            },

            "figures_or_tables_title": {
                "font": "宋体",
                "size": "小四 (12pt)",
                "alignment": "居中",
                "line_spacing": "固定值20pt"
            },

            "references": {
                "title": {
                    "font": "黑体",
                    "size": "小二 (18pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "content": {
                    "font": "宋体",
                    "size": "五号 (10.5pt)",
                    "alignment": "左对齐",
                    "line_spacing": "固定值20pt"
                }
            },

            "acknowledgments": {
                "title": {
                    "font": "黑体",
                    "size": "小二 (18pt)",
                    "alignment": "居中",
                    "line_spacing": "固定值20pt"
                },
                "content": {
                    "font": "宋体",
                    "size": "小四 (12pt)",
                    "alignment": "两端对齐",
                    "line_spacing": "固定值20pt"
                }
            }
        }
        
        
    def _get_font_size_pt(self, size_option: str) -> float | None:
        """从字号选项中提取Pt值"""
        if size_option in self.size_mapping:
            return self.size_mapping[size_option]
        # 尝试用正则表达式提取数字
        match = re.search(r'(\d+(\.\d+)?)', size_option)
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                return None
        return None

    def update_format(self, section_key, settings, prefix):
        font_key = f"{prefix}_font"
        size_key = f"{prefix}_size"
        align_key = f"{prefix}_align"
        line_spacing_key = f"{prefix}_line_spacing"

        if font_key in settings:
            self.format_rules[section_key][prefix]["font"] = settings[font_key]
        
        if size_key in settings:
            self.format_rules[section_key][prefix]["size"] = settings[size_key]

        if align_key in settings:
            self.format_rules[section_key][prefix]["alignment"] = settings[align_key]
        
        if line_spacing_key in settings:
            self.format_rules[section_key][prefix]["line_spacing"] = settings[line_spacing_key]
        
        
    def update_formats(self, formats: Dict[str, Dict[str, str]]) -> None:
        """
        更新格式设置
        
        Args:
            formats: 格式设置字典
        """
        for section_name, settings in formats.items():
            section_key = self.get_section_key(section_name)
            
            if section_key == "cover":
                # 更新封面格式
                self.update_format(section_key, settings, "school")
                    
                self.update_format(section_key, settings, "title")
                
                self.update_format(section_key, settings, "personal_information")
            
                    
            elif section_key == "statement":
                # 更新声明格式
                self.update_format(section_key, settings, "title")
                    
                self.update_format(section_key, settings, "content")
                    
            elif section_key == "abstract_keyword":
                # 更新摘要和关键词格式
                # 中文部分
                self.update_format(section_key, settings, "chinese_title")
                    
                self.update_format(section_key, settings, "chinese_content")
                    
                self.update_format(section_key, settings, "chinese_keyword_title")
                    
                self.update_format(section_key, settings, "chinese_keyword")
                    
                # 英文部分
                self.update_format(section_key, settings, "english_title")
                    
                self.update_format(section_key, settings, "english_content")
                    
                self.update_format(section_key, settings, "english_keyword_title")
                    
                self.update_format(section_key, settings, "english_keyword")
                    
            elif section_key == "main_text":
                # 更新正文格式
                if "font" in settings:
                    self.format_rules[section_key]["font"] = settings["font"]
                if "size" in settings:
                    self.format_rules[section_key]["size"] = settings["size"]
                if "align" in settings:
                    self.format_rules[section_key]["alignment"] = settings["align"]
                if "line_spacing" in settings:
                    self.format_rules[section_key]["line_spacing"] = settings["line_spacing"]
                    
            elif section_key == "headings":
                # 更新标题格式
                self.update_format(section_key, settings, "chapter")
                    
                for level in range(1, 4):
                    self.update_format(section_key, settings, f"level{level}")
                    
            elif section_key == "figures_or_tables_title":
                # 更新图|表标题格式
                if "font" in settings:
                    self.format_rules[section_key]["font"] = settings["font"]
                if "size" in settings:
                    self.format_rules[section_key]["size"] = settings["size"]
                if "align" in settings:
                    self.format_rules[section_key]["alignment"] = settings["align"]
                if "line_spacing" in settings:
                    self.format_rules[section_key]["line_spacing"] = settings["line_spacing"]
                    
            elif section_key == "references":
                # 更新参考文献格式
                self.update_format(section_key, settings, "title")
                    
                self.update_format(section_key, settings, "content")
                    
            elif section_key == "acknowledgments":
                # 更新致谢格式
                self.update_format(section_key, settings, "title")
                    
                self.update_format(section_key, settings, "content")
    
    def get_section_key(self, section_name: str) -> str:
        """
        根据部分名称获取键
        
        Args:
            section_name: 部分名称
            
        Returns:
            str: 部分键
        """
        section_mapping = {
            "封面": "cover",
            "声明": "statement",
            "摘要与关键词": "abstract_keyword",
            "目录": "catalogue",
            "正文": "main_text",
            "标题": "headings",
            "图|表题": "figures_or_tables_title",
            "图": "figures",
            "表格": "tables",
            "参考文献": "references",
            "致谢": "acknowledgments"
        }
        return section_mapping.get(section_name, section_name)
    
    def modify_format(self, file_path: str) -> str:
        """
        修改文档格式
        
        Args:
            file_path: 文档路径
            
        Returns:
            str: 修改后的文档路径
        """

        self._modify_section_format(self.sections["cover"], self.format_rules["cover"])

        # 修改封面格式
        self._modify_section_format(self.sections["cover"], self.format_rules["cover"])
        
        # 修改原创性声明格式
        self._modify_section_format(self.sections["statement"], self.format_rules["statement"])
        
        # 修改摘要和关键词格式
        self._modify_section_format(self.sections["abstract_keyword"], self.format_rules["abstract_keyword"])
        
        # 修改正文格式
        self._modify_paragraphs_format(self.sections["main_text"], self.format_rules["main_text"])
        
        # 修改标题格式
        self._modify_section_format(self.sections["headings"], self.format_rules["headings"])

        # 修改图|表题格式
        self._modify_paragraphs_format(self.sections["figures_or_tables_title"], self.format_rules["figures_or_tables_title"])
        
        # 修改参考文献格式
        self._modify_section_format(self.sections["references"], self.format_rules["references"])
        
        # 修改致谢格式
        self._modify_section_format(self.sections["acknowledgments"], self.format_rules["acknowledgments"])

        # 保存修改后的文档
        dir_path = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        new_file_name = f"formatted_{file_name}"
        new_file_path = os.path.join(dir_path, new_file_name)
        
        self.doc.save(new_file_path)
        return new_file_path
    
    def _modify_section_format(self, section_content, section_rules):
        """
        修改该部分格式
        
        Args:
            section_content: 内容
            section_rules: 格式规则
        """
        
        for key, content in section_content.items():
            if content is None:
                continue

            if isinstance(content, list):
                self._modify_paragraphs_format(content, section_rules[key])
            else:
                self._modify_paragraph_format(content, section_rules[key])

    def _modify_paragraph_format(self, paragraph, rules):
        """
        修改段落格式
        
        Args:
            paragraph: 段落
            rules: 格式规则
        """
        
        for run in paragraph.runs:
            # 设置字体
            if "font" in rules:
                run.font.name = self.font_mapping.get(rules["font"], rules["font"])
                
            
            # 设置字号
            if "size" in rules:
                size_pt = self._get_font_size_pt(rules["size"])
                if size_pt is not None:
                    run.font.size = Pt(size_pt)

        # 设置对齐方式
        if "alignment" in rules:
            paragraph.alignment = self.alignment_mapping[rules["alignment"]]
        
        # 设置行距
        if "line_spacing" in rules:
            line_spacing_rule = self.line_spacing_rule_mapping[rules["line_spacing"]]
            line_spacing = self.line_spacing_mapping[rules["line_spacing"]]

            if line_spacing is not None and line_spacing_rule is not None:
                paragraph.paragraph_format.line_spacing_rule = line_spacing_rule
                paragraph.paragraph_format.line_spacing = Pt(line_spacing) if line_spacing_rule == WD_LINE_SPACING.EXACTLY else line_spacing

    def _modify_paragraphs_format(self, paragraphs, rules):
        """
        修改多个段落格式
        
        Args:
            paragraph: 段落
            rules: 格式规则
        """
        
        for paragraph in paragraphs:
            self._modify_paragraph_format(paragraph, rules)
        