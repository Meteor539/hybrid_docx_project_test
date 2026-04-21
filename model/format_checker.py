from typing import Dict, Any
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re # 导入正则表达式模块
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn

class FormatChecker:
    def __init__(self):
        """初始化格式检查器"""
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
        
        # 字号映射 (仅用于比较时提取pt值)
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

        self.alignment_mapping = {
            "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
            "居中": WD_ALIGN_PARAGRAPH.CENTER,
            "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
            "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "分散对齐": WD_ALIGN_PARAGRAPH.DISTRIBUTE
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

        # 字体别名映射（用于降低 python-docx 字体名读取差异导致的误报）
        self.font_alias_groups = {
            "华文中宋": {"华文中宋", "STZhongsong", "ST中宋"},
            "宋体": {"宋体", "SimSun", "Songti"},
            "黑体": {"黑体", "SimHei", "Heiti"},
            "楷体": {"楷体", "KaiTi", "楷体_GB2312"},
            "仿宋": {"仿宋", "FangSong", "仿宋_GB2312"},
            "微软雅黑": {"微软雅黑", "Microsoft YaHei", "MSYH"},
            "等线": {"等线", "DengXian"},
            "Times New Roman": {"Times New Roman", "TimesNewRoman"},
            "Arial": {"Arial"},
        }
        self.font_alias_lookup = self._build_font_alias_lookup()

    def _get_font_size_pt(self, size_option: str) -> float | None:
        """从字号选项中提取Pt值"""
        if size_option in self.size_mapping:
            return self.size_mapping[size_option]
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
        
    def check_format(self, doc_content: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """
        检查文档格式
        
        Args:
            doc_content: 文档内容
            
        Returns:
            格式检查结果
        """
        results = {}
        
        # 检查封面格式
        results["cover"] = self._check_section_format(doc_content["cover"], self.format_rules["cover"])
        
        # 检查原创性声明格式
        results["statement"] = self._check_section_format(doc_content["statement"], self.format_rules["statement"])
        
        # 检查摘要和关键词格式
        results["abstract_keyword"] = self._check_section_format(doc_content["abstract_keyword"], self.format_rules["abstract_keyword"])
        
        # 检查正文格式
        results["main_text"] = self._check_paragraphs_format(doc_content["main_text"], self.format_rules["main_text"])
        
        # 检查标题格式
        results["headings"] = self._check_section_format(doc_content["headings"], self.format_rules["headings"])

        # 检查图|表题格式
        results["figures_or_tables_title"] = self._check_paragraphs_format(doc_content["figures_or_tables_title"], self.format_rules["figures_or_tables_title"])
        
        # 检查参考文献格式
        results["references"] = self._check_section_format(doc_content["references"], self.format_rules["references"])
        
        # 检查致谢格式
        results["acknowledgments"] = self._check_section_format(doc_content["acknowledgments"], self.format_rules["acknowledgments"])

        # 检查参考文献引用格式
        results["references_check"] = self.check_references_format(doc_content["references"]["content"])
        
        return results
    
    def check_references_format(self, references: list):
        results = []

        for reference in references:
            if not self._check_reference_format(reference.text):
                result = {
                    "参考文献": reference.text,
                    "检查结果": "引用格式不符合标准"
                }

                results.append(result)

        if results:
            return results
        return {"检查结果": "引用格式匹配无误"}

    def _check_reference_format(self, reference):
        """检查参考文献格式"""
        # 期刊文献： [序号] 作者.题目[J]. 刊名，年，卷号（期号）起止页码. 或年（期号）：起止页码.
        journal_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,\s*(\d+)\s*\,(.*?):\s*(.*?)\.\s*$' # 年，卷号（期号）:起止页码
        journal_pattern2 = r'^(.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,\s*(\d+)\s*\,(.*?):\s*(.*?)\.\s*$' # 无序号
        journal_pattern3 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,(.*?):(.*?)\.\s*$' # 年（期号）：起止页码
        journal_pattern4 = r'^(.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,(.*?):(.*?)\.\s*$' # 无序号
        journal_pattern5 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,\s*(\d+)\s*\,(.*?):(.*?)\.\s*$' # 年,(期号)：起止页码
        journal_pattern6 = r'^(.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,\s*(\d+)\s*\,(.*?):(.*?)\.\s*$' # 无序号

        # 专（译）著文献: [序号] 作者. 书名[M]. 译者. 出版地：出版者，出版年：起止页码.
        monograph_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*M\s*\]\s*\.(.*?)\.(.*?):(.*?)\,\s*(\d+)\s*:(.*?)\.\s*$'
        monograph_pattern2 = r'^(.*?)\.(.*?)\[\s*M\s*\]\s*\.(.*?)\.(.*?):(.*?)\,\s*(\d+)\s*:(.*?)\.\s*$'

        # 论文集文献：    [序号] 作者.论文集名称[C].出版地：出版者，出版年： 起止页码 
        proceedings_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*C\s*\]\s*\.(.*?):(.*?)\,\s*(\d+)\s*:(.*?)\.\s*$'
        proceedings_pattern2 = r'^(.*?)\.(.*?)\[\s*C\s*\]\s*\.(.*?):(.*?)\,\s*(\d+)\s*:(.*?)\.\s*$'
        
        # 会议录文献 ： [序号] 编者. 会议名称，会议地点，会议年份[C]. 出版地：出版者，出版年.
        conference_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\,(.*?)\,(.*?)\[\s*C\s*\]\s*\.(.*?):(.*?)\,\s*(\d+)\s*\.\s*$'
        conference_pattern2 = r'^(.*?)\.(.*?)\,(.*?)\,(.*?)\[\s*C\s*\]\s*\.(.*?):(.*?)\,\s*(\d+)\s*\.\s*$'

        # 学位论文文献：  [序号] 姓名. 题目［D］. 授予单位所在地：授予单位，授予年.
        thesis_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*D\s*\]\s*\.(.*?):(.*?)\,\s*(\d+)\s*\.\s*$'
        thesis_pattern2 = r'^(.*?)\.(.*?)\[\s*D\s*\]\s*\.(.*?):(.*?)\,\s*(\d+)\s*\.\s*$'

        # 专利文献: [序号] 申请人. 专利名. 国名，专利文献种类专利号[P].日期.
        patent_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\.(.*?)\,(.*?)(\[\s*P\s*\]\s*)\.(.*?)\.\s*$'
        patent_pattern2 = r'^(.*?)\.(.*?)\.(.*?)\,(.*?)(\[\s*P\s*\]\s*)\.(.*?)\.\s*$'

        # 技术标准文献: [序号] 发布单位. 技术标准代号. 技术标准名称[S]. 出版地：出版者，出版年
        standard_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\.(.*?)(\[\s*S\s*\]\s*)\.(.*?):(.*?)\,\s*(\d+)\s*\.\s*$'
        standard_pattern2 = r'^(.*?)\.(.*?)\.(.*?)(\[\s*S\s*\]\s*)\.(.*?):(.*?)\,\s*(\d+)\s*\.\s*$'

        # 电子文献 : 作者. 题目： 其他题目信息[DB、CP、EB / MT、DK、CD、OL]. 出版地： 出版者，出版年（更新或修改日期）[引用日期]. 获取和访问路径.
        electronic_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?):(.*?)\.(.*?):(.*?)\,\s*(\d+)(\s*\[(.*?)\])\s*\.(.*?)\.\s*$'
        electronic_pattern2 = r'^(.*?)\.(.*?):(.*?)\.(.*?):(.*?)\,\s*(\d+)(\s*\[(.*?)\])\s*\.(.*?)\.\s*$'

        # 专著中析出的文献:析出文献主要作者,析出文献题目[M]. 析出文献其他作者//专著主要作者. 专著题目：其他题目信息. 出版地. 出版者, 出版年：析出文献的页码[引用日期]. 获取和访问路径.
        monograph_extracted_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\,(.*?)(\[\s*M\s*\]\s*)\.(.*?)\.(.*?):(.*?)\.(.*?)\.(.*?)\,\s*(\d+)\s*:(.*?)(\s*\[(.*?)\])\s*\.(.*?)\.\s*$'
        monograph_extracted_pattern2 = r'^(.*?)\,(.*?)(\[\s*M\s*\]\s*)\.(.*?)\.(.*?):(.*?)\.(.*?)\.(.*?)\,\s*(\d+)\s*:(.*?)(\s*\[(.*?)\])\s*\.(.*?)\.\s*$'

        # 出版物析出的文献: [序号] 作者.文献题目[J]. 连续出版物题目，年，卷（期）：页码[引用日期].获取和访问路径.
        publication_extracted_pattern1 = r'^\s*\[\s*(\d+)\s*\](.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,\s*(\d+)\s*\,\s*(\d+\s*\(\s*\d+\s*\)\s*):(.*?)(\s*\[(.*?)\])\s*\.(.*?)\.\s*$'
        publication_extracted_pattern2 = r'^(.*?)\.(.*?)\[\s*J\s*\]\s*\.(.*?)\,\s*(\d+)\s*\,\s*(\d+\s*\(\s*\d+\s*\)\s*):(.*?)(\s*\[(.*?)\])\s*\.(.*?)\.\s*$'
        
        patterns = [
            journal_pattern1, journal_pattern2, journal_pattern3, journal_pattern4, journal_pattern5, journal_pattern6,
            monograph_pattern1, monograph_pattern2,
            proceedings_pattern1, proceedings_pattern2,
            conference_pattern1, conference_pattern2,
            thesis_pattern1, thesis_pattern2,
            patent_pattern1, patent_pattern2,
            standard_pattern1, standard_pattern2,
            electronic_pattern1, electronic_pattern2,
            monograph_extracted_pattern1, monograph_extracted_pattern2,
            publication_extracted_pattern1, publication_extracted_pattern2
        ]

        for pattern in patterns:
            if re.fullmatch(pattern, reference):
                return True
        return False
            
    def _check_section_format(self, section_content, section_rules):
        """检查某个部分的格式"""
        results = {}

        for key, content in section_content.items():
            if content is None:
                results[key] = {"检查结果": "该部分内容为空，无法进行格式检查"}
                continue

            if isinstance(content, list):
                results[key] = self._check_paragraphs_format(content, section_rules[key])
            else:
                results[key] = self._check_paragraph_format_when_section_check(content, section_rules[key])
                
        return results
        
    def _check_paragraphs_format(self, paragraphs, rules):
        """检查多个段落的格式"""
        results = [res for res in (self._check_paragraph_format(para, rules) for para in paragraphs) if res is not None]

        return results if results else {"检查结果": "格式匹配无误"}
        

        
    def _check_paragraph_format(self, paragraph, rules):
        """检查单个段落的格式"""
        font_result = self._check_font(paragraph, rules.get("font"))
        size_result = self._check_size(paragraph, rules.get("size"))
        align_result = self._check_alignment(paragraph, rules.get("alignment"))
        line_spacing_result = self._check_line_spacing(paragraph, rules.get("line_spacing"))

        if not (font_result and size_result and align_result and line_spacing_result):
            result = {
                "段落": paragraph.text,
                "字体": font_result,
                "字号": size_result,
                "对齐方式": align_result,
                "行间距": line_spacing_result
            }   

            return result
        
        return None
    
    def _check_paragraph_format_when_section_check(self, paragraph, rules):
        """检查单个段落的格式"""
        font_result = self._check_font(paragraph, rules.get("font"))
        size_result = self._check_size(paragraph, rules.get("size"))
        align_result = self._check_alignment(paragraph, rules.get("alignment"))
        line_spacing_result = self._check_line_spacing(paragraph, rules.get("line_spacing"))

        if not (font_result and size_result and align_result):
            result = {
                "段落": paragraph.text,
                "字体": font_result,
                "字号": size_result,
                "对齐方式": align_result,
                "行间距": line_spacing_result
            }   

            return result
        
        return {"检查结果": "格式匹配无误"}

    @staticmethod
    def _normalize_font_name(font_name: str | None) -> str:
        if not isinstance(font_name, str):
            return ""
        return re.sub(r"[\s_\-]", "", font_name).lower()

    def _build_font_alias_lookup(self) -> Dict[str, str]:
        lookup: Dict[str, str] = {}
        for canonical, aliases in self.font_alias_groups.items():
            canonical_norm = self._normalize_font_name(canonical)
            lookup[canonical_norm] = canonical_norm
            for alias in aliases:
                lookup[self._normalize_font_name(alias)] = canonical_norm
        return lookup

    def _is_font_equivalent(self, expected_font: str, actual_font: str) -> bool:
        expected_norm = self._normalize_font_name(expected_font)
        actual_norm = self._normalize_font_name(actual_font)
        if not expected_norm or not actual_norm:
            return False
        expected_canonical = self.font_alias_lookup.get(expected_norm, expected_norm)
        actual_canonical = self.font_alias_lookup.get(actual_norm, actual_norm)
        return expected_canonical == actual_canonical

    def _get_run_font_candidates(self, run) -> list[str]:
        fonts: list[str] = []

        # 常规路径
        run_font = run.font.name if getattr(run, "font", None) else None
        if run_font:
            fonts.append(str(run_font))

        # 底层 XML 路径（eastAsia/ascii/hAnsi/cs）
        try:
            rpr = run._element.rPr
            if rpr is not None:
                rfonts = rpr.rFonts
                if rfonts is not None:
                    for key in ("eastAsia", "ascii", "hAnsi", "cs"):
                        value = rfonts.get(qn(f"w:{key}"))
                        if value:
                            fonts.append(str(value))
        except Exception:
            pass

        return [f for f in fonts if f]

    def _get_paragraph_style_font_candidates(self, paragraph) -> list[str]:
        fonts: list[str] = []
        style = getattr(paragraph, "style", None)
        if style is None:
            return fonts

        # 段落样式字体
        try:
            style_font = getattr(style, "font", None)
            style_font_name = getattr(style_font, "name", None) if style_font else None
            if style_font_name:
                fonts.append(str(style_font_name))
        except Exception:
            pass

        # 底层 XML 字体
        try:
            style_element = getattr(style, "_element", None)
            rpr = getattr(style_element, "rPr", None) if style_element is not None else None
            rfonts = getattr(rpr, "rFonts", None) if rpr is not None else None
            if rfonts is not None:
                for key in ("eastAsia", "ascii", "hAnsi", "cs"):
                    value = rfonts.get(qn(f"w:{key}"))
                    if value:
                        fonts.append(str(value))
        except Exception:
            pass

        return [f for f in fonts if f]
        
    def _check_font(self, paragraph, expected_font):
        """检查字体"""
        if not expected_font:
            return True

        found_explicit_font = False
        for run in paragraph.runs:
            if run.text.strip() == "":
                continue

            candidate_fonts = self._get_run_font_candidates(run)
            if not candidate_fonts:
                continue
            found_explicit_font = True

            if not any(self._is_font_equivalent(expected_font, x) for x in candidate_fonts):
                return False

        if found_explicit_font:
            return True

        style_fonts = self._get_paragraph_style_font_candidates(paragraph)
        if style_fonts:
            return any(self._is_font_equivalent(expected_font, x) for x in style_fonts)

        return True
        
    def _check_size(self, paragraph, expected_size):
        """检查字号"""
        target_size = self._get_font_size_pt(expected_size)
        if target_size is None:
            return True

        for run in paragraph.runs:
            if run.text.strip() == "":
                continue

            actual_size = run.font.size.pt if run.font.size else None

            if actual_size is None:
                continue
            
            # Word 中字号常有小数偏差，允许 0.5pt 容差
            if abs(actual_size - target_size) > 0.5:
                return False
            
        return True
        
    def _check_alignment(self, paragraph, expected_alignment):
        """检查对齐方式"""
        actual_alignment = paragraph.alignment
        target_alignment = self.alignment_mapping[expected_alignment]

        # 段落未显式设置对齐时，尝试读取样式；仍为空则按“未指定”处理，不判错
        if actual_alignment is None:
            style = getattr(paragraph, "style", None)
            style_format = getattr(style, "paragraph_format", None) if style else None
            actual_alignment = getattr(style_format, "alignment", None) if style_format else None
            if actual_alignment is None:
                return True

        return actual_alignment == target_alignment
    
    def _check_line_spacing(self, paragraph, expected_line_spacing):
        """检查行间距"""
        actual_line_spacing_rule = paragraph.paragraph_format.line_spacing_rule if paragraph.paragraph_format.line_spacing_rule else None
        actual_line_spacing = paragraph.paragraph_format.line_spacing if paragraph.paragraph_format.line_spacing else None

        if actual_line_spacing_rule is None or actual_line_spacing is None:
            return True

        target_line_spacing_rule = self.line_spacing_rule_mapping[expected_line_spacing]
        target_line_spacing = self.line_spacing_mapping[expected_line_spacing]

        if expected_line_spacing == "0.5倍行距":
            return actual_line_spacing_rule == target_line_spacing_rule and abs(actual_line_spacing - target_line_spacing) <= 0.1
        elif expected_line_spacing == "单倍行距":
            return (actual_line_spacing_rule == target_line_spacing_rule and abs(actual_line_spacing - target_line_spacing) <= 0.1) or actual_line_spacing_rule == WD_LINE_SPACING.SINGLE
        elif expected_line_spacing == "1.5倍行距":
            return (actual_line_spacing_rule == target_line_spacing_rule and abs(actual_line_spacing - target_line_spacing) <= 0.1) or actual_line_spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE
        elif expected_line_spacing == "2倍行距":
            return (actual_line_spacing_rule == target_line_spacing_rule and abs(actual_line_spacing - target_line_spacing) <= 0.1) or actual_line_spacing_rule == WD_LINE_SPACING.DOUBLE
        elif expected_line_spacing == "3倍行距":    
            return actual_line_spacing_rule == target_line_spacing_rule and abs(actual_line_spacing - target_line_spacing) <= 0.1
        elif "固定值" in expected_line_spacing: 
            return actual_line_spacing_rule == target_line_spacing_rule and abs(actual_line_spacing.pt - target_line_spacing) <= 0.5

        
