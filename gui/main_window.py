import json
from typing import Dict
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QLabel, QScrollArea,
                             QGroupBox, QFormLayout, QTabWidget, QComboBox,
                             QGroupBox, QFormLayout, QTabWidget, QComboBox,
                             QMessageBox)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
import os
from model.document_parser import DocumentParser
from model.format_checker import FormatChecker
from model.format_modifier import FormatModifier
from model import create_default_hybrid_processor

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # --- 定义更丰富的选项 --- 
        self.font_options = ["华文中宋", "宋体", "黑体", "楷体", "仿宋", "微软雅黑", "等线", "Times New Roman", "Arial"]
        
        self.size_options = [
            "六号 (7.5pt)", "小五 (9pt)", "五号 (10.5pt)", "小四 (12pt)", "四号 (14pt)", 
            "小三 (15pt)", "三号 (16pt)", "小二 (18pt)", "二号 (22pt)", "小一 (24pt)", 
            "一号 (26pt)", "小初 (36pt)", "初号 (42pt)",
            "8pt", "9pt", "10pt", "11pt", "12pt", "14pt", "16pt", "18pt", "20pt", 
            "22pt", "24pt", "26pt", "28pt", "36pt", "48pt", "72pt"
        ]
        self.line_spacing_options = ["0.5倍行距","单倍行距", "1.5倍行距", "2倍行距","3倍行距", "固定值12pt", "固定值15pt", "固定值20pt"]

        self.alignment_options = ["左对齐", "居中", "右对齐", "两端对齐", "分散对齐"]

        self.table_style_options = [
            "无样式", "表格网格", "浅色底纹", "浅色列表", "浅色网格", 
            "中度底纹 - 强调文字颜色 1", "中度列表 - 强调文字颜色 2", 
            "中度网格 - 强调文字颜色 3", "深色列表"
        ]

        self.setWindowTitle("Word文档格式检查与修改软件")
        self.setMinimumSize(800, 600)
        
        # 启用拖拽
        self.setAcceptDrops(True)
        
        # 初始化文档解析器和格式检查器
        self.parser = DocumentParser()
        self.checker = FormatChecker()
        self.modifier = FormatModifier()
        self.hybrid_processor = create_default_hybrid_processor()
        
        # 创建主窗口部件
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        
        # 创建主布局
        self.main_layout = QVBoxLayout(self.main_widget)
        
        # 创建文件操作区
        self.create_file_operation_area()
        
        # 创建格式设置区
        self.create_format_settings_area()
        
        # 创建操作按钮区
        self.create_operation_buttons()
        
        # 创建检查结果区
        self.create_result_area()
        
    def create_file_operation_area(self):
        """创建文件操作区域"""
        file_group = QGroupBox("文件操作")
        file_layout = QVBoxLayout()  # 改为垂直布局以添加提示文本
        
        # 添加拖拽提示
        drag_hint = QLabel("将Word文档拖拽到此处，或点击下方按钮选择")
        drag_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        file_layout.addWidget(drag_hint)
        
        # 文件操作按钮区域
        button_area = QHBoxLayout()
        self.file_path_label = QLabel("未选择文件")
        self.select_file_btn = QPushButton("选择文件")
        self.select_file_btn.clicked.connect(self.select_file)
        
        button_area.addWidget(self.file_path_label)
        button_area.addWidget(self.select_file_btn)
        
        file_layout.addLayout(button_area)
        file_group.setLayout(file_layout)
        
        # 设置文件操作区接受拖拽
        file_group.setAcceptDrops(True)
        
        self.main_layout.addWidget(file_group)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        """处理拖拽进入事件"""
        if event.mimeData().hasUrls():
            # 获取拖拽的文件路径
            file_path = event.mimeData().urls()[0].toLocalFile()
            # 检查是否为docx文件
            if file_path.lower().endswith('.docx'):
                event.acceptProposedAction()
                
    def dropEvent(self, event: QDropEvent):
        """处理拖拽释放事件"""
        file_path = event.mimeData().urls()[0].toLocalFile()
        if file_path.lower().endswith('.docx'):
            self.file_path_label.setText(file_path)
            event.acceptProposedAction()
        else:
            # 如果不是docx文件，显示错误信息
            self.file_path_label.setText("请选择Word文档(.docx)文件")
            
    def create_format_settings_area(self):
        """创建格式设置区域"""
        settings_group = QGroupBox("格式设置")
        settings_layout = QVBoxLayout()
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        self.format_combos = {} # 用于存储所有的下拉菜单

        self.create_cover_tab() # 创建封面选项卡

        self.create_statement_tab() # 创建原创性声明选项卡

        self.create_abstract_keyword_tab() # 创建摘要和关键词选项卡

        self.create_main_text_tab() # 创建正文选项卡
        
        self.create_headings_tab() # 创建标题选项卡

        self.create_figures_or_tables_title_tab() # 创建图|表题选页卡

        self.create_references_tab() # 创建参考文献选项卡

        self.create_acknowledgments_tab() # 创建致谢选项卡
        
        # 设置默认值
        self.set_default_format_values()
        
        settings_layout.addWidget(self.tab_widget)
        settings_group.setLayout(settings_layout)
        self.main_layout.addWidget(settings_group)

    def create_combo_box(self, section, layout, label, key):
        self.format_combos[section][key] = QComboBox()
        self.format_combos[section][key].addItems(self.font_options if 'font' in key else self.size_options if 'size' in key else self.alignment_options if 'align' in key else self.line_spacing_options)
        layout.addRow(label, self.format_combos[section][key])

    def create_sub_tab(self, section, prefix):
        sub_tab = QWidget()
        sub_tab_layout = QFormLayout(sub_tab)

        self.create_combo_box(section, sub_tab_layout, "字体", f"{prefix}_font")
        self.create_combo_box(section, sub_tab_layout, "字号", f"{prefix}_size")
        self.create_combo_box(section, sub_tab_layout, "对齐方式", f"{prefix}_align")
        self.create_combo_box(section, sub_tab_layout, "行间距", f"{prefix}_line_spacing")

        return sub_tab
        
    def create_cover_tab(self, section = "封面"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)

        # 创建二级选项卡
        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        # 学校名称设置
        school_tab = self.create_sub_tab(section, "school")
        sub_tab_widget.addTab(school_tab, "学校")
                
        # 论文题目设置
        title_tab = self.create_sub_tab(section, "title")
        sub_tab_widget.addTab(title_tab, "论文题目")
                
        # 个人信息设置
        info_tab = self.create_sub_tab(section, "personal_information")
        sub_tab_widget.addTab(info_tab, "个人信息")
                
        tab_layout.addWidget(sub_tab_widget)

        self.tab_widget.addTab(tab, section)

    def create_statement_tab(self, section = "声明"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)

        # 创建二级选项卡
        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        # 声明标题设置
        title_tab = self.create_sub_tab(section, "title")
        sub_tab_widget.addTab(title_tab, "声明标题")
                
        # 声明内容设置
        content_tab = self.create_sub_tab(section, "content")
        sub_tab_widget.addTab(content_tab, "声明内容")
                
        tab_layout.addWidget(sub_tab_widget)

        self.tab_widget.addTab(tab, section)

    def create_abstract_keyword_tab(self, section = "摘要与关键词"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)

        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        # 中文摘要设置
        chinese_abstract_tab = QWidget()
        chinese_abstract_layout = QFormLayout(chinese_abstract_tab)

        # 摘要标题设置
        chinese_title_tab = self.create_sub_tab(section, "chinese_title")

        # 摘要内容设置
        chinese_content_tab = self.create_sub_tab(section, "chinese_content")

        # 关键词标题设置
        chinese_keyword_title_tab = self.create_sub_tab(section, "chinese_keyword_title")
        
        # 关键词内容设置
        chinese_keyword_tab = self.create_sub_tab(section, "chinese_keyword")
        

        # 为中文摘要添加三级 tab
        chinese_sub_tab_widget = QTabWidget()

        chinese_sub_tab_widget.addTab(chinese_title_tab, "摘要标题")
        chinese_sub_tab_widget.addTab(chinese_content_tab, "摘要内容")
        chinese_sub_tab_widget.addTab(chinese_keyword_title_tab, "关键词标题")
        chinese_sub_tab_widget.addTab(chinese_keyword_tab, "关键词内容")

        chinese_abstract_layout.addWidget(chinese_sub_tab_widget)
        
        sub_tab_widget.addTab(chinese_abstract_tab, "中文摘要")
        
        # 英文摘要设置
        english_abstract_tab = QWidget()
        english_abstract_layout = QFormLayout(english_abstract_tab)

        # 摘要标题设置
        english_title_tab = self.create_sub_tab(section, "english_title")
        
        # 摘要内容设置
        english_content_tab = self.create_sub_tab(section, "english_content")

        # 关键词标题设置
        english_keyword_title_tab = self.create_sub_tab(section, "english_keyword_title")

        # 关键词内容设置
        english_keyword_tab = self.create_sub_tab(section, "english_keyword")

        # 为英文摘要添加三级 tab
        english_sub_tab_widget = QTabWidget()
        english_sub_tab_widget.addTab(english_title_tab, "标题")
        english_sub_tab_widget.addTab(english_content_tab, "内容")
        english_sub_tab_widget.addTab(english_keyword_title_tab, "关键词标题")
        english_sub_tab_widget.addTab(english_keyword_tab, "关键词内容")

        english_abstract_layout.addWidget(english_sub_tab_widget)
        
        sub_tab_widget.addTab(english_abstract_tab, "英文摘要")
        
        tab_layout.addWidget(sub_tab_widget)

        self.tab_widget.addTab(tab, section)

    def create_main_text_tab(self, section = "正文"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)
        
        # 字体设置
        self.format_combos[section]["font"] = QComboBox()
        self.format_combos[section]["font"].addItems(self.font_options)
        tab_layout.addRow("字体", self.format_combos[section]["font"])
        
        # 字号设置
        self.format_combos[section]["size"] = QComboBox()
        self.format_combos[section]["size"].addItems(self.size_options)
        tab_layout.addRow("字号", self.format_combos[section]["size"])
        
        # 对齐方式
        self.format_combos[section]["align"] = QComboBox()
        self.format_combos[section]["align"].addItems(self.alignment_options)
        tab_layout.addRow("对齐方式", self.format_combos[section]["align"])

        # 行间距设置
        self.format_combos[section]["line_spacing"] = QComboBox()
        self.format_combos[section]["line_spacing"].addItems(self.line_spacing_options)
        tab_layout.addRow("行间距", self.format_combos[section]["line_spacing"])

        self.tab_widget.addTab(tab, section)

    def create_headings_tab(self, section = "标题"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)

        # 创建二级选项卡
        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        chapter_tab = self.create_sub_tab(section, "chapter")
        
        sub_tab_widget.addTab(chapter_tab, f"章节标题")
        
        # 为各级标题添加设置
        for level in range(1, 4):
            level_tab = self.create_sub_tab(section, f"level{level}")
            
            sub_tab_widget.addTab(level_tab, f"{level}级标题")
        
        tab_layout.addWidget(sub_tab_widget)

        self.tab_widget.addTab(tab, section)

    def create_figures_or_tables_title_tab(self, section = "图|表题"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)
        
        # 字体设置
        self.format_combos[section]["font"] = QComboBox()
        self.format_combos[section]["font"].addItems(self.font_options)
        tab_layout.addRow("字体", self.format_combos[section]["font"])
        
        # 字号设置
        self.format_combos[section]["size"] = QComboBox()
        self.format_combos[section]["size"].addItems(self.size_options)
        tab_layout.addRow("字号", self.format_combos[section]["size"])
        
        # 对齐方式
        self.format_combos[section]["align"] = QComboBox()
        self.format_combos[section]["align"].addItems(self.alignment_options)
        tab_layout.addRow("对齐方式", self.format_combos[section]["align"])

        # 行间距设置
        self.format_combos[section]["line_spacing"] = QComboBox()
        self.format_combos[section]["line_spacing"].addItems(self.line_spacing_options)
        tab_layout.addRow("行间距", self.format_combos[section]["line_spacing"])

        self.tab_widget.addTab(tab, section)

    def create_references_tab(self, section = "参考文献"):
        """创建参考文献格式设置选项卡"""
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)

        # 标题字体
        title_font_combo = QComboBox()
        title_font_combo.addItems(self.font_options)
        self.format_combos[section]["title_font"] = title_font_combo
        tab_layout.addRow("标题字体", title_font_combo)

        # 标题字号
        title_size_combo = QComboBox()
        title_size_combo.addItems(self.size_options)
        self.format_combos[section]["title_size"] = title_size_combo
        tab_layout.addRow("标题字号", title_size_combo)

        # 标题对齐方式
        title_align_combo = QComboBox()
        title_align_combo.addItems(self.alignment_options)
        self.format_combos[section]["title_align"] = title_align_combo
        tab_layout.addRow("标题对齐方式", title_align_combo)

        # 标题行间距
        title_line_spacing_combo = QComboBox()
        title_line_spacing_combo.addItems(self.line_spacing_options)
        self.format_combos[section]["title_line_spacing"] = title_line_spacing_combo
        tab_layout.addRow("标题行间距", title_line_spacing_combo)

        # 内容字体
        content_font_combo = QComboBox()
        content_font_combo.addItems(self.font_options)
        self.format_combos[section]["content_font"] = content_font_combo
        tab_layout.addRow("内容字体", content_font_combo)

        # 内容字号
        content_size_combo = QComboBox()
        content_size_combo.addItems(self.size_options)
        self.format_combos[section]["content_size"] = content_size_combo
        tab_layout.addRow("内容字号", content_size_combo)

        # 内容对齐方式
        content_align_combo = QComboBox()
        content_align_combo.addItems(self.alignment_options)
        self.format_combos[section]["content_align"] = content_align_combo
        tab_layout.addRow("内容对齐方式", content_align_combo)

        # 内容行间距
        content_line_spacing_combo = QComboBox()
        content_line_spacing_combo.addItems(self.line_spacing_options)
        self.format_combos[section]["content_line_spacing"] = content_line_spacing_combo
        tab_layout.addRow("内容行间距", content_line_spacing_combo)

        self.tab_widget.addTab(tab, section)

    def create_acknowledgments_tab(self, section = "致谢"):
        """创建致谢格式设置选项卡"""
        self.format_combos[section] = {}
        tab = QWidget()
        tab_layout = QFormLayout(tab)

        # 标题字体
        title_font_combo = QComboBox()
        title_font_combo.addItems(self.font_options)
        self.format_combos[section]["title_font"] = title_font_combo
        tab_layout.addRow("标题字体", title_font_combo)

        # 标题字号
        title_size_combo = QComboBox()
        title_size_combo.addItems(self.size_options)
        self.format_combos[section]["title_size"] = title_size_combo
        tab_layout.addRow("标题字号", title_size_combo)

        # 标题对齐方式
        title_align_combo = QComboBox()
        title_align_combo.addItems(self.alignment_options)
        self.format_combos[section]["title_align"] = title_align_combo
        tab_layout.addRow("标题对齐方式", title_align_combo)

        # 标题行间距
        title_line_spacing_combo = QComboBox()
        title_line_spacing_combo.addItems(self.line_spacing_options)
        self.format_combos[section]["title_line_spacing"] = title_line_spacing_combo
        tab_layout.addRow("标题行间距", title_line_spacing_combo)

        # 内容字体
        content_font_combo = QComboBox()
        content_font_combo.addItems(self.font_options)
        self.format_combos[section]["content_font"] = content_font_combo
        tab_layout.addRow("内容字体", content_font_combo)

        # 内容字号
        content_size_combo = QComboBox()
        content_size_combo.addItems(self.size_options)
        self.format_combos[section]["content_size"] = content_size_combo
        tab_layout.addRow("内容字号", content_size_combo)

        # 内容对齐方式
        content_align_combo = QComboBox()
        content_align_combo.addItems(self.alignment_options)
        self.format_combos[section]["content_align"] = content_align_combo
        tab_layout.addRow("内容对齐方式", content_align_combo)

        # 内容行间距
        content_line_spacing_combo = QComboBox()
        content_line_spacing_combo.addItems(self.line_spacing_options)
        self.format_combos[section]["content_line_spacing"] = content_line_spacing_combo
        tab_layout.addRow("内容行间距", content_line_spacing_combo)

        self.tab_widget.addTab(tab, section) 

    def set_default_format_value(self, section, key, font, size, align, line_spacing):
        """设置默认的格式值"""
        self.format_combos[section][f"{key}_font"].setCurrentText(font)
        self.format_combos[section][f"{key}_size"].setCurrentText(size)
        self.format_combos[section][f"{key}_align"].setCurrentText(align)
        self.format_combos[section][f"{key}_line_spacing"].setCurrentText(line_spacing)

    def set_default_format_values(self):
        """设置默认的格式值"""
        # 封面
        if "封面" in self.format_combos:
            # 学校名称默认值
            self.set_default_format_value("封面", "school", "华文中宋", "一号 (26pt)", "居中", "3倍行距")
            
            # 论文题目默认值
            self.set_default_format_value("封面", "title", "黑体", "二号 (22pt)", "居中", "固定值20pt")
            
            # 个人信息默认值
            self.set_default_format_value("封面", "personal_information", "华文中宋", "三号 (16pt)", "两端对齐", "固定值20pt")
            
        # 原创性声明
        if "声明" in self.format_combos:
            # 声明标题默认值
            self.set_default_format_value("声明", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")

            # 声明内容默认值
            self.set_default_format_value("声明", "content", "宋体", "小四 (12pt)", "两端对齐", "固定值20pt")
            
        # 摘要和关键词
        if "摘要与关键词" in self.format_combos:
            # 中文摘要默认值
            self.set_default_format_value("摘要与关键词", "chinese_title", "黑体", "小二 (18pt)", "居中", "固定值20pt")

            self.set_default_format_value("摘要与关键词", "chinese_content", "宋体", "小四 (12pt)", "两端对齐", "固定值20pt")

            self.set_default_format_value("摘要与关键词", "chinese_keyword_title", "黑体", "四号 (14pt)", "两端对齐", "固定值20pt")

            self.set_default_format_value("摘要与关键词", "chinese_keyword", "宋体", "小四 (12pt)", "两端对齐", "固定值20pt")
            
            # 英文摘要默认值
            self.set_default_format_value("摘要与关键词", "english_title", "Times New Roman", "小二 (18pt)", "居中", "固定值20pt")

            self.set_default_format_value("摘要与关键词", "english_content", "Times New Roman", "小四 (12pt)", "两端对齐", "固定值20pt")

            self.set_default_format_value("摘要与关键词", "english_keyword_title", "Times New Roman", "四号 (14pt)", "两端对齐", "固定值20pt")

            self.set_default_format_value("摘要与关键词", "english_keyword", "Times New Roman", "小四 (12pt)", "两端对齐", "固定值20pt")
            
        # 正文
        if "正文" in self.format_combos:
            self.format_combos["正文"]["font"].setCurrentText("宋体")
            self.format_combos["正文"]["size"].setCurrentText("小四 (12pt)")
            self.format_combos["正文"]["align"].setCurrentText("两端对齐")
            self.format_combos["正文"]["line_spacing"].setCurrentText("固定值20pt")
            
        # 标题
        if "标题" in self.format_combos:
            # 章节标题
            self.set_default_format_value("标题", "chapter", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            
            # 一级标题
            self.set_default_format_value("标题", "level1", "黑体", "三号 (16pt)", "左对齐", "固定值20pt")
            
            # 二级标题
            self.set_default_format_value("标题", "level2", "黑体", "四号 (14pt)", "左对齐", "固定值20pt")

            # 三级标题
            self.set_default_format_value("标题", "level3", "黑体", "小四 (12pt)", "左对齐", "固定值20pt")

        # 图|表题
        if "图|表题" in self.format_combos:
            self.format_combos["图|表题"]["font"].setCurrentText("黑体")
            self.format_combos["图|表题"]["size"].setCurrentText("小四 (12pt)")
            self.format_combos["图|表题"]["align"].setCurrentText("居中")
            self.format_combos["图|表题"]["line_spacing"].setCurrentText("固定值20pt")
            
        # 参考文献
        if "参考文献" in self.format_combos:
            # 参考文献标题  
            self.set_default_format_value("参考文献", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            
            # 参考文献内容
            self.set_default_format_value("参考文献", "content", "宋体", "小四 (12pt)", "左对齐", "固定值20pt")
            
        # 致谢
        if "致谢" in self.format_combos:
            # 致谢标题
            self.set_default_format_value("致谢", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            
            # 致谢内容
            self.set_default_format_value("致谢", "content", "宋体", "小四 (12pt)", "两端对齐", "固定值20pt")
        
    def create_operation_buttons(self):
        """创建操作按钮区域"""
        button_layout = QHBoxLayout()
        
        self.check_btn = QPushButton("检查格式")
        self.check_btn.clicked.connect(self.check_format)

        self.order_check_btn = QPushButton("检查排版顺序")
        self.order_check_btn.clicked.connect(self.check_order)

        self.hybrid_check_btn = QPushButton("混合检查")
        self.hybrid_check_btn.clicked.connect(self.hybrid_check)
        
        self.modify_btn = QPushButton("确认修改")
        self.modify_btn.clicked.connect(self.modify_format)
        
        button_layout.addWidget(self.check_btn)
        button_layout.addWidget(self.order_check_btn)
        button_layout.addWidget(self.hybrid_check_btn)
        button_layout.addWidget(self.modify_btn)
        
        self.main_layout.addLayout(button_layout)
        
    def create_result_area(self):
        """创建检查结果区域"""
        result_group = QGroupBox("检查结果（结果仅供参考）")
        result_layout = QVBoxLayout()
        
        self.result_scroll = QScrollArea()
        self.result_widget = QWidget()
        self.result_layout = QVBoxLayout(self.result_widget)
        
        self.result_scroll.setWidget(self.result_widget)
        self.result_scroll.setWidgetResizable(True)
        
        result_layout.addWidget(self.result_scroll)
        result_group.setLayout(result_layout)
        
        self.main_layout.addWidget(result_group)
        
    def select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx)"
        )
        if file_path:
            self.file_path_label.setText(file_path)
            
    def get_format_settings(self):
        """获取用户设置的格式"""
        formats = {}
        
        for section, combos in self.format_combos.items():
            section_formats = {}
            for key, combo in combos.items():
                section_formats[key] = combo.currentText()
            formats[section] = section_formats
            
        return formats
            
    def check_format(self):
        """检查格式"""
        if not self.file_path_label.text() or self.file_path_label.text() == "未选择文件":
            QMessageBox.warning(
                self,
                "提示",
                "未选择文件，请选择Word文档",
                QMessageBox.StandardButton.Ok
            )
            return
            
        # 获取用户设置的格式
        user_formats = self.get_format_settings()
        
        # 更新检查器中的格式设置
        self.checker.update_formats(user_formats)
            
        # 解析文档
        self.doc_section = self.parser.parse_document(self.file_path_label.text())
        
        # 检查格式
        results = self.checker.check_format(self.doc_section)
        
        # 显示结果
        self.display_results(results)

    def check_order(self):
        """检查排版顺序"""
        if not self.file_path_label.text() or self.file_path_label.text() == "未选择文件":
            QMessageBox.warning(
                self,
                "提示",
                "未选择文件，请选择Word文档",
                QMessageBox.StandardButton.Ok
            )
            return
            
        # 解析文档
        self.doc_section = self.parser.parse_document(self.file_path_label.text())
        
        # 检查排版顺序
        self.parser.check_order()

    def hybrid_check(self):
        """执行混合检查（docx + pdf + ocr占位）"""
        if not self.file_path_label.text() or self.file_path_label.text() == "未选择文件":
            QMessageBox.warning(
                self,
                "提示",
                "未选择文件，请选择Word文档",
                QMessageBox.StandardButton.Ok
            )
            return

        try:
            user_formats = self.get_format_settings()
            result = self.hybrid_processor.process(
                self.file_path_label.text(),
                user_formats=user_formats
            )
            self.display_hybrid_results(result)
        except Exception as e:
            QMessageBox.critical(
                self,
                "错误",
                f"混合检查失败：\n{str(e)}",
                QMessageBox.StandardButton.Ok
            )
        
    def modify_format(self):
        """修改格式"""
        if not self.file_path_label.text() or self.file_path_label.text() == "未选择文件":
            QMessageBox.warning(
                self,
                "提示",
                "未选择文件，请选择Word文档",
                QMessageBox.StandardButton.Ok
            )
            return
            
        try:
            # 获取用户设置的格式
            user_formats = self.get_format_settings()
            
            # 更新修改器中的格式设置
            self.modifier.update_formats(user_formats)

            # 获取文档地址
            self.modifier.doc = self.parser.doc

            # 获取解析好的section
            self.modifier.sections = self.parser.sections
                
            # 修改格式
            new_file_path = self.modifier.modify_format(self.file_path_label.text())
            
            # 获取原始文件名和目录
            original_dir = os.path.dirname(self.file_path_label.text())
            new_file_name = os.path.basename(new_file_path)
            
            # 显示成功消息
            QMessageBox.information(
                self,
                "修改完成",
                f"文档格式已修改完成！\n\n"
                f"修改后的文件已保存为：\n{new_file_name}\n\n"
                f"保存位置：\n{original_dir}",
                QMessageBox.StandardButton.Ok
            )
            
            # 清空之前的检查结果
            while self.result_layout.count():
                item = self.result_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
                    
            # 添加成功提示到结果区
            success_label = QLabel(f"✓ 格式修改成功")
            success_label.setStyleSheet("color: green; font-weight: bold; font-size: 14px;")
            success_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.result_layout.addWidget(success_label)
            
            file_info_label = QLabel(
                f"修改后的文件：{new_file_name}\n"
                f"保存位置：{original_dir}"
            )
            file_info_label.setWordWrap(True)
            file_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.result_layout.addWidget(file_info_label)
            
        except Exception as e:
            # 显示错误消息
            QMessageBox.critical(
                self,
                "错误",
                f"修改文档格式时发生错误：\n{str(e)}",
                QMessageBox.StandardButton.Ok
            )
        
    def display_results(self, results):
        """显示检查结果"""
        # 清空现有结果
        while self.result_layout.count():
            item = self.result_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
                
        # 添加新结果
        for section, result in results.items():
            section_widget = QGroupBox(self.get_section_name(section))
            section_layout = QFormLayout()
            result_str = ""
            if isinstance(result, Dict):
                result_str = '\n'.join([f"{key}: {value}" for key, value in result.items()])
            else:
                for result in result:
                    line = ", ".join(f"'{key}': '{value}'" for key, value in result.items())
                    result_str += "{" + line + "}" + "\n"

            content_label = QLabel(result_str)

            section_layout.addRow(content_label)
            
            section_widget.setLayout(section_layout)
            self.result_layout.addWidget(section_widget)

    def display_hybrid_results(self, result):
        """显示混合检查结果"""
        while self.result_layout.count():
            item = self.result_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        summary = result.get("summary", {})
        engine_counts = result.get("engine_counts", {})
        context_status = result.get("context_status", {})
        issues = result.get("issues", [])

        overview_widget = QGroupBox("检查概览")
        overview_layout = QFormLayout()
        overview_text = (
            f"总问题数：{summary.get('total', 0)}\n"
            f"错误：{summary.get('errors', 0)}，警告：{summary.get('warnings', 0)}，提示：{summary.get('infos', 0)}\n"
            f"各引擎问题数：{engine_counts}\n"
            f"上下文状态：{context_status}"
        )
        overview_label = QLabel(overview_text)
        overview_label.setWordWrap(True)
        overview_layout.addRow(overview_label)
        overview_widget.setLayout(overview_layout)
        self.result_layout.addWidget(overview_widget)

        if not issues:
            no_issue_label = QLabel("✓ 未发现格式问题")
            no_issue_label.setStyleSheet("color: green; font-weight: bold; font-size: 14px;")
            no_issue_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.result_layout.addWidget(no_issue_label)
            return

        grouped = {}
        for issue in issues:
            source = issue.get("source", "unknown")
            grouped.setdefault(source, []).append(issue)

        for source, source_issues in grouped.items():
            source_widget = QGroupBox(f"{source} 检查结果")
            source_layout = QFormLayout()
            lines = []
            for issue in source_issues:
                lines.append(
                    f"[{issue.get('severity', 'info')}] {issue.get('title', '')}: "
                    f"{issue.get('message', '')}"
                )
            content_label = QLabel("\n".join(lines))
            content_label.setWordWrap(True)
            source_layout.addRow(content_label)
            source_widget.setLayout(source_layout)
            self.result_layout.addWidget(source_widget)
            
    def get_section_name(self, section_key):
        """根据section键获取中文名称"""
        section_names = {
            "cover": "封面",
            "statement": "声明",
            "abstract_keyword": "摘要与关键词",
            "catalogue": "目录",
            "main_text": "正文",
            "headings": "标题",
            "figures_or_tables_title": "图|表题",
            "figures": "图",
            "tables": "表",
            "references": "参考文献",
            "references_check": "文献引用格式",
            "acknowledgments": "致谢"
        }
        return section_names.get(section_key, section_key)
