import json
import re
from typing import Dict
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QLabel, QScrollArea,
                             QGroupBox, QFormLayout, QTabWidget, QComboBox,
                             QGroupBox, QFormLayout, QTabWidget, QComboBox,
                             QSpinBox, QSizePolicy, QListWidget, QListWidgetItem,
                             QMessageBox)
from PyQt6.QtCore import Qt, QMimeData
from PyQt6.QtGui import (
    QDragEnterEvent,
    QDropEvent,
    QPixmap,
    QImage,
    QPainter,
    QColor,
    QPen,
    QFont,
)
import os
from model.document_parser import DocumentParser
from model.format_checker import FormatChecker
from model.format_modifier import FormatModifier
from model import create_default_hybrid_processor
from model.pdf_engine import DocxToPdfConverter
from model.pdf_engine.extractor import PdfExtractor
from model.pdf_engine.page_roles import (
    PAGE_ROLE_BACKMATTER,
    PAGE_ROLE_CATALOGUE,
    PAGE_ROLE_CN_ABSTRACT,
    PAGE_ROLE_EN_ABSTRACT,
    PAGE_ROLE_MAIN,
    PAGE_ROLE_OTHER,
    build_page_roles,
    is_catalogue_page,
)

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

        self.setWindowTitle("Word文档格式检查系统")
        self.setMinimumSize(800, 600)
        
        # 启用拖拽
        self.setAcceptDrops(True)
        
        # 初始化文档解析器和格式检查器
        self.parser = DocumentParser()
        self.checker = FormatChecker()
        self.modifier = FormatModifier()
        self.hybrid_processor = create_default_hybrid_processor()
        self.docx_to_pdf_converter = DocxToPdfConverter()
        self.pdf_extractor = PdfExtractor()
        self.generated_pdf_path = None
        self.visual_issue_state = None
        self.issue_widget_map = {}
        
        # 创建主窗口部件
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        
        # 创建主布局
        self.main_layout = QHBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        self.main_layout.setSpacing(10)

        self.left_panel_widget = QWidget()
        self.left_panel_layout = QVBoxLayout(self.left_panel_widget)
        self.left_panel_layout.setContentsMargins(0, 0, 0, 0)
        self.left_panel_layout.setSpacing(8)

        self.right_panel_widget = QWidget()
        self.right_panel_layout = QVBoxLayout(self.right_panel_widget)
        self.right_panel_layout.setContentsMargins(0, 0, 0, 0)
        self.right_panel_layout.setSpacing(8)

        self.main_layout.addWidget(self.left_panel_widget, 5)
        self.main_layout.addWidget(self.right_panel_widget, 6)
        
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
        file_layout = QVBoxLayout()

        self.document_check_btn = QPushButton("文档检查")
        self.document_check_btn.clicked.connect(self.hybrid_check)
        file_layout.addWidget(self.document_check_btn)
        
        # 文件操作按钮区域
        button_area = QHBoxLayout()
        self.file_path_label = QLabel("未选择文件")
        self.select_file_btn = QPushButton("选择文件")
        self.select_file_btn.clicked.connect(self.select_file)
        
        button_area.addWidget(self.file_path_label)
        button_area.addWidget(self.select_file_btn)
        
        file_layout.addLayout(button_area)

        pdf_area = QHBoxLayout()
        self.pdf_path_label = QLabel("未选择PDF（将自动尝试同名PDF）")
        self.select_pdf_btn = QPushButton("选择PDF")
        self.select_pdf_btn.clicked.connect(self.select_pdf)
        pdf_area.addWidget(self.pdf_path_label)
        pdf_area.addWidget(self.select_pdf_btn)
        file_layout.addLayout(pdf_area)
        file_group.setLayout(file_layout)
        
        # 设置文件操作区接受拖拽
        file_group.setAcceptDrops(True)
        
        self.right_panel_layout.addWidget(file_group)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        """处理拖拽进入事件"""
        if event.mimeData().hasUrls():
            # 获取拖拽的文件路径
            file_path = event.mimeData().urls()[0].toLocalFile()
            # 检查是否为docx文件
            if file_path.lower().endswith('.docx'):
                event.acceptProposedAction()
                
    @staticmethod
    def normalize_display_path(path: str) -> str:
        if not path:
            return path
        return os.path.normpath(path)

    def dropEvent(self, event: QDropEvent):
        """处理拖拽释放事件"""
        file_path = event.mimeData().urls()[0].toLocalFile()
        if file_path.lower().endswith('.docx'):
            self.file_path_label.setText(self.normalize_display_path(file_path))
            self.sync_pdf_path(file_path)
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

        self.create_catalogue_tab() # 创建目录选项卡

        self.create_main_text_tab() # 创建正文选项卡
        
        self.create_headings_tab() # 创建标题选项卡

        self.create_figures_or_tables_title_tab() # 创建图|表题选页卡

        self.create_references_tab() # 创建参考文献选项卡

        self.create_acknowledgments_tab() # 创建致谢选项卡
        
        # 设置默认值
        self.set_default_format_values()
        self.apply_post_defaults()
        
        settings_layout.addWidget(self.tab_widget)
        settings_group.setLayout(settings_layout)
        settings_group.setMaximumHeight(280)
        self.right_panel_layout.addWidget(settings_group)

    def create_combo_box(self, section, layout, label, key):
        self.format_combos[section][key] = QComboBox()
        self.format_combos[section][key].addItems(self.font_options if 'font' in key else self.size_options if 'size' in key else self.alignment_options if 'align' in key else self.line_spacing_options)
        layout.addRow(label, self.format_combos[section][key])

    def apply_post_defaults(self):
        """补充设置需要明确覆盖的默认值。"""
        if "参考文献" in self.format_combos:
            self.set_default_format_value("参考文献", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            self.set_default_format_value("参考文献", "content", "宋体", "五号 (10.5pt)", "左对齐", "固定值20pt")

        if "致谢" in self.format_combos:
            self.set_default_format_value("致谢", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            self.set_default_format_value("致谢", "content", "宋体", "小四 (12pt)", "两端对齐", "固定值20pt")

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
        english_sub_tab_widget.addTab(english_title_tab, "摘要标题")
        english_sub_tab_widget.addTab(english_content_tab, "摘要内容")
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

    def create_catalogue_tab(self, section = "目录"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)
        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        sub_tab_widget.addTab(self.create_sub_tab(section, "title"), "目录标题")
        sub_tab_widget.addTab(self.create_sub_tab(section, "content"), "目录内容")

        tab_layout.addWidget(sub_tab_widget)
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

    def create_references_tab(self, section = "参考文献"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)
        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        sub_tab_widget.addTab(self.create_sub_tab(section, "title"), "参考文献标题")
        sub_tab_widget.addTab(self.create_sub_tab(section, "content"), "参考文献内容")

        tab_layout.addWidget(sub_tab_widget)
        self.tab_widget.addTab(tab, section)

    def create_acknowledgments_tab(self, section = "致谢"):
        self.format_combos[section] = {}

        tab = QWidget()
        tab_layout = QFormLayout(tab)
        sub_tab_widget = QTabWidget()
        sub_tab_widget.setTabPosition(QTabWidget.TabPosition.North)

        sub_tab_widget.addTab(self.create_sub_tab(section, "title"), "致谢标题")
        sub_tab_widget.addTab(self.create_sub_tab(section, "content"), "致谢内容")

        tab_layout.addWidget(sub_tab_widget)
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

        # 目录
        if "目录" in self.format_combos:
            self.set_default_format_value("目录", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            self.set_default_format_value("目录", "content", "宋体", "小四 (12pt)", "左对齐", "固定值20pt")
            
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
            self.format_combos["图|表题"]["font"].setCurrentText("宋体")
            self.format_combos["图|表题"]["size"].setCurrentText("小四 (12pt)")
            self.format_combos["图|表题"]["align"].setCurrentText("居中")
            self.format_combos["图|表题"]["line_spacing"].setCurrentText("固定值20pt")
            
        # 参考文献
        if "参考文献" in self.format_combos:
            # 参考文献标题  
            self.set_default_format_value("参考文献", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            
            # 参考文献内容
            self.set_default_format_value("参考文献", "content", "宋体", "五号 (10.5pt)", "左对齐", "固定值20pt")
            
        # 致谢
        if "致谢" in self.format_combos:
            # 致谢标题
            self.set_default_format_value("致谢", "title", "黑体", "小二 (18pt)", "居中", "固定值20pt")
            
            # 致谢内容
            self.set_default_format_value("致谢", "content", "宋体", "小四 (12pt)", "两端对齐", "固定值20pt")
        
    def create_operation_buttons(self):
        """创建操作按钮区域"""
        return
        
    def create_result_area(self):
        """创建检查结果区域"""
        self.visual_result_group = QGroupBox("结果可视化")
        self.visual_result_layout = QVBoxLayout()
        self.visual_result_layout.setContentsMargins(6, 6, 6, 6)
        self.visual_result_layout.setSpacing(6)
        self.visual_layout = self.visual_result_layout
        self.visual_result_group.setLayout(self.visual_result_layout)

        self.text_result_group = QGroupBox("检查结果（结果仅供参考）")
        self.text_result_layout = QVBoxLayout()
        self.result_scroll = QScrollArea()
        self.result_widget = QWidget()
        self.result_layout = QVBoxLayout(self.result_widget)
        self.result_layout.setContentsMargins(0, 0, 0, 0)
        self.result_scroll.setWidget(self.result_widget)
        self.result_scroll.setWidgetResizable(True)
        self.text_result_layout.addWidget(self.result_scroll)
        self.text_result_group.setLayout(self.text_result_layout)

        self.left_panel_layout.addWidget(self.visual_result_group, 1)
        self.right_panel_layout.addWidget(self.text_result_group, 1)

    @staticmethod
    def clear_qt_layout(layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif child_layout is not None:
                MainWindow.clear_qt_layout(child_layout)
        
    def select_file(self):
        """选择文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文档",
            "",
            "Word文档 (*.docx)"
        )
        if file_path:
            self.file_path_label.setText(self.normalize_display_path(file_path))
            self.sync_pdf_path(file_path)

    def select_pdf(self):
        """选择用于图像侧分析的 PDF 文件。"""
        pdf_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择PDF文件",
            "",
            "PDF文件 (*.pdf)"
        )
        if pdf_path:
            self.pdf_path_label.setText(self.normalize_display_path(pdf_path))

    def sync_pdf_path(self, docx_path: str):
        """自动匹配同名 PDF，方便触发图像侧分析。"""
        if not docx_path:
            self.pdf_path_label.setText("未选择PDF（将自动尝试同名PDF）")
            return

        default_pdf = os.path.splitext(docx_path)[0] + ".pdf"
        if os.path.exists(default_pdf):
            self.pdf_path_label.setText(self.normalize_display_path(default_pdf))
        else:
            self.pdf_path_label.setText("未选择PDF（将自动尝试同名PDF）")

    def get_selected_pdf_path(self):
        """获取当前选中的 PDF 路径。"""
        pdf_path = self.pdf_path_label.text().strip()
        if not pdf_path or pdf_path == "未选择PDF（将自动尝试同名PDF）":
            return None
        return pdf_path

    def ensure_pdf_for_hybrid_check(self, docx_path: str):
        """确保混合检查可用的 PDF 输入。"""
        pdf_path = self.get_selected_pdf_path()
        if pdf_path and os.path.exists(pdf_path) and pdf_path != self.generated_pdf_path:
            return pdf_path

        generated_pdf_path = self.docx_to_pdf_converter.convert(docx_path)
        self.generated_pdf_path = generated_pdf_path
        self.pdf_path_label.setText(self.normalize_display_path(generated_pdf_path))
        return generated_pdf_path
             
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
            pdf_path = self.ensure_pdf_for_hybrid_check(self.file_path_label.text())
            result = self.hybrid_processor.process(
                self.file_path_label.text(),
                pdf_path=pdf_path,
                user_formats=user_formats
            )
            self.display_hybrid_results(result)
        except RuntimeError as e:
            QMessageBox.warning(
                self,
                "PDF转换失败",
                f"{str(e)}\n\n本次将仅执行文档侧检查。",
                QMessageBox.StandardButton.Ok
            )
            try:
                user_formats = self.get_format_settings()
                result = self.hybrid_processor.process(
                    self.file_path_label.text(),
                    pdf_path=None,
                    user_formats=user_formats
                )
                self.display_hybrid_results(result)
            except Exception as inner_e:
                QMessageBox.critical(
                    self,
                    "错误",
                    f"混合检查失败：\n{str(inner_e)}",
                    QMessageBox.StandardButton.Ok
                )
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
        self.clear_qt_layout(self.result_layout)
        self.clear_qt_layout(self.visual_layout)
        self.issue_widget_map = {}
                
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
        self.clear_qt_layout(self.result_layout)
        self.clear_qt_layout(self.visual_layout)
        self.issue_widget_map = {}

        summary = result.get("summary", {})
        engine_counts = result.get("engine_counts", {})
        context_status = result.get("context_status", {})
        issues = self.sort_and_number_issues(result.get("issues", []))

        visualization_widget = self.create_issue_visualization_widget(context_status, issues)
        self.visual_layout.addWidget(visualization_widget)

        overview_widget = QGroupBox("检查概览")
        overview_layout = QFormLayout()
        overview_text = self.build_overview_text(summary, engine_counts, context_status)
        overview_label = QLabel(overview_text)
        overview_label.setWordWrap(True)
        overview_layout.addRow(overview_label)
        overview_widget.setLayout(overview_layout)
        self.result_layout.addWidget(overview_widget)

        if not issues:
            no_issue_label = QLabel("未发现格式问题，但混合检查链路已执行。")
            no_issue_label.setStyleSheet("color: green; font-weight: bold; font-size: 14px;")
            no_issue_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.result_layout.addWidget(no_issue_label)
            return

        issue_widget = QGroupBox("问题列表")
        issue_layout = QVBoxLayout()
        for issue in issues:
            issue_item_widget = self.create_issue_item_widget(issue)
            issue_layout.addWidget(issue_item_widget)
            self.issue_widget_map[issue.get("display_index")] = issue_item_widget
        issue_widget.setLayout(issue_layout)
        self.result_layout.addWidget(issue_widget)
        self.result_layout.addStretch(1)

    @staticmethod
    def shorten_issue_text(text, limit=120):
        text = str(text or "").replace("\n", " ").strip()
        if len(text) > limit:
            return text[:limit] + "..."
        return text

    def resolve_issue_section_label(self, issue):
        metadata = issue.get("metadata") or {}
        section = str(metadata.get("section") or "").strip()
        title = str(issue.get("title") or "").strip()
        rule_id = str(issue.get("rule_id") or "").strip()

        section = self.normalize_issue_section_label(rule_id, section, title)
        if section:
            return section
        return "未分类"

    @staticmethod
    def normalize_issue_section_label(rule_id, section, title):
        rule_id = str(rule_id or "").strip()
        section = str(section or "").strip()
        title = str(title or "").strip()

        if section in {"参考文献", "引文与参考文献"}:
            if ".citation_reference_consistency." in rule_id and title == "引文与参考文献":
                return "正文内容" if section != "参考文献" else "参考文献内容"
            return "参考文献内容"

        if section == "正文":
            return "正文内容"

        if section == "目录":
            return "目录内容"

        if section == "注释":
            return "注释内容"

        if section == "公式":
            return "公式内容"

        if title == "公式编号":
            return "公式编号"

        if title == "公式":
            return "公式内容"

        if title == "引文与参考文献":
            return "正文内容"

        if title == "参考文献":
            return "参考文献内容"

        if title == "目录":
            return "目录内容"

        if title == "注释":
            return "注释内容"

        return section or title

    def resolve_issue_content_text(self, issue):
        metadata = issue.get("metadata") or {}
        original_content = metadata.get("original_content")
        if original_content:
            return self.shorten_issue_text(original_content)

        content = metadata.get("content")
        if content:
            return self.shorten_issue_text(content)

        detail = metadata.get("detail")
        if isinstance(detail, dict):
            para_text = detail.get("段落") or detail.get("参考文献") or detail.get("内容")
            if para_text:
                return self.shorten_issue_text(para_text)

        message = issue.get("message")
        if message:
            return self.shorten_issue_text(message)
        return "未提供"

    def resolve_issue_problem_text(self, issue):
        metadata = issue.get("metadata") or {}
        problem = metadata.get("problem")
        if problem:
            detail = metadata.get("problem_detail")
            if detail:
                return self.shorten_issue_text(f"{problem}：{detail}", limit=200)
            return self.shorten_issue_text(problem, limit=200)

        detail = metadata.get("detail")
        if isinstance(detail, dict):
            failed_items = []
            for k in ("字体", "字号", "对齐方式", "行间距"):
                if detail.get(k) is False:
                    failed_items.append(k)
            if failed_items:
                return f"以下项目可能不符合要求：{', '.join(failed_items)}"

            check_result = detail.get("检查结果")
            if isinstance(check_result, str) and check_result.strip():
                return self.shorten_issue_text(check_result, limit=200)

        message = issue.get("message")
        if message:
            return self.shorten_issue_text(message, limit=200)
        return "未提供"

    def format_issue_block(self, issue):
        index = issue.get("display_index", "")
        severity = str(issue.get("severity", "info")).strip().lower() or "info"
        section = self.resolve_issue_section_label(issue)
        content = self.resolve_issue_content_text(issue)
        problem = self.resolve_issue_problem_text(issue)
        return (
            f"{index}. [{severity}] {section}\n"
            f"出错内容：{content}\n"
            f"错误描述：{problem}"
        )

    def create_issue_item_widget(self, issue):
        widget = QGroupBox()
        widget.setObjectName(f"issue-{issue.get('display_index')}")
        widget.setStyleSheet(
            "QGroupBox {"
            "border: 1px solid #d9d9d9;"
            "border-radius: 6px;"
            "margin-top: 4px;"
            "padding: 6px;"
            "background-color: #ffffff;"
            "}"
        )
        layout = QVBoxLayout()
        content_label = QLabel(self.format_issue_block(issue))
        content_label.setWordWrap(True)
        content_label.setCursor(Qt.CursorShape.PointingHandCursor)
        layout.addWidget(content_label)
        widget.setLayout(layout)
        widget.setCursor(Qt.CursorShape.PointingHandCursor)

        def handle_click(_event, target_issue=issue):
            self.navigate_to_issue(target_issue)

        widget.mousePressEvent = handle_click
        content_label.mousePressEvent = handle_click
        return widget

    @staticmethod
    def sort_and_number_issues(issues):
        def sort_key(issue):
            page = issue.get("page")
            has_page = isinstance(page, int) and page > 0
            rule_id = str(issue.get("rule_id") or "")
            section = str((issue.get("metadata") or {}).get("section") or "")
            title = str(issue.get("title") or "")
            return (
                0 if has_page else 1,
                int(page) if has_page else 10**9,
                rule_id,
                section,
                title,
            )

        sorted_issues = sorted(list(issues or []), key=sort_key)
        numbered = []
        for index, issue in enumerate(sorted_issues, start=1):
            copied = dict(issue)
            copied["display_index"] = index
            numbered.append(copied)
        return numbered

    def build_overview_text(self, summary, engine_counts, context_status):
        """构建检查概览文本。"""
        docx_error = context_status.get("docx_parse_error")
        docx_parts_count = context_status.get("docx_parts_order_count", 0)
        pdf_path = context_status.get("pdf_path")
        pdf_error = context_status.get("pdf_extract_error")

        docx_count = engine_counts.get("docx", 0)
        pdf_count = engine_counts.get("pdf", 0)

        if docx_error:
            docx_status = f"DOCX 规则未完成检查：文档解析失败（{docx_error}）。"
        else:
            docx_status = f"DOCX 规则已完成检查：解析到 {docx_parts_count} 个板块，发现问题 {docx_count}。"

        if not pdf_path:
            pdf_status = "PDF 页面规则未执行检查：未提供可用于页面分析的 PDF。"
        elif pdf_error:
            pdf_status = f"PDF 页面规则未完成检查：页面提取失败（{pdf_error}）。"
        else:
            pdf_status = f"PDF 页面规则已完成检查：发现问题 {pdf_count}。"

        return (
            f"总问题数：{summary.get('total', 0)}\n"
            f"错误：{summary.get('errors', 0)}，警告：{summary.get('warnings', 0)}，提示：{summary.get('infos', 0)}\n"
            f"{docx_status}\n"
            f"{pdf_status}"
        )

    def get_issue_group_name(self, issue):
        """将 pdf/ocr 统一归并为图像侧展示。"""
        source = issue.get("source", "unknown")
        if source == "docx":
            return "docx"
        if source in {"pdf", "ocr"}:
            return "图像侧"
        return source

    def create_issue_visualization_widget(self, context_status, issues):
        """创建错误位置可视化区域。"""
        visual_widget = QWidget()
        visual_layout = QVBoxLayout()
        visual_layout.setContentsMargins(0, 0, 0, 0)
        visual_layout.setSpacing(4)
        visual_widget.setLayout(visual_layout)

        pdf_path = str(context_status.get("pdf_path") or "").strip()
        if not pdf_path or not os.path.exists(pdf_path):
            hint_label = QLabel("当前没有可用于标注的 PDF，无法生成页面可视化。")
            hint_label.setWordWrap(True)
            visual_layout.addWidget(hint_label)
            return visual_widget

        pdf_pages, pdf_error = self.pdf_extractor.extract(pdf_path)
        if pdf_error:
            hint_label = QLabel(f"PDF 页面提取失败，无法生成标注：{pdf_error}")
            hint_label.setWordWrap(True)
            visual_layout.addWidget(hint_label)
            return visual_widget

        visual_issues = self.collect_visual_issues(issues, pdf_pages)

        total_pages = len(pdf_pages)
        page_numbers = list(range(1, total_pages + 1))
        if not page_numbers:
            hint_label = QLabel("当前没有可切换展示的标注页。")
            hint_label.setWordWrap(True)
            visual_layout.addWidget(hint_label)
            return visual_widget

        self.visual_issue_state = {
            "pdf_path": pdf_path,
            "pdf_pages": pdf_pages,
            "issues": visual_issues,
            "page_numbers": page_numbers,
            "total_pages": total_pages,
            "zoom": 1.6,
        }

        control_layout = QHBoxLayout()
        control_layout.setContentsMargins(0, 0, 0, 0)
        control_layout.setSpacing(6)
        page_label = QLabel("页码：")
        page_spin = QSpinBox()
        page_spin.setMinimum(min(page_numbers))
        page_spin.setMaximum(max(page_numbers))
        page_spin.setValue(page_numbers[0])
        page_spin.setSingleStep(1)
        control_layout.addWidget(page_label)
        control_layout.addWidget(page_spin)

        zoom_label = QLabel("缩放：")
        zoom_combo = QComboBox()
        zoom_combo.addItems(["适应宽度", "40%", "50%", "60%", "70%", "80%", "100%", "120%", "150%"])
        zoom_combo.setCurrentText("适应宽度")
        control_layout.addWidget(zoom_label)
        control_layout.addWidget(zoom_combo)

        page_issue_label = QLabel()
        page_issue_label.setWordWrap(True)
        control_layout.addWidget(page_issue_label, 1)
        visual_layout.addLayout(control_layout)

        image_scroll = QScrollArea()
        image_scroll.setWidgetResizable(True)
        image_label = QLabel("正在生成标注预览...")
        image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        image_label.setScaledContents(False)
        image_scroll.setWidget(image_label)
        visual_layout.addWidget(image_scroll, 1)

        page_issue_panel = QWidget()
        page_issue_panel_layout = QVBoxLayout(page_issue_panel)
        page_issue_panel_layout.setContentsMargins(0, 18, 0, 0)
        page_issue_panel_layout.setSpacing(4)

        page_issue_title = QLabel("本页问题定位")
        page_issue_title.setStyleSheet("font-weight: bold;")
        page_issue_title.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        page_issue_panel_layout.addWidget(page_issue_title)

        page_issue_list = QListWidget()
        page_issue_list.setMaximumHeight(104)
        page_issue_list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        page_issue_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        page_issue_list.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        page_issue_panel_layout.addWidget(page_issue_list)

        page_issue_panel.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        visual_layout.addWidget(page_issue_panel)

        self.visual_issue_state["image_scroll"] = image_scroll
        self.visual_issue_state["image_label"] = image_label
        self.visual_issue_state["page_issue_list"] = page_issue_list
        self.visual_issue_state["page_spin"] = page_spin
        self.visual_issue_state["zoom_combo"] = zoom_combo
        self.visual_issue_state["fit_to_width"] = True
        viewport = image_scroll.viewport()
        original_resize_event = viewport.resizeEvent

        def wrapped_resize_event(event):
            original_resize_event(event)
            self.refresh_visualization_pixmap()

        viewport.resizeEvent = wrapped_resize_event

        zoom_combo.currentTextChanged.connect(self.handle_visual_zoom_changed)

        visual_widget.setLayout(visual_layout)

        page_spin.valueChanged.connect(
            lambda value: self.update_issue_visualization_page(
                value,
                image_label,
                page_issue_label,
            )
        )

        self.update_issue_visualization_page(page_numbers[0], image_label, page_issue_label)
        return visual_widget

    @staticmethod
    def normalize_match_text(text):
        return re.sub(r"[\s\u3000]+", "", str(text or "")).strip()

    @staticmethod
    def group_page_lines(page):
        rows = []
        spans = sorted(
            [span for span in getattr(page, "spans", []) if getattr(span, "text", "").strip()],
            key=lambda item: ((item.bbox[1] + item.bbox[3]) / 2, item.bbox[0]),
        )
        for span in spans:
            bbox = getattr(span, "bbox", None) or []
            if len(bbox) != 4:
                continue
            x0, y0, x1, y1 = [float(value) for value in bbox]
            center_y = (y0 + y1) / 2
            target_row = None
            for row in rows:
                if abs(center_y - row["center_y"]) <= 4.0:
                    target_row = row
                    break
            if target_row is None:
                target_row = {"center_y": center_y, "items": []}
                rows.append(target_row)
            target_row["items"].append({"text": span.text.strip(), "bbox": [x0, y0, x1, y1]})

        lines = []
        for row in rows:
            items = sorted(row["items"], key=lambda item: item["bbox"][0])
            text = "".join(item["text"] for item in items).strip()
            if not text:
                continue
            bbox = [
                min(item["bbox"][0] for item in items),
                min(item["bbox"][1] for item in items),
                max(item["bbox"][2] for item in items),
                max(item["bbox"][3] for item in items),
            ]
            lines.append({"text": text, "bbox": bbox})
        return lines

    def build_pdf_page_roles(self, pdf_pages):
        return build_page_roles(pdf_pages)

    def allowed_roles_for_issue(self, issue):
        section_label = self.resolve_issue_section_label(issue)
        rule_id = str(issue.get("rule_id") or "")
        title = str(issue.get("title") or "")
        if section_label in {"中文摘要标题", "中文摘要内容", "中文关键词标题", "中文关键词内容", "摘要与关键词"}:
            return {PAGE_ROLE_CN_ABSTRACT}
        if section_label in {"英文摘要标题", "英文摘要内容", "英文关键词标题", "英文关键词内容"}:
            return {PAGE_ROLE_EN_ABSTRACT}
        if section_label in {"目录", "目录内容", "目录标题"}:
            return {PAGE_ROLE_CATALOGUE}
        if section_label in {"参考文献标题", "参考文献内容", "致谢标题", "致谢内容", "附录"}:
            return {PAGE_ROLE_BACKMATTER}
        if "citation_superscript" in rule_id or title == "引文标识":
            return {PAGE_ROLE_MAIN}
        if section_label in {"页码", "页眉", "页脚"}:
            return {
                PAGE_ROLE_CN_ABSTRACT,
                PAGE_ROLE_EN_ABSTRACT,
                PAGE_ROLE_MAIN,
                PAGE_ROLE_BACKMATTER,
                PAGE_ROLE_CATALOGUE,
                PAGE_ROLE_OTHER,
            }
        return {PAGE_ROLE_MAIN, PAGE_ROLE_BACKMATTER}

    @staticmethod
    def merge_bboxes(bboxes):
        if not bboxes:
            return None
        return [
            min(item[0] for item in bboxes),
            min(item[1] for item in bboxes),
            max(item[2] for item in bboxes),
            max(item[3] for item in bboxes),
        ]

    def build_line_windows(self, lines, max_window_size=4):
        windows = []
        total = len(lines)
        for start in range(total):
            for size in range(1, max_window_size + 1):
                end = start + size
                if end > total:
                    break
                window_lines = lines[start:end]
                merged_text = "".join(line["text"] for line in window_lines).strip()
                if not merged_text:
                    continue
                merged_bbox = self.merge_bboxes([line["bbox"] for line in window_lines])
                if merged_bbox is None:
                    continue
                windows.append(
                    {
                        "text": merged_text,
                        "bbox": merged_bbox,
                        "line_count": size,
                    }
                )
        return windows

    def iter_issue_candidate_pages(self, issue, pdf_pages, page_roles, *, avoid_catalogue=None):
        allowed_roles = self.allowed_roles_for_issue(issue)
        skip_catalogue = self.should_avoid_catalogue_for_issue(issue) if avoid_catalogue is None else avoid_catalogue
        for page in pdf_pages:
            page_no = int(getattr(page, "page_no", 0) or 0)
            page_role = page_roles.get(page_no, PAGE_ROLE_OTHER)
            if page_role not in allowed_roles:
                continue
            if skip_catalogue and is_catalogue_page(page):
                continue
            yield page, page_no

    def iter_page_windows(self, page, *, max_window_size=4):
        lines = self.group_page_lines(page)
        for window in self.build_line_windows(lines, max_window_size=max_window_size):
            yield window

    @staticmethod
    def choose_better_match(best_match, score, page_no, bbox):
        if best_match is None or score > best_match["score"]:
            return {
                "page": page_no,
                "bbox": [float(value) for value in bbox],
                "score": score,
            }
        return best_match

    def build_issue_search_candidates(self, issue):
        candidates = []
        metadata = issue.get("metadata") or {}
        for value in (
            metadata.get("original_content"),
            metadata.get("content"),
            (metadata.get("detail") or {}).get("段落") if isinstance(metadata.get("detail"), dict) else None,
            (metadata.get("detail") or {}).get("内容") if isinstance(metadata.get("detail"), dict) else None,
        ):
            normalized = self.normalize_match_text(value)
            if normalized and len(normalized) >= 4:
                candidates.append(normalized)
                if len(normalized) > 40:
                    candidates.append(normalized[:80])
                    candidates.append(normalized[-80:])

        message = self.normalize_match_text(issue.get("message"))
        if message:
            if len(message) >= 12:
                candidates.append(message[:120])
                if len(message) > 120:
                    candidates.append(message[-120:])
            candidates.extend(
                snippet
                for snippet in re.split(r"[，。；：,.!?（）()\[\]\-]", message)
                if len(snippet) >= 6 and not snippet.isdigit()
            )
        # 去重并优先短一些、可检索的片段
        unique = []
        seen = set()
        for item in candidates:
            clipped = item[:120]
            if clipped not in seen and not clipped.isdigit():
                seen.add(clipped)
                unique.append(clipped)
        return unique[:10]

    def should_avoid_catalogue_for_issue(self, issue):
        section_label = self.resolve_issue_section_label(issue)
        rule_id = str(issue.get("rule_id") or "")
        title = str(issue.get("title") or "")
        if section_label in {"目录", "目录内容"}:
            return False
        if "catalogue" in rule_id or "toc." in rule_id or title == "目录":
            return False
        return True

    def locate_docx_marker_bbox(self, issue, pdf_pages, page_roles=None):
        metadata = issue.get("metadata") or {}
        marker = str(metadata.get("marker") or "").strip()
        if not marker:
            return None

        page_roles = page_roles or self.build_pdf_page_roles(pdf_pages)
        best_match = None
        normalized_marker = self.normalize_match_text(marker)
        for page, page_no in self.iter_issue_candidate_pages(
            issue,
            pdf_pages,
            page_roles,
            avoid_catalogue=False,
        ):
            for line in self.group_page_lines(page):
                line_text = self.normalize_match_text(line["text"])
                if normalized_marker not in line_text:
                    continue
                start_pos = line_text.find(normalized_marker)
                score = (len(normalized_marker) * 20) - (start_pos * 5)
                best_match = self.choose_better_match(best_match, score, page_no, line["bbox"])
        return best_match

    def locate_reference_entry_boxes(self, issue, pdf_pages, page_roles=None):
        section_label = self.resolve_issue_section_label(issue)
        if section_label != "参考文献内容":
            return []

        metadata = issue.get("metadata") or {}
        raw_content = str(metadata.get("original_content") or metadata.get("content") or "").strip()
        if not raw_content or "[" not in raw_content:
            return []

        entries = [part.strip() for part in raw_content.split("|") if part.strip()]
        if not entries:
            return []

        page_roles = page_roles or self.build_pdf_page_roles(pdf_pages)
        locations = []
        for entry in entries:
            marker_match = re.match(r"^\s*(\[\d+\])", entry)
            marker = marker_match.group(1) if marker_match else None
            normalized_entry = self.normalize_match_text(entry)
            best_match = None
            for page, page_no in self.iter_issue_candidate_pages(issue, pdf_pages, page_roles, avoid_catalogue=False):
                if page_roles.get(page_no) != PAGE_ROLE_BACKMATTER:
                    continue
                for window in self.iter_page_windows(page, max_window_size=6):
                    line_text = self.normalize_match_text(window["text"])
                    if len(line_text) < 4:
                        continue
                    score = None
                    if marker:
                        marker_pos = line_text.find(self.normalize_match_text(marker))
                        if marker_pos < 0:
                            continue
                        score = 1000 - (marker_pos * 10)
                        if normalized_entry and normalized_entry in line_text:
                            score += min(800, len(normalized_entry))
                        if line_text.startswith(self.normalize_match_text(marker)):
                            score += 300
                    elif normalized_entry and normalized_entry in line_text:
                        start_pos = line_text.find(normalized_entry)
                        score = (len(normalized_entry) * 10) - (start_pos * 3)

                    if score is None:
                        continue

                    best_match = self.choose_better_match(best_match, score, page_no, window["bbox"])
            if best_match is not None:
                locations.append(best_match)
        return locations

    def locate_docx_issue_bbox(self, issue, pdf_pages, page_roles=None):
        page_roles = page_roles or self.build_pdf_page_roles(pdf_pages)
        marker_located = self.locate_docx_marker_bbox(issue, pdf_pages, page_roles=page_roles)
        if marker_located is not None:
            return marker_located

        candidates = self.build_issue_search_candidates(issue)
        if not candidates:
            return None

        section_label = self.resolve_issue_section_label(issue)
        allow_header_footer = section_label in {"页码", "页眉", "页脚"}
        best_match = None
        for page, page_no in self.iter_issue_candidate_pages(issue, pdf_pages, page_roles):
            for line in self.iter_page_windows(page):
                line_text = self.normalize_match_text(line["text"])
                if len(line_text) < 4:
                    continue
                if not allow_header_footer:
                    bbox = line["bbox"]
                    page_height = float(getattr(page, "height", 0.0) or 0.0)
                    if page_height > 0:
                        center_y = (bbox[1] + bbox[3]) / 2
                        if center_y <= page_height * 0.10 or center_y >= page_height * 0.88:
                            continue
                for candidate in candidates:
                    if len(candidate) < 4 or len(line_text) < 4:
                        continue
                    start_pos = line_text.find(candidate)
                    if start_pos >= 0:
                        extra_chars = max(0, len(line_text) - len(candidate))
                        score = (
                            (len(candidate) * 10)
                            - (extra_chars * 2)
                            - ((line.get("line_count", 1) - 1) * 6)
                            - (start_pos * 3)
                        )
                        best_match = self.choose_better_match(best_match, score, page_no, line["bbox"])
        return best_match

    def collect_visual_issues(self, issues, pdf_pages):
        """筛选并补齐可用于页面标注的问题。"""
        visual_issues = []
        page_roles = self.build_pdf_page_roles(pdf_pages)
        for issue in issues:
            source = issue.get("source")
            page = issue.get("page")
            bbox = issue.get("bbox")
            if isinstance(page, int) and page > 0 and isinstance(bbox, (list, tuple)) and len(bbox) == 4:
                try:
                    x1, y1, x2, y2 = [float(value) for value in bbox]
                except Exception:
                    x1 = y1 = x2 = y2 = 0.0
                if x2 > x1 and y2 > y1:
                    visual_issue = dict(issue)
                    visual_issue["bbox"] = [x1, y1, x2, y2]
                    visual_issue["visual_source"] = source
                    visual_issues.append(visual_issue)
                    continue

            if source != "docx":
                continue

            reference_locations = self.locate_reference_entry_boxes(issue, pdf_pages, page_roles=page_roles)
            if reference_locations:
                for idx, located in enumerate(reference_locations, start=1):
                    visual_issue = dict(issue)
                    visual_issue["page"] = located["page"]
                    visual_issue["bbox"] = located["bbox"]
                    visual_issue["visual_source"] = "docx_reference_located"
                    visual_issue["visual_instance"] = idx
                    visual_issues.append(visual_issue)
                continue

            located = self.locate_docx_issue_bbox(issue, pdf_pages, page_roles=page_roles)
            if not located:
                continue

            visual_issue = dict(issue)
            visual_issue["page"] = located["page"]
            visual_issue["bbox"] = located["bbox"]
            visual_issue["visual_source"] = "docx_located"
            visual_issues.append(visual_issue)
        return visual_issues

    def update_issue_visualization_page(self, page_no, image_label, page_issue_label):
        """更新当前页的标注图片。"""
        state = self.visual_issue_state or {}
        pdf_path = state.get("pdf_path")
        visual_issues = state.get("issues") or []
        page_issues = [issue for issue in visual_issues if issue.get("page") == page_no]
        error_count = sum(1 for issue in page_issues if str(issue.get("severity")).lower() == "error")
        warning_count = sum(1 for issue in page_issues if str(issue.get("severity")).lower() == "warning")
        info_count = sum(1 for issue in page_issues if str(issue.get("severity")).lower() == "info")
        page_issue_label.setText(
            f"当前页可标注问题数：{len(page_issues)}（错误：{error_count}，警告：{warning_count}，提示：{info_count}）"
        )

        if not pdf_path:
            image_label.setText("未找到可用于标注的 PDF。")
            image_label.setPixmap(QPixmap())
            return

        try:
            pixmap = self.render_pdf_page_with_issue_boxes(
                pdf_path,
                page_no,
                page_issues,
                zoom=float(state.get("zoom") or 1.6),
            )
        except Exception as exc:
            image_label.setText(f"标注预览生成失败：{exc}")
            image_label.setPixmap(QPixmap())
            return

        state["current_pixmap"] = pixmap
        self.refresh_visualization_pixmap()
        self.populate_current_page_issue_buttons(page_issues)

    def render_pdf_page_with_issue_boxes(self, pdf_path, page_no, page_issues, zoom=1.6):
        """将 PDF 页渲染为图片并叠加问题框。"""
        try:
            import fitz  # type: ignore
        except Exception as exc:
            raise RuntimeError("PyMuPDF 未安装，无法生成页面标注预览。") from exc

        with fitz.open(pdf_path) as doc:
            if page_no < 1 or page_no > len(doc):
                raise ValueError(f"页码超出范围：{page_no}")

            page = doc.load_page(page_no - 1)
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)

        image_format = QImage.Format.Format_RGB888
        qimage = QImage(pix.samples, pix.width, pix.height, pix.stride, image_format).copy()
        pixmap = QPixmap.fromImage(qimage)

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        font = QFont()
        font.setPointSize(10)
        painter.setFont(font)

        for issue in page_issues:
            color = self.get_issue_draw_color(issue.get("severity"))
            pen = QPen(color)
            pen.setWidth(3 if str(issue.get("severity")).lower() == "error" else 2)
            painter.setPen(pen)

            x1, y1, x2, y2 = issue["bbox"]
            sx1 = int(round(x1 * zoom))
            sy1 = int(round(y1 * zoom))
            sx2 = int(round(x2 * zoom))
            sy2 = int(round(y2 * zoom))
            painter.drawRect(sx1, sy1, max(1, sx2 - sx1), max(1, sy2 - sy1))

            label = f"{issue.get('display_index', '')}. {self.resolve_issue_section_label(issue)}"
            text_y = max(18, sy1 - 6)
            painter.fillRect(
                sx1,
                text_y - 16,
                min(320, max(80, len(label) * 12)),
                18,
                QColor(255, 255, 255, 110),
            )
            painter.drawText(sx1 + 4, text_y - 2, label)

        painter.end()
        return pixmap

    def refresh_visualization_pixmap(self):
        """按可视化区域宽度自适应显示标注图片，保持长宽比不变。"""
        state = self.visual_issue_state or {}
        pixmap = state.get("current_pixmap")
        image_label = state.get("image_label")
        image_scroll = state.get("image_scroll")
        if not isinstance(pixmap, QPixmap) or image_label is None or image_scroll is None:
            return

        if state.get("fit_to_width", True):
            viewport_width = max(200, image_scroll.viewport().width() - 12)
            scaled = pixmap.scaledToWidth(viewport_width, Qt.TransformationMode.SmoothTransformation)
        else:
            zoom_factor = float(state.get("display_zoom_factor") or 1.0)
            target_width = max(1, int(round(pixmap.width() * zoom_factor)))
            scaled = pixmap.scaledToWidth(target_width, Qt.TransformationMode.SmoothTransformation)
        image_label.setPixmap(scaled)
        image_label.resize(scaled.size())

    def handle_visual_zoom_changed(self, value):
        state = self.visual_issue_state or {}
        text = str(value or "").strip()
        if text == "适应宽度":
            state["fit_to_width"] = True
            state["display_zoom_factor"] = 1.0
        else:
            state["fit_to_width"] = False
            try:
                state["display_zoom_factor"] = max(0.1, float(text.rstrip("%")) / 100.0)
            except Exception:
                state["display_zoom_factor"] = 1.0
        self.refresh_visualization_pixmap()

    def populate_current_page_issue_buttons(self, page_issues):
        state = self.visual_issue_state or {}
        page_issue_list = state.get("page_issue_list")
        if page_issue_list is None:
            return

        page_issue_list.clear()

        if not page_issues:
            item = QListWidgetItem("当前页没有可标注问题。")
            page_issue_list.addItem(item)
            return

        for issue in page_issues:
            index = issue.get("display_index", "")
            section = self.resolve_issue_section_label(issue)
            item = QListWidgetItem(f"{index}. {section}")
            item.setData(Qt.ItemDataRole.UserRole, index)
            page_issue_list.addItem(item)

        try:
            page_issue_list.itemClicked.disconnect()
        except Exception:
            pass
        page_issue_list.itemClicked.connect(self.handle_page_issue_item_clicked)

    def handle_page_issue_item_clicked(self, item):
        issue_index = item.data(Qt.ItemDataRole.UserRole)
        if issue_index is None:
            return
        self.scroll_to_issue(issue_index)

    def navigate_to_issue(self, issue):
        state = self.visual_issue_state or {}
        page_spin = state.get("page_spin")
        target_page = None
        for visual_issue in state.get("issues") or []:
            if visual_issue.get("display_index") == issue.get("display_index"):
                mapped_page = visual_issue.get("page")
                if isinstance(mapped_page, int) and mapped_page > 0:
                    target_page = mapped_page
                    break

        if not (isinstance(target_page, int) and target_page > 0):
            raw_page = issue.get("page")
            if isinstance(raw_page, int) and raw_page > 0:
                target_page = raw_page

        if isinstance(target_page, int) and target_page > 0 and page_spin is not None:
            page_spin.setValue(target_page)

    def scroll_to_issue(self, issue_index):
        target_widget = self.issue_widget_map.get(issue_index)
        if target_widget is None:
            return
        self.result_scroll.ensureWidgetVisible(target_widget, 20, 20)

    @staticmethod
    def get_issue_draw_color(severity):
        severity = str(severity or "info").strip().lower()
        if severity == "error":
            return QColor(220, 53, 69)
        if severity == "warning":
            return QColor(255, 193, 7)
        return QColor(13, 110, 253)

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
