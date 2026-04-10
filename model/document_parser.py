import re
from docx import Document

try:
    from PyQt6.QtCore import Qt
    from PyQt6.QtWidgets import QMessageBox
except Exception:  # noqa: BLE001
    Qt = None
    QMessageBox = None


class DocumentParser:
    def __init__(self):
        self.doc = None
        self.sections = self._empty_sections()
        self.parts_order: list[str] = []
        self.message_box = None

    def _empty_sections(self):
        return {
            "cover": {"school": None, "title": None, "personal_information": None},
            "statement": {"title": [], "content": []},
            "abstract_keyword": {
                "chinese_title": None,
                "chinese_content": None,
                "chinese_keyword_title": None,
                "chinese_keyword": None,
                "english_title": None,
                "english_content": None,
                "english_keyword_title": None,
                "english_keyword": None,
            },
            "catalogue": {"title": None, "content": None},
            "main_text": [],
            "headings": {"chapter": [], "level1": [], "level2": [], "level3": []},
            "figures_or_tables_title": [],
            "figures": [],
            "tables": [],
            "references": {"title": None, "content": []},
            "acknowledgments": {"title": None, "content": None},
        }

    def parse_document(self, file_path):
        self.doc = Document(file_path)
        self.sections = self._empty_sections()
        self.parts_order = []

        paragraphs = [p for p in self.doc.paragraphs if p.text and p.text.strip()]
        if not paragraphs:
            return self.sections

        idx_statement = self._find_index(paragraphs, ["原创性声明", "使用授权"])
        idx_cn_abs = self._find_index(paragraphs, ["摘要"])
        idx_en_abs = self._find_index(paragraphs, ["abstract"], lower=True)
        idx_catalogue = self._find_index(paragraphs, ["目录"])
        idx_main = self._find_chapter_index(paragraphs)
        idx_ref = self._find_index(paragraphs, ["参考文献", "references", "bibliography"], lower=True)
        idx_ack = self._find_index(paragraphs, ["致谢", "acknowledg"], lower=True)

        cover_end = self._first_valid_index(
            [idx_statement, idx_cn_abs, idx_catalogue, idx_main, idx_ref, idx_ack],
            len(paragraphs),
        )
        cover = paragraphs[:cover_end]
        if cover:
            self.parts_order.append("cover")
            self.sections["cover"]["school"] = cover[0]
            if len(cover) > 1:
                self.sections["cover"]["title"] = cover[1]
            if len(cover) > 2:
                self.sections["cover"]["personal_information"] = cover[2:]

        if idx_statement is not None:
            self.parts_order.append("statement1")
            st_end = self._first_index_after(
                idx_statement,
                [idx_cn_abs, idx_en_abs, idx_catalogue, idx_main, idx_ref, idx_ack],
                len(paragraphs),
            )
            st = paragraphs[idx_statement:st_end]
            if st:
                self.sections["statement"]["title"] = [st[0]]
                if len(st) > 1:
                    self.sections["statement"]["content"] = st[1:]

        if idx_cn_abs is not None:
            self.parts_order.append("chinese_abstract")
            cn_end = self._first_index_after(
                idx_cn_abs, [idx_en_abs, idx_catalogue, idx_main, idx_ref, idx_ack], len(paragraphs)
            )
            cn = paragraphs[idx_cn_abs:cn_end]
            if cn:
                self.sections["abstract_keyword"]["chinese_title"] = cn[0]
                if len(cn) > 2:
                    self.sections["abstract_keyword"]["chinese_content"] = cn[1:-2]
                    self.sections["abstract_keyword"]["chinese_keyword_title"] = cn[-2]
                    self.sections["abstract_keyword"]["chinese_keyword"] = cn[-1]
                elif len(cn) == 2:
                    self.sections["abstract_keyword"]["chinese_keyword"] = cn[-1]

        if idx_en_abs is not None:
            self.parts_order.append("english_abstract")
            en_end = self._first_index_after(idx_en_abs, [idx_catalogue, idx_main, idx_ref, idx_ack], len(paragraphs))
            en = paragraphs[idx_en_abs:en_end]
            if en:
                self.sections["abstract_keyword"]["english_title"] = en[0]
                if len(en) > 2:
                    self.sections["abstract_keyword"]["english_content"] = en[1:-2]
                    self.sections["abstract_keyword"]["english_keyword_title"] = en[-2]
                    self.sections["abstract_keyword"]["english_keyword"] = en[-1]
                elif len(en) == 2:
                    self.sections["abstract_keyword"]["english_keyword"] = en[-1]

        if idx_catalogue is not None:
            self.parts_order.append("catalogue")
            catalogue_end = self._first_index_after(idx_catalogue, [idx_main, idx_ref, idx_ack], len(paragraphs))
            cat = paragraphs[idx_catalogue:catalogue_end]
            if cat:
                self.sections["catalogue"]["title"] = cat[0]
                if len(cat) > 1:
                    self.sections["catalogue"]["content"] = cat[1:]

        main_start = idx_main if idx_main is not None else 0
        main_end = self._first_index_after(main_start, [idx_ref, idx_ack], len(paragraphs))
        if main_start < main_end:
            self.parts_order.append("main_text")
            for para in paragraphs[main_start:main_end]:
                text = para.text.strip()
                if self._is_chapter_title(text):
                    self.sections["headings"]["chapter"].append(para)
                elif self._is_level2_title(text):
                    self.sections["headings"]["level2"].append(para)
                elif self._is_level1_title(text):
                    self.sections["headings"]["level1"].append(para)
                elif self._is_level3_title(text):
                    self.sections["headings"]["level3"].append(para)
                elif self._is_figure_or_table_title(text):
                    self.sections["figures_or_tables_title"].append(para)
                else:
                    self.sections["main_text"].append(para)

        if idx_ref is not None:
            self.parts_order.append("references")
            ref_end = self._first_index_after(idx_ref, [idx_ack], len(paragraphs))
            refs = paragraphs[idx_ref:ref_end]
            if refs:
                self.sections["references"]["title"] = refs[0]
                if len(refs) > 1:
                    self.sections["references"]["content"] = refs[1:]

        if idx_ack is not None:
            self.parts_order.append("acknowledgments")
            ack = paragraphs[idx_ack:]
            if ack:
                self.sections["acknowledgments"]["title"] = ack[0]
                if len(ack) > 1:
                    self.sections["acknowledgments"]["content"] = ack[1:]

        self.sections["figures"] = self.doc.inline_shapes
        self.sections["tables"] = self.doc.tables
        return self.sections

    def check_order(self):
        correct_orders = [
            "cover",
            "statement1",
            "statement2",
            "chinese_abstract",
            "english_abstract",
            "main_text",
            "references",
            "acknowledgments",
        ]
        part_name = {
            "cover": "封面",
            "statement1": "原创性声明",
            "statement2": "使用授权声明",
            "chinese_abstract": "中文摘要",
            "english_abstract": "英文摘要",
            "main_text": "正文",
            "references": "参考文献",
            "acknowledgments": "致谢",
        }

        missing = [p for p in correct_orders if p not in self.parts_order]
        if missing:
            self.show_message("可能缺失的部分: " + ", ".join(part_name.get(p, p) for p in missing))
            return

        for i in range(len(self.parts_order) - 1):
            current = self.parts_order[i]
            nxt = self.parts_order[i + 1]
            if current in correct_orders and nxt in correct_orders:
                if correct_orders.index(current) > correct_orders.index(nxt):
                    self.show_message(
                        f"可能错误的排版顺序: {part_name.get(current, current)} -> {part_name.get(nxt, nxt)}"
                    )
                    return

        self.show_message("排版顺序检查通过，各部分顺序正确。")

    def show_message(self, message_text):
        if QMessageBox is None or Qt is None:
            print(message_text)
            return
        self.message_box = QMessageBox()
        self.message_box.setWindowTitle("排版顺序检查")
        self.message_box.setText(message_text)
        self.message_box.setIcon(QMessageBox.Icon.Information)
        self.message_box.setMinimumSize(800, 800)
        self.message_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        self.message_box.setWindowModality(Qt.WindowModality.NonModal)
        self.message_box.show()

    @staticmethod
    def _find_index(paragraphs, keywords, lower=False):
        for idx, para in enumerate(paragraphs):
            text = para.text.strip()
            source = text.lower() if lower else text
            keys = [k.lower() for k in keywords] if lower else keywords
            if any(k in source for k in keys):
                return idx
        return None

    @staticmethod
    def _find_chapter_index(paragraphs):
        for idx, para in enumerate(paragraphs):
            if DocumentParser._is_chapter_title(para.text.strip()):
                return idx
        return None

    @staticmethod
    def _first_valid_index(values, default):
        valid = [v for v in values if isinstance(v, int)]
        return min(valid) if valid else default

    @staticmethod
    def _first_index_after(base, values, default):
        valid = [v for v in values if isinstance(v, int) and v > base]
        return min(valid) if valid else default

    @staticmethod
    def _is_chapter_title(text):
        return bool(re.match(r"^第[一二三四五六七八九十百千0-9]+章", text) or re.match(r"^chapter\s+\d+", text.lower()))

    @staticmethod
    def _is_level1_title(text):
        return bool(re.match(r"^\d+\.\d+\s*", text) or re.match(r"^[一二三四五六七八九十]+、", text))

    @staticmethod
    def _is_level2_title(text):
        return bool(re.match(r"^\d+\.\d+\.\d+\s*", text) or re.match(r"^（[一二三四五六七八九十]+）", text))

    @staticmethod
    def _is_level3_title(text):
        return bool(re.match(r"^\d+\.\s*", text) or re.match(r"^\d+\)", text))

    @staticmethod
    def _is_figure_or_table_title(text):
        return bool(re.match(r"^(图|表)\s*\d+([\.．]\d+)*", text))

