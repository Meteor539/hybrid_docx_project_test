import re
from docx import Document
from docx.text.paragraph import Paragraph

try:
    from PyQt6.QtCore import Qt
    from PyQt6.QtWidgets import QMessageBox
except Exception:  # noqa: BLE001
    Qt = None
    QMessageBox = None


class DocumentParser:
    _statement_headings = [
        "原创性声明",
        "学位论文原创性声明",
        "学位论文版权使用授权书",
        "版权使用授权书",
        "使用授权书",
        "使用授权声明",
    ]

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

        paragraphs = [p for p in self._collect_body_paragraphs() if p.text and p.text.strip()]
        if not paragraphs:
            return self.sections

        idx_statement = self._find_statement_index(paragraphs)
        idx_cn_abs = self._find_exact_heading_index(paragraphs, ["摘要", "摘 要"])
        idx_en_abs = self._find_exact_heading_index(paragraphs, ["abstract"], lower=True)
        idx_catalogue = self._find_exact_heading_index(paragraphs, ["目录"])
        idx_main = self._find_chapter_index(paragraphs)
        idx_ref = self._find_exact_heading_index(paragraphs, ["参考文献", "references", "bibliography"], lower=True)
        idx_ack = self._find_ack_index(paragraphs)

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
                statement_titles, statement_contents = self._split_statement_block(st)
                self.sections["statement"]["title"] = statement_titles
                self.sections["statement"]["content"] = statement_contents

        if idx_cn_abs is not None:
            self.parts_order.append("chinese_abstract")
            cn_end = self._first_index_after(
                idx_cn_abs, [idx_en_abs, idx_catalogue, idx_main, idx_ref, idx_ack], len(paragraphs)
            )
            cn = paragraphs[idx_cn_abs:cn_end]
            if cn:
                self.sections["abstract_keyword"]["chinese_title"] = cn[0]
                self._assign_abstract_parts(cn, language="chinese")

        if idx_en_abs is not None:
            self.parts_order.append("english_abstract")
            en_end = self._first_index_after(idx_en_abs, [idx_catalogue, idx_main, idx_ref, idx_ack], len(paragraphs))
            en = paragraphs[idx_en_abs:en_end]
            if en:
                self.sections["abstract_keyword"]["english_title"] = en[0]
                self._assign_abstract_parts(en, language="english")

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

    def _collect_body_paragraphs(self):
        """按文档顺序收集正文段落，并补上 sdt 中的自动目录段落。"""
        body = getattr(self.doc, "_body", None)
        body_element = getattr(body, "_element", None) if body is not None else None
        if body_element is None:
            return list(getattr(self.doc, "paragraphs", []) or [])

        paragraphs = []
        for child in body_element.iterchildren():
            paragraphs.extend(self._extract_paragraphs_from_element(child, body))
        return paragraphs

    def _extract_paragraphs_from_element(self, element, parent):
        paragraphs = []
        tag = getattr(element, "tag", "")
        if tag.endswith("}p"):
            paragraphs.append(Paragraph(element, parent))
            return paragraphs

        if tag.endswith("}sdt") or tag.endswith("}sdtContent"):
            for child in element.iterchildren():
                paragraphs.extend(self._extract_paragraphs_from_element(child, parent))
        return paragraphs

    def _assign_abstract_parts(self, paragraphs, *, language: str):
        if not paragraphs:
            return

        content_key = f"{language}_content"
        keyword_title_key = f"{language}_keyword_title"
        keyword_key = f"{language}_keyword"

        body = paragraphs[1:]
        if not body:
            return

        keyword_idx = self._find_keyword_paragraph_index(body, language=language)
        if keyword_idx is not None:
            keyword_paragraph = body[keyword_idx]
            content_paragraphs = body[:keyword_idx]
            self.sections["abstract_keyword"][content_key] = content_paragraphs or None
            # 关键词通常与内容同段，这里把同一段同时登记为标题和内容，供后续细粒度检查拆分。
            self.sections["abstract_keyword"][keyword_title_key] = keyword_paragraph
            self.sections["abstract_keyword"][keyword_key] = keyword_paragraph
            return

        if len(body) >= 2:
            self.sections["abstract_keyword"][content_key] = body[:-1]
            self.sections["abstract_keyword"][keyword_key] = body[-1]
        else:
            self.sections["abstract_keyword"][keyword_key] = body[-1]

    @staticmethod
    def _find_keyword_paragraph_index(paragraphs, *, language: str):
        if language == "english":
            pattern = r"^\s*key\s*words?\s*[:：]"
        else:
            pattern = r"^\s*(关键词|关\s*键\s*词)\s*[:：]"

        for idx, para in enumerate(paragraphs):
            text = (para.text or "").strip()
            if re.match(pattern, text, flags=re.IGNORECASE):
                return idx
        return None

    def check_order(self):
        correct_orders = [
            "cover",
            "statement1",
            "chinese_abstract",
            "english_abstract",
            "catalogue",
            "main_text",
            "references",
            "acknowledgments",
        ]
        part_name = {
            "cover": "封面",
            "statement1": "原创性声明",
            "chinese_abstract": "中文摘要",
            "english_abstract": "英文摘要",
            "catalogue": "目录",
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
    def _normalize_heading_text(text: str, *, lower: bool = False) -> str:
        normalized = re.sub(r"[\s\u3000]+", "", text or "")
        if lower:
            normalized = normalized.lower()
        return normalized

    @classmethod
    def _find_exact_heading_index(cls, paragraphs, headings, lower=False):
        normalized_headings = {
            cls._normalize_heading_text(h, lower=lower)
            for h in headings
        }
        for idx, para in enumerate(paragraphs):
            text = para.text.strip()
            normalized = cls._normalize_heading_text(text, lower=lower)
            if normalized in normalized_headings:
                return idx
        return None

    @classmethod
    def _find_statement_index(cls, paragraphs):
        return cls._find_exact_heading_index(paragraphs, cls._statement_headings)

    @classmethod
    def _is_statement_heading(cls, text: str) -> bool:
        normalized = cls._normalize_heading_text(text)
        normalized_headings = {cls._normalize_heading_text(item) for item in cls._statement_headings}
        return normalized in normalized_headings

    @classmethod
    def _split_statement_block(cls, paragraphs):
        titles = []
        contents = []
        for paragraph in paragraphs or []:
            text = (getattr(paragraph, "text", "") or "").strip()
            if not text:
                continue
            if cls._is_statement_heading(text):
                titles.append(paragraph)
            else:
                contents.append(paragraph)
        return titles, contents

    @classmethod
    def _find_ack_index(cls, paragraphs):
        for idx, para in enumerate(paragraphs):
            normalized = cls._normalize_heading_text(para.text, lower=True)
            if normalized in {"致谢", "acknowledgments", "acknowledgement", "acknowledgments"}:
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
        if not text:
            return False
        normalized = text.strip()
        if re.search(r"[，。；：！？!?]", normalized):
            return False
        if len(normalized) > 40:
            return False
        return bool(
            re.fullmatch(r"第[一二三四五六七八九十百千0-9]+章\s*[^\s，。；：！？!?]{1,30}", normalized)
            or re.fullmatch(r"chapter\s+\d+\s+.{1,30}", normalized, flags=re.IGNORECASE)
        )

    @staticmethod
    def _is_level1_title(text):
        if not text:
            return False
        normalized = text.strip()
        if re.search(r"[，。；：！？!?]", normalized):
            return False
        if len(normalized) > 40:
            return False
        return bool(
            re.fullmatch(r"\d+\.\d+\s*[^\s，。；：！？!?].{0,30}", normalized)
            or re.fullmatch(r"[一二三四五六七八九十]+、[^\s，。；：！？!?].{0,30}", normalized)
        )

    @staticmethod
    def _is_level2_title(text):
        if not text:
            return False
        normalized = text.strip()
        if re.search(r"[，。；：！？!?]", normalized):
            return False
        if len(normalized) > 40:
            return False
        return bool(
            re.fullmatch(r"\d+\.\d+\.\d+\s*[^\s，。；：！？!?].{0,30}", normalized)
            or re.fullmatch(r"（[一二三四五六七八九十]+）[^\s，。；：！？!?].{0,30}", normalized)
        )

    @staticmethod
    def _is_level3_title(text):
        if not text:
            return False
        normalized = text.strip()
        if re.search(r"[，。；：！？!?]", normalized):
            return False
        if len(normalized) > 30:
            return False
        return bool(
            re.fullmatch(r"\d+\.\s*[^\s，。；：！？!?].{0,20}", normalized)
            or re.fullmatch(r"\d+\)[^\s，。；：！？!?].{0,20}", normalized)
        )

    @staticmethod
    def _is_figure_or_table_title(text):
        return bool(re.match(r"^(图|表)\s*\d+([\.．]\d+)*", text))
