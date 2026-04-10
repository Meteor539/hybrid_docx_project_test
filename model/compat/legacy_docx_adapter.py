from typing import Any

import re
from docx import Document


class LegacyDocxAdapter:
    """对现有 python-docx 解析链路的适配器。"""

    def __init__(self) -> None:
        self.parser = None
        self.parts_order: list[str] = []

    def parse(self, file_path: str) -> tuple[Any, dict[str, Any]]:
        try:
            if self.parser is None:
                from model.document_parser import DocumentParser

                self.parser = DocumentParser()
            sections = self.parser.parse_document(file_path)
            self.parts_order = list(getattr(self.parser, "parts_order", []) or [])
            return self.parser.doc, sections
        except Exception:
            # 旧解析器不可用时，回退到纯 python-docx 解析。
            return self._fallback_parse(file_path)

    def _fallback_parse(self, file_path: str) -> tuple[Any, dict[str, Any]]:
        doc = Document(file_path)
        paragraphs = [p for p in doc.paragraphs if p.text and p.text.strip()]

        sections = self._empty_sections()
        self.parts_order = []

        if not paragraphs:
            return doc, sections

        # 粗粒度边界识别：声明、摘要、正文、参考文献、致谢
        idx_statement = self._find_index(paragraphs, ["原创性声明", "使用授权"])
        idx_cn_abs = self._find_index(paragraphs, ["摘要"])
        idx_en_abs = self._find_index(paragraphs, ["abstract"], lower=True)
        idx_main = self._find_chapter_index(paragraphs)
        idx_ref = self._find_index(paragraphs, ["参考文献", "references", "bibliography"], lower=True)
        idx_ack = self._find_index(paragraphs, ["致谢", "acknowledg"], lower=True)

        # 封面
        cover_end = self._first_valid_index(
            [idx_statement, idx_cn_abs, idx_main, idx_ref, idx_ack],
            default=len(paragraphs),
        )
        cover = paragraphs[:cover_end]
        if cover:
            self.parts_order.append("cover")
            sections["cover"]["school"] = cover[0]
            if len(cover) > 1:
                sections["cover"]["title"] = cover[1]
            if len(cover) > 2:
                sections["cover"]["personal_information"] = cover[2:]

        # 声明（简单聚合）
        if idx_statement is not None:
            self.parts_order.append("statement1")
            st_end = self._first_index_after(
                idx_statement,
                [idx_cn_abs, idx_en_abs, idx_main, idx_ref, idx_ack],
                default=len(paragraphs),
            )
            st = paragraphs[idx_statement:st_end]
            if st:
                sections["statement"]["title"] = [st[0]]
                if len(st) > 1:
                    sections["statement"]["content"] = st[1:]

        # 中文摘要
        if idx_cn_abs is not None:
            self.parts_order.append("chinese_abstract")
            cn_end = self._first_index_after(
                idx_cn_abs,
                [idx_en_abs, idx_main, idx_ref, idx_ack],
                default=len(paragraphs),
            )
            cn = paragraphs[idx_cn_abs:cn_end]
            if cn:
                sections["abstract_keyword"]["chinese_title"] = cn[0]
                if len(cn) > 2:
                    sections["abstract_keyword"]["chinese_content"] = cn[1:-2]
                    sections["abstract_keyword"]["chinese_keyword_title"] = cn[-2]
                    sections["abstract_keyword"]["chinese_keyword"] = cn[-1]
                elif len(cn) == 2:
                    sections["abstract_keyword"]["chinese_keyword"] = cn[-1]

        # 英文摘要
        if idx_en_abs is not None:
            self.parts_order.append("english_abstract")
            en_end = self._first_index_after(
                idx_en_abs,
                [idx_main, idx_ref, idx_ack],
                default=len(paragraphs),
            )
            en = paragraphs[idx_en_abs:en_end]
            if en:
                sections["abstract_keyword"]["english_title"] = en[0]
                if len(en) > 2:
                    sections["abstract_keyword"]["english_content"] = en[1:-2]
                    sections["abstract_keyword"]["english_keyword_title"] = en[-2]
                    sections["abstract_keyword"]["english_keyword"] = en[-1]
                elif len(en) == 2:
                    sections["abstract_keyword"]["english_keyword"] = en[-1]

        # 正文及标题分类
        main_start = idx_main if idx_main is not None else 0
        main_end = self._first_index_after(
            main_start,
            [idx_ref, idx_ack],
            default=len(paragraphs),
        )
        if main_start < main_end:
            self.parts_order.append("main_text")
            for para in paragraphs[main_start:main_end]:
                text = para.text.strip()
                if self._is_chapter_title(text):
                    sections["headings"]["chapter"].append(para)
                elif self._is_level2_title(text):
                    sections["headings"]["level2"].append(para)
                elif self._is_level1_title(text):
                    sections["headings"]["level1"].append(para)
                elif self._is_level3_title(text):
                    sections["headings"]["level3"].append(para)
                elif self._is_figure_or_table_title(text):
                    sections["figures_or_tables_title"].append(para)
                else:
                    sections["main_text"].append(para)

        # 参考文献
        if idx_ref is not None:
            self.parts_order.append("references")
            ref_end = self._first_index_after(idx_ref, [idx_ack], default=len(paragraphs))
            refs = paragraphs[idx_ref:ref_end]
            if refs:
                sections["references"]["title"] = refs[0]
                if len(refs) > 1:
                    sections["references"]["content"] = refs[1:]

        # 致谢
        if idx_ack is not None:
            self.parts_order.append("acknowledgments")
            ack = paragraphs[idx_ack:]
            if ack:
                sections["acknowledgments"]["title"] = ack[0]
                if len(ack) > 1:
                    sections["acknowledgments"]["content"] = ack[1:]

        sections["figures"] = doc.inline_shapes
        sections["tables"] = doc.tables
        return doc, sections

    @staticmethod
    def _empty_sections() -> dict[str, Any]:
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

    @staticmethod
    def _find_index(paragraphs, keywords: list[str], lower: bool = False):
        for idx, para in enumerate(paragraphs):
            text = para.text.strip()
            candidate = text.lower() if lower else text
            keys = [k.lower() for k in keywords] if lower else keywords
            if any(k in candidate for k in keys):
                return idx
        return None

    @staticmethod
    def _find_chapter_index(paragraphs):
        for idx, para in enumerate(paragraphs):
            if LegacyDocxAdapter._is_chapter_title(para.text.strip()):
                return idx
        return None

    @staticmethod
    def _first_valid_index(values: list[int | None], default: int) -> int:
        valid = [v for v in values if isinstance(v, int)]
        return min(valid) if valid else default

    @staticmethod
    def _first_index_after(base: int, values: list[int | None], default: int) -> int:
        valid = [v for v in values if isinstance(v, int) and v > base]
        return min(valid) if valid else default

    @staticmethod
    def _is_chapter_title(text: str) -> bool:
        return bool(
            re.match(r"^第[一二三四五六七八九十百千0-9]+章", text)
            or re.match(r"^chapter\s+\d+", text.lower())
        )

    @staticmethod
    def _is_level1_title(text: str) -> bool:
        return bool(re.match(r"^\d+\.\d+\s*", text) or re.match(r"^[一二三四五六七八九十]+、", text))

    @staticmethod
    def _is_level2_title(text: str) -> bool:
        return bool(re.match(r"^\d+\.\d+\.\d+\s*", text) or re.match(r"^（[一二三四五六七八九十]+）", text))

    @staticmethod
    def _is_level3_title(text: str) -> bool:
        return bool(re.match(r"^\d+\.\s*", text) or re.match(r"^\d+\)", text))

    @staticmethod
    def _is_figure_or_table_title(text: str) -> bool:
        return bool(re.match(r"^(图|表)\s*\d+([\.．]\d+)*", text))
