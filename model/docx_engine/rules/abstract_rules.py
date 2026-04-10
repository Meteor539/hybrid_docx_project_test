from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class AbstractSectionPresenceRule(BaseRule):
    rule_id = "abstract_cn_en.font_layout"
    display_name = "Abstract section presence"
    spec_ref = "撰写规范（2）（3）中英文摘要与关键词"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        sections = ctx.docx_sections or {}
        abstract = sections.get("abstract_keyword", {})
        if not isinstance(abstract, dict):
            abstract = {}

        cn_title = abstract.get("chinese_title")
        en_title = abstract.get("english_title")
        if cn_title is not None and en_title is not None:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="中英文摘要可能不完整",
                message="docx 解析中未同时识别到中英文摘要标题。",
                severity=Severity.WARNING,
                source=Source.DOCX,
                page=1,
                fixable=False,
            )
        ]

