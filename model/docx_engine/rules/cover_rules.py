from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class CoverTitlePresenceRule(BaseRule):
    rule_id = "cover.font_layout"
    display_name = "Cover title presence"
    spec_ref = "撰写规范（4）论文封面"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        sections = ctx.docx_sections or {}
        cover = sections.get("cover", {})
        title = cover.get("title") if isinstance(cover, dict) else None
        if title is not None:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="封面题目可能缺失",
                message="未在 docx 解析结果中识别到封面题目段落。",
                severity=Severity.WARNING,
                source=Source.DOCX,
                page=1,
                fixable=False,
            )
        ]

