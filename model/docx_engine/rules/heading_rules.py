from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class MainTextPresenceRule(BaseRule):
    rule_id = "heading.level_font"
    display_name = "Main text and heading presence"
    spec_ref = "撰写规范（3）（9）标题与正文"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        sections = ctx.docx_sections or {}
        headings = sections.get("headings", {})
        main_text = sections.get("main_text", [])
        has_heading = isinstance(headings, dict) and any(
            bool(headings.get(k)) for k in ("chapter", "level1", "level2", "level3")
        )
        has_main = isinstance(main_text, list) and len(main_text) > 0

        if has_heading and has_main:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="正文或标题识别不足",
                message="docx 解析中未稳定识别到正文或分级标题。",
                severity=Severity.INFO,
                source=Source.DOCX,
                page=1,
                fixable=False,
            )
        ]

