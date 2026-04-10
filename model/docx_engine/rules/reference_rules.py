from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class ReferenceSectionPresenceRule(BaseRule):
    rule_id = "references.font_layout"
    display_name = "Reference section presence"
    spec_ref = "撰写规范（19）参考文献"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        sections = ctx.docx_sections or {}
        references = sections.get("references", {})
        if not isinstance(references, dict):
            references = {}

        title = references.get("title")
        content = references.get("content")
        has_content = isinstance(content, list) and len(content) > 0

        if title is not None and has_content:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="参考文献章节可能缺失",
                message="docx 解析中未识别到参考文献标题或条目内容。",
                severity=Severity.WARNING,
                source=Source.DOCX,
                fixable=False,
            )
        ]

