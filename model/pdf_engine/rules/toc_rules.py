from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class TocPresencePdfRule(BaseRule):
    rule_id = "toc.exists_and_match"
    display_name = "TOC presence (pdf)"
    spec_ref = "撰写规范（8）目录"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []
        all_text = "\n".join(getattr(p, "text", "") for p in pages)
        if "目录" in all_text:
            return []
        return [
            Issue(
                rule_id=self.rule_id,
                title="目录标题未识别",
                message="PDF 文本中未识别到“目录”关键词，建议人工确认目录页。",
                severity=Severity.INFO,
                source=Source.PDF,
                fixable=False,
            )
        ]

