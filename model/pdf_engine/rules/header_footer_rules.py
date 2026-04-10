import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class PageNumberPresencePdfRule(BaseRule):
    rule_id = "page_number.presence"
    display_name = "Page number presence (pdf)"
    spec_ref = "撰写规范（7）页码"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        pattern = re.compile(r"^\d+$")
        issues: list[Issue] = []
        for page in pages:
            found = False
            spans = getattr(page, "spans", [])
            for span in spans[-8:]:
                text = getattr(span, "text", "").strip()
                if pattern.fullmatch(text):
                    found = True
                    break
            if not found:
                issues.append(
                    Issue(
                        rule_id=self.rule_id,
                        title="页码可能缺失",
                        message=f"第 {getattr(page, 'page_no', '?')} 页底部未识别到纯数字页码。",
                        severity=Severity.INFO,
                        source=Source.PDF,
                        page=getattr(page, "page_no", None),
                        fixable=False,
                    )
                )
        return issues

