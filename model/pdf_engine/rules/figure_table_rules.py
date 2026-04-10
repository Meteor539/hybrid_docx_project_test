import re

from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class FigureTableCaptionHintRule(BaseRule):
    rule_id = "figure_table.caption_hint"
    display_name = "Figure/table caption format hint (pdf)"
    spec_ref = "撰写规范（16）（17）表题与图题"
    engine = "pdf"

    def check(self, ctx: RuleContext) -> list[Issue]:
        pages = ctx.pdf_pages or []
        if not pages:
            return []

        pat_figure = re.compile(r"^图\s*\d+([\.．]\d+)*")
        pat_table = re.compile(r"^表\s*\d+([\.．]\d+)*")

        issues: list[Issue] = []
        for page in pages:
            for span in getattr(page, "spans", []):
                text = getattr(span, "text", "").strip()
                if not text:
                    continue
                if text.startswith("图") and not pat_figure.search(text):
                    issues.append(
                        Issue(
                            rule_id=self.rule_id,
                            title="图题编号格式可疑",
                            message=f"第 {getattr(page, 'page_no', '?')} 页存在“图”开头文本但未匹配常见编号格式。",
                            severity=Severity.INFO,
                            source=Source.PDF,
                            page=getattr(page, "page_no", None),
                            fixable=False,
                        )
                    )
                if text.startswith("表") and not pat_table.search(text):
                    issues.append(
                        Issue(
                            rule_id=self.rule_id,
                            title="表题编号格式可疑",
                            message=f"第 {getattr(page, 'page_no', '?')} 页存在“表”开头文本但未匹配常见编号格式。",
                            severity=Severity.INFO,
                            source=Source.PDF,
                            page=getattr(page, "page_no", None),
                            fixable=False,
                        )
                    )
        return issues

