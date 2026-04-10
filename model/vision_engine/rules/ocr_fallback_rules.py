from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class OcrFallbackAvailabilityRule(BaseRule):
    rule_id = "ocr.fallback.status"
    display_name = "OCR fallback availability"
    spec_ref = "撰写规范（7）（8）（16）（17）图像兜底链路"
    engine = "ocr"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.ocr_pages:
            return []

        # 以提示级别告知：图像兜底链路当前未启用。
        return [
            Issue(
                rule_id=self.rule_id,
                title="OCR兜底链路未启用",
                message="当前骨架未接入 OCR 分析器，图像兜底规则暂未执行。",
                severity=Severity.INFO,
                source=Source.OCR,
                fixable=False,
            )
        ]
