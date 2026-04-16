from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


def _cm_value(length_obj) -> float | None:
    if length_obj is None:
        return None
    try:
        return float(length_obj.cm)
    except Exception:  # noqa: BLE001
        return None


def _approx_equal(value: float | None, target: float, tolerance: float = 0.15) -> bool:
    if value is None:
        return False
    return abs(value - target) <= tolerance


class PageSettingsRule(BaseRule):
    rule_id = "docx.stage2.page_settings"
    display_name = "Page settings"
    spec_ref = "撰写规范（6）页面设置及（2）打印规范"
    engine = "docx"

    def check(self, ctx: RuleContext) -> list[Issue]:
        if ctx.extras.get("docx_parse_error"):
            return []

        doc = ctx.docx_obj
        if doc is None:
            return []

        sections = getattr(doc, "sections", None) or []
        if not sections:
            return []

        section = sections[0]
        actual = {
            "纸张宽度": _cm_value(getattr(section, "page_width", None)),
            "纸张高度": _cm_value(getattr(section, "page_height", None)),
            "上边距": _cm_value(getattr(section, "top_margin", None)),
            "下边距": _cm_value(getattr(section, "bottom_margin", None)),
            "左边距": _cm_value(getattr(section, "left_margin", None)),
            "右边距": _cm_value(getattr(section, "right_margin", None)),
            "页眉距离": _cm_value(getattr(section, "header_distance", None)),
            "页脚距离": _cm_value(getattr(section, "footer_distance", None)),
        }
        expected = {
            "纸张宽度": 21.0,
            "纸张高度": 29.7,
            "上边距": 2.5,
            "下边距": 2.0,
            "左边距": 2.5,
            "右边距": 2.0,
            "页眉距离": 2.6,
            "页脚距离": 2.4,
        }

        failed_items: list[str] = []
        for key, target in expected.items():
            if not _approx_equal(actual.get(key), target):
                value = actual.get(key)
                if value is None:
                    failed_items.append(f"{key}（未读取到）")
                else:
                    failed_items.append(f"{key}（当前约 {value:.2f}cm，规范为 {target:.2f}cm）")

        if not failed_items:
            return []

        return [
            Issue(
                rule_id=self.rule_id,
                title="页面设置",
                message="；".join(failed_items),
                severity=Severity.WARNING,
                source=Source.DOCX,
                fixable=False,
                metadata={
                    "section": "页面设置",
                    "content": "首节页面属性",
                    "problem": "部分页面设置可能不符合规范",
                    "actual": actual,
                    "expected": expected,
                },
            )
        ]
