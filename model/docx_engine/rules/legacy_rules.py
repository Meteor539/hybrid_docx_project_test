from model.compat.legacy_format_adapter import LegacyFormatAdapter
from model.core.base_rule import BaseRule
from model.core.context import RuleContext
from model.core.issue import Issue, Severity, Source


class LegacyFormatCheckRule(BaseRule):
    rule_id = "legacy.docx.format_check"
    display_name = "旧版格式检查器复用"
    spec_ref = "撰写规范（3）（4）（5）（19）等"
    engine = "docx"

    def __init__(self) -> None:
        self.adapter = LegacyFormatAdapter()

    def check(self, ctx: RuleContext) -> list[Issue]:
        parse_error = ctx.extras.get("docx_parse_error")
        if parse_error:
            return [
                Issue(
                    rule_id=self.rule_id,
                    title="文档解析失败",
                    message=f"无法执行旧版格式检查：{parse_error}",
                    severity=Severity.ERROR,
                    source=Source.DOCX,
                    fixable=False,
                )
            ]

        if not ctx.docx_sections:
            return [
                Issue(
                    rule_id=self.rule_id,
                    title="未获取到可检查内容",
                    message="文档解析结果为空，旧版格式检查未执行。",
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                )
            ]

        user_formats = ctx.extras.get("user_formats")
        raw_results = self.adapter.check_sections(ctx.docx_sections, user_formats=user_formats)
        ctx.extras["legacy_format_results"] = raw_results

        failures = self.adapter.extract_failures(raw_results)
        issues: list[Issue] = []
        for index, item in enumerate(failures, start=1):
            path = item.get("path", "root")
            detail = item.get("detail", {})
            detail_text = self._build_detail_text(path, detail)
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.{index}",
                    title="格式不匹配",
                    message=detail_text,
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={"path": path, "detail": detail},
                )
            )

        return issues

    @staticmethod
    def _build_detail_text(path: str, detail: dict) -> str:
        """把旧版检查器的原始结果转成更易读的提示文本。"""
        if not isinstance(detail, dict):
            return f"路径 {path} 存在格式问题。"

        paragraph = str(detail.get("段落") or detail.get("参考文献") or "").strip()
        paragraph = paragraph.replace("\n", " ")
        if len(paragraph) > 60:
            paragraph = f"{paragraph[:60]}..."

        failed_items: list[str] = []
        for key in ("字体", "字号", "对齐方式", "行间距"):
            if detail.get(key) is False:
                failed_items.append(key)

        check_result = detail.get("检查结果")
        if isinstance(check_result, str) and ("无误" not in check_result and "匹配" not in check_result):
            failed_items.append("检查结果")

        message = f"路径 {path} 存在格式问题"
        if failed_items:
            message += f"（问题项：{', '.join(failed_items)}）"
        if paragraph:
            message += f"；段落：{paragraph}"
        if isinstance(check_result, str) and check_result:
            message += f"；说明：{check_result}"
        return message + "。"


class LegacyOrderCheckRule(BaseRule):
    rule_id = "legacy.docx.order_check"
    display_name = "旧版排版顺序检查复用"
    spec_ref = "撰写规范（一）内容顺序"
    engine = "docx"

    def __init__(self) -> None:
        self.adapter = LegacyFormatAdapter()
        self.part_name_map = {
            "cover": "封面",
            "statement1": "原创性声明",
            "statement2": "使用授权声明",
            "chinese_abstract": "中文摘要",
            "english_abstract": "英文摘要",
            "main_text": "正文",
            "references": "参考文献",
            "acknowledgments": "致谢",
        }

    def check(self, ctx: RuleContext) -> list[Issue]:
        parts_order = ctx.extras.get("docx_parts_order")
        if not isinstance(parts_order, list):
            return []

        order_result = self.adapter.check_order(parts_order)
        ctx.extras["legacy_order_results"] = order_result

        issues: list[Issue] = []
        if order_result["missing"]:
            names = [self.part_name_map.get(x, x) for x in order_result["missing"]]
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.missing",
                    title="可能缺失章节",
                    message=f"顺序检查发现可能缺失：{', '.join(names)}",
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={"missing": order_result["missing"]},
                )
            )

        if order_result["illegal"]:
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.illegal",
                    title="出现未识别章节标记",
                    message=f"检测到未识别章节键：{', '.join(order_result['illegal'])}",
                    severity=Severity.INFO,
                    source=Source.DOCX,
                    fixable=False,
                    metadata={"illegal": order_result["illegal"]},
                )
            )

        for idx, err in enumerate(order_result["order_errors"], start=1):
            current = self.part_name_map.get(err["current"], err["current"])
            nxt = self.part_name_map.get(err["next"], err["next"])
            issues.append(
                Issue(
                    rule_id=f"{self.rule_id}.order.{idx}",
                    title="排版顺序可能错误",
                    message=f"检测到顺序异常：{current} -> {nxt}",
                    severity=Severity.WARNING,
                    source=Source.DOCX,
                    fixable=False,
                    metadata=err,
                )
            )

        return issues
