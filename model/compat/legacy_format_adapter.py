from typing import Any

from model.format_checker import FormatChecker


class LegacyFormatAdapter:
    """复用现有 FormatChecker 的适配器。"""

    def __init__(self) -> None:
        self.checker = FormatChecker()

    def check_sections(
        self,
        sections: dict[str, Any],
        user_formats: dict[str, dict[str, str]] | None = None,
    ) -> dict[str, Any]:
        if user_formats:
            self.checker.update_formats(user_formats)
        return self.checker.check_format(sections)

    def extract_failures(self, results: dict[str, Any]) -> list[dict[str, Any]]:
        failures: list[dict[str, Any]] = []
        self._walk_result(results, path=[], collector=failures)
        return failures

    def _walk_result(
        self,
        node: Any,
        path: list[str],
        collector: list[dict[str, Any]],
    ) -> None:
        if isinstance(node, list):
            for item in node:
                if isinstance(item, dict):
                    collector.append(
                        {
                            "path": ".".join(path) if path else "root",
                            "detail": item,
                        }
                    )
            return

        if isinstance(node, dict):
            if self._is_success_dict(node):
                return

            # 叶子失败字典：直接作为问题收集
            if self._is_leaf_dict(node):
                collector.append(
                    {
                        "path": ".".join(path) if path else "root",
                        "detail": node,
                    }
                )
                return

            for key, value in node.items():
                self._walk_result(value, [*path, str(key)], collector)

    @staticmethod
    def _is_leaf_dict(data: dict[str, Any]) -> bool:
        return all(not isinstance(v, (dict, list)) for v in data.values())

    @staticmethod
    def _is_success_dict(data: dict[str, Any]) -> bool:
        if len(data) != 1:
            return False
        value = next(iter(data.values()))
        if not isinstance(value, str):
            return False
        return "无误" in value or "匹配" in value

    def check_order(self, parts_order: list[str]) -> dict[str, Any]:
        expected = [
            "cover",
            "statement1",
            "statement2",
            "chinese_abstract",
            "english_abstract",
            "main_text",
            "references",
            "acknowledgments",
        ]
        expected_set = set(expected)

        missing = [x for x in expected if x not in parts_order]
        illegal = [x for x in parts_order if x not in expected_set]

        order_errors: list[dict[str, str]] = []
        filtered = [x for x in parts_order if x in expected_set]
        for i in range(len(filtered) - 1):
            cur = filtered[i]
            nxt = filtered[i + 1]
            if expected.index(cur) > expected.index(nxt):
                order_errors.append({"current": cur, "next": nxt})

        return {
            "missing": missing,
            "illegal": illegal,
            "order_errors": order_errors,
            "parts_order": parts_order,
            "ok": (not missing and not illegal and not order_errors),
        }

