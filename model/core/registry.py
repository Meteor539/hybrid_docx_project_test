from collections import defaultdict

from model.core.base_rule import BaseRule


class RuleRegistry:
    def __init__(self) -> None:
        self._rules: dict[str, BaseRule] = {}

    def register(self, rule: BaseRule) -> None:
        if not rule.rule_id:
            raise ValueError("rule_id is required")
        self._rules[rule.rule_id] = rule

    def get(self, rule_id: str) -> BaseRule | None:
        return self._rules.get(rule_id)

    def enable(self, rule_id: str, enabled: bool = True) -> None:
        rule = self.get(rule_id)
        if rule is not None:
            rule.enabled = enabled

    def all(self) -> list[BaseRule]:
        return list(self._rules.values())

    def by_engine(self, engine: str) -> list[BaseRule]:
        return [r for r in self._rules.values() if r.engine == engine and r.enabled]

    def grouped_enabled(self) -> dict[str, list[BaseRule]]:
        grouped: dict[str, list[BaseRule]] = defaultdict(list)
        for rule in self._rules.values():
            if rule.enabled:
                grouped[rule.engine].append(rule)
        return dict(grouped)

