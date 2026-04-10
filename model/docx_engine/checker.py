from model.core.context import RuleContext
from model.core.issue import Issue
from model.core.registry import RuleRegistry


class DocxRuleEngine:
    def __init__(self, registry: RuleRegistry) -> None:
        self.registry = registry

    def run(self, ctx: RuleContext) -> list[Issue]:
        issues: list[Issue] = []
        for rule in self.registry.by_engine("docx"):
            issues.extend(rule.check(ctx))
        return issues

