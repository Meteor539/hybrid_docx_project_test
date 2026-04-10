from model.core.base_rule import BaseRule
from model.core.registry import RuleRegistry


class RuleRouter:
    ENGINE_ORDER = ["docx", "pdf", "ocr", "hybrid"]

    def __init__(self, registry: RuleRegistry) -> None:
        self.registry = registry

    def select(self) -> dict[str, list[BaseRule]]:
        grouped = self.registry.grouped_enabled()
        return {engine: grouped.get(engine, []) for engine in self.ENGINE_ORDER}

