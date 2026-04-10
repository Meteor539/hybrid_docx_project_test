from abc import ABC, abstractmethod

from model.core.context import RuleContext
from model.core.issue import Issue


class BaseRule(ABC):
    rule_id: str = ""
    display_name: str = ""
    spec_ref: str = ""
    engine: str = ""
    enabled: bool = True

    @abstractmethod
    def check(self, ctx: RuleContext) -> list[Issue]:
        raise NotImplementedError

    def fix(self, ctx: RuleContext, issues: list[Issue]) -> None:
        # 初始骨架阶段默认不执行自动修复。
        return
