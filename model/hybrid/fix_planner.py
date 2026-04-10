from model.core.context import RuleContext
from model.core.issue import Issue


class FixPlanner:
    """
    修复阶段在最小骨架中默认关闭。
    """

    def apply(self, ctx: RuleContext, issues: list[Issue]) -> None:
        return
