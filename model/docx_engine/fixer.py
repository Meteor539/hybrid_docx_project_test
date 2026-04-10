from model.core.context import RuleContext
from model.core.issue import Issue


class DocxFixer:
    """
    自动格式修复的预留扩展点。
    当前骨架版本默认关闭。
    """

    def apply(self, ctx: RuleContext, issues: list[Issue]) -> None:
        return
