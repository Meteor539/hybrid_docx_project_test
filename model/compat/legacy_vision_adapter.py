class LegacyVisionAdapter:
    """
    旧图像项目的占位适配器。
    该类作为后续集成的连接点保留。
    """

    def analyze(self, image_paths: list[str]) -> list[dict]:
        # 最小骨架阶段暂未接入真实实现。
        return []
