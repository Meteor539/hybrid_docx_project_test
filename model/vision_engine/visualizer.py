class VisionVisualizer:
    def annotate(self, image_path: str, issues: list[dict]) -> str:
        # 预留接缝：骨架阶段直接返回原图路径。
        return image_path
