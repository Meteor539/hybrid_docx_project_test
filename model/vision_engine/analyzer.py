class OcrAnalyzer:
    """
    OCR 分析器接缝层。

    当前默认混合检查流程未接入该链路；图像侧主流程使用 PDF 页面提取
    与坐标分析。此类仅作为后续真实 OCR 接入的预留入口。
    """

    def analyze_images(self, image_paths: list[str]) -> tuple[list[dict], str | None]:
        if not image_paths:
            return [], None
        return [], "OCR analyzer is not wired in the minimal skeleton"
