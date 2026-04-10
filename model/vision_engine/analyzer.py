class OcrAnalyzer:
    """
    OCR 分析器接缝层。
    后续可在此接入旧图像项目的真实能力。
    """

    def analyze_images(self, image_paths: list[str]) -> tuple[list[dict], str | None]:
        if not image_paths:
            return [], None
        return [], "OCR analyzer is not wired in the minimal skeleton"
