from pathlib import Path


class DocumentRenderer:
    """
    docx -> 图片渲染的预留接缝。
    当前骨架仅校验路径，不实际产出图片。
    """

    def render_to_images(self, file_path: str, output_dir: str) -> tuple[list[str], str | None]:
        path = Path(file_path)
        if not path.exists():
            return [], f"Document not found: {file_path}"
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        return [], None
