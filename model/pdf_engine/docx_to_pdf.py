import os
import platform
import subprocess
import sys
from pathlib import Path


class DocxToPdfConverter:
    """Convert a .docx file to PDF for page-level analysis."""

    def convert(self, docx_path: str, output_path: str | None = None) -> str:
        source = Path(docx_path)
        if not source.exists():
            raise FileNotFoundError(f"Word 文档不存在: {docx_path}")
        if source.suffix.lower() != ".docx":
            raise ValueError(f"仅支持 .docx 文档: {docx_path}")

        target_path = output_path or self._build_temp_pdf_path(source)

        system_name = platform.system()
        if system_name == "Linux":
            self._convert_with_soffice(source, target_path)
        else:
            self._convert_with_docx2pdf(source, target_path)

        if not os.path.exists(target_path) or os.path.getsize(target_path) == 0:
            raise RuntimeError("自动生成 PDF 失败，输出文件不存在或为空。")

        return target_path

    @staticmethod
    def _build_temp_pdf_path(source: Path) -> str:
        return str(source.with_name(f"{source.stem}_hybrid.pdf"))

    @staticmethod
    def _convert_with_soffice(source: Path, target_path: str) -> None:
        output_dir = os.path.dirname(target_path)
        expected_pdf = os.path.join(output_dir, source.with_suffix(".pdf").name)
        try:
            completed = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    output_dir,
                    str(source),
                ],
                check=True,
                capture_output=True,
                text=True,
            )
        except subprocess.CalledProcessError as exc:
            message = exc.stderr or exc.stdout or str(exc)
            raise RuntimeError(
                "自动转 PDF 失败：LibreOffice soffice 调用出错。\n"
                f"{message}"
            ) from exc
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(
                f"自动转 PDF 失败：无法调用 LibreOffice soffice。\n{exc}"
            ) from exc

        if not os.path.exists(expected_pdf):
            stdout = completed.stdout.strip()
            stderr = completed.stderr.strip()
            raise RuntimeError(
                "自动转 PDF 失败：LibreOffice 未生成预期 PDF 文件。\n"
                f"stdout: {stdout}\n"
                f"stderr: {stderr}"
            )

        os.replace(expected_pdf, target_path)

    @staticmethod
    def _convert_with_docx2pdf(source: Path, target_path: str) -> None:
        py_code = (
            "from docx2pdf import convert\n"
            f"convert(r'{str(source)}', r'{target_path}')\n"
        )
        try:
            subprocess.run(
                [sys.executable, "-c", py_code],
                check=True,
                capture_output=True,
                text=True,
            )
        except subprocess.CalledProcessError as exc:
            message = exc.stderr or exc.stdout or str(exc)
            raise RuntimeError(
                "自动转 PDF 失败：docx2pdf 调用出错。\n"
                f"{message}"
            ) from exc
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(
                f"自动转 PDF 失败：无法启动 docx2pdf 子进程。\n{exc}"
            ) from exc
