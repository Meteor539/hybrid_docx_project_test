from pathlib import Path
from typing import Any

from model.core.context import RuleContext
from model.core.issue import Issue
from model.core.merger import IssueMerger
from model.core.registry import RuleRegistry
from model.docx_engine.parser import DocxContextBuilder
from model.docx_engine.rules import (
    AlignmentFormatRule,
    AppendixFormatRule,
    HeadingNumberHierarchyRule,
    NoteMarkerSuperscriptRule,
    NoteMarkerConsistencyRule,
    PageNumberFormatRule,
    ReferenceAuthorCountRule,
    CatalogueHeadingConsistencyRule,
    CatalogueNumberFontRule,
    CaptionFormatRule,
    CitationSuperscriptRule,
    CharacterSpacingFormatRule,
    ChineseAbstractLengthRule,
    CoverTitleLengthRule,
    CitationReferenceConsistencyRule,
    FontSizeFormatRule,
    FooterFormatRule,
    FormulaAlignmentRule,
    FormulaNumberRightAlignedRule,
    FirstStageSectionPresenceRule,
    FormulaNumberFormatRule,
    HeaderFormatRule,
    HeadingCitationRule,
    HeadingPunctuationRule,
    KeywordCountRule,
    LineSpacingFormatRule,
    PageSettingsRule,
    ReferenceEntryFormatRule,
    ReferenceNumberSequenceRule,
    ReferenceCountRule,
    ReferenceTerminalPeriodRule,
    SectionOrderRule,
)
from model.hybrid.fix_planner import FixPlanner
from model.hybrid.router import RuleRouter
from model.pdf_engine.extractor import PdfExtractor
from model.pdf_engine.rules import (
    ChapterStartsNewPagePdfRule,
    CoverTitleCenterPdfRule,
    FormulaNumberRightAlignedPdfRule,
    HeaderStartBoundaryPdfRule,
    FigureCaptionBelowPdfRule,
    FigureCaptionCenterPdfRule,
    FigureTableCaptionHintRule,
    FigureTableSplitAcrossPagesPdfRule,
    HeaderTopContentPdfRule,
    PageNumberBottomCenterPdfRule,
    PageNumberPresencePdfRule,
    PageNumberStyleSequencePdfRule,
    TableCaptionCenterPdfRule,
    TableCaptionAbovePdfRule,
    TocLevelPresentationPdfRule,
    TocPresencePdfRule,
)
from model.vision_engine.analyzer import OcrAnalyzer
from model.vision_engine.rules import OcrFallbackAvailabilityRule


class HybridProcessor:
    def __init__(
        self,
        registry: RuleRegistry,
        router: RuleRouter,
        merger: IssueMerger,
        fix_planner: FixPlanner,
    ) -> None:
        self.registry = registry
        self.router = router
        self.merger = merger
        self.fix_planner = fix_planner

        self.docx_builder = DocxContextBuilder()
        self.pdf_extractor = PdfExtractor()
        self.ocr_analyzer = OcrAnalyzer()

    def process(
        self,
        file_path: str,
        *,
        pdf_path: str | None = None,
        user_formats: dict[str, dict[str, str]] | None = None,
        enable_fix: bool = False,
    ) -> dict[str, Any]:
        ctx = self._build_context(file_path=file_path, pdf_path=pdf_path, user_formats=user_formats)

        grouped_issues: dict[str, list[Issue]] = {}
        for engine, rules in self.router.select().items():
            engine_issues: list[Issue] = []
            for rule in rules:
                try:
                    engine_issues.extend(rule.check(ctx))
                except Exception as exc:  # noqa: BLE001
                    engine_issues.append(
                        Issue(
                            rule_id=f"{rule.rule_id}.runtime_error",
                            title=f"{rule.display_name} 执行失败",
                            message=str(exc),
                            severity=rule_error_severity(),
                            source=rule_error_source(engine),
                            fixable=False,
                            metadata={"engine": engine},
                        )
                    )
            grouped_issues[engine] = filter_display_issues(engine_issues)

        merged = self.merger.merge(grouped_issues)
        if enable_fix:
            self.fix_planner.apply(ctx, merged)

        return {
            "file_path": file_path,
            "summary": build_summary(merged),
            "issues": [issue_to_dict(x) for x in merged],
            "engine_counts": {k: len(v) for k, v in grouped_issues.items()},
            "context_status": build_context_status(ctx.extras),
        }

    def _build_context(
        self,
        *,
        file_path: str,
        pdf_path: str | None,
        user_formats: dict[str, dict[str, str]] | None,
    ) -> RuleContext:
        ctx = self.docx_builder.build(file_path)
        ctx.extras["user_formats"] = user_formats

        resolved_pdf = pdf_path
        if not resolved_pdf:
            default_pdf = str(Path(file_path).with_suffix(".pdf"))
            if Path(default_pdf).exists():
                resolved_pdf = default_pdf

        if resolved_pdf:
            pdf_pages, pdf_error = self.pdf_extractor.extract(resolved_pdf)
            ctx.pdf_pages = pdf_pages
            ctx.extras["pdf_path"] = resolved_pdf
            ctx.extras["pdf_extract_error"] = pdf_error
        else:
            ctx.extras["pdf_path"] = None
            ctx.extras["pdf_extract_error"] = "PDF not provided"

        ocr_pages, ocr_error = self.ocr_analyzer.analyze_images([])
        ctx.ocr_pages = ocr_pages
        ctx.extras["ocr_error"] = ocr_error
        return ctx


def rule_error_severity():
    from model.core.issue import Severity

    return Severity.WARNING


def rule_error_source(engine: str):
    from model.core.issue import Source

    if engine == "docx":
        return Source.DOCX
    if engine == "pdf":
        return Source.PDF
    if engine == "ocr":
        return Source.OCR
    return Source.HYBRID


def issue_to_dict(issue: Issue) -> dict[str, Any]:
    return {
        "rule_id": issue.rule_id,
        "title": issue.title,
        "message": issue.message,
        "severity": issue.severity.value,
        "source": issue.source.value,
        "page": issue.page,
        "bbox": issue.bbox,
        "confidence": issue.confidence,
        "fixable": issue.fixable,
        "fix_action": issue.fix_action,
        "metadata": issue.metadata,
    }


def filter_display_issues(issues: list[Issue]) -> list[Issue]:
    """Filter out status-only placeholder issues from the main result list."""
    return [issue for issue in issues if issue.rule_id not in STATUS_ONLY_RULE_IDS]


def build_summary(issues: list[Issue]) -> dict[str, int]:
    return {
        "total": len(issues),
        "errors": sum(1 for i in issues if i.severity.value == "error"),
        "warnings": sum(1 for i in issues if i.severity.value == "warning"),
        "infos": sum(1 for i in issues if i.severity.value == "info"),
    }


def create_default_registry() -> RuleRegistry:
    registry = RuleRegistry()

    # 文档对象侧规则（复用旧版完整检查）
    registry.register(FirstStageSectionPresenceRule())
    registry.register(CoverTitleLengthRule())
    registry.register(ChineseAbstractLengthRule())
    registry.register(KeywordCountRule())
    registry.register(HeadingPunctuationRule())
    registry.register(HeadingCitationRule())
    registry.register(ReferenceTerminalPeriodRule())
    registry.register(ReferenceCountRule())
    registry.register(CitationSuperscriptRule())
    registry.register(ReferenceEntryFormatRule())
    registry.register(ReferenceNumberSequenceRule())
    registry.register(PageSettingsRule())
    registry.register(FontSizeFormatRule())
    registry.register(AlignmentFormatRule())
    registry.register(LineSpacingFormatRule())
    registry.register(CharacterSpacingFormatRule())
    registry.register(FormulaAlignmentRule())
    registry.register(FormulaNumberRightAlignedRule())
    registry.register(FormulaNumberFormatRule())
    registry.register(HeaderFormatRule())
    registry.register(CaptionFormatRule())
    registry.register(AppendixFormatRule())
    registry.register(HeadingNumberHierarchyRule())
    registry.register(NoteMarkerSuperscriptRule())
    registry.register(NoteMarkerConsistencyRule())
    registry.register(FooterFormatRule())
    registry.register(PageNumberFormatRule())
    registry.register(ReferenceAuthorCountRule())
    registry.register(SectionOrderRule())
    registry.register(CatalogueHeadingConsistencyRule())
    registry.register(CatalogueNumberFontRule())
    registry.register(CitationReferenceConsistencyRule())
    # 版面文本侧规则
    registry.register(CoverTitleCenterPdfRule())
    registry.register(TocPresencePdfRule())
    registry.register(TocLevelPresentationPdfRule())
    registry.register(PageNumberPresencePdfRule())
    registry.register(PageNumberBottomCenterPdfRule())
    registry.register(PageNumberStyleSequencePdfRule())
    registry.register(HeaderStartBoundaryPdfRule())
    registry.register(HeaderTopContentPdfRule())
    registry.register(ChapterStartsNewPagePdfRule())
    registry.register(FormulaNumberRightAlignedPdfRule())
    registry.register(FigureTableCaptionHintRule())
    registry.register(FigureCaptionBelowPdfRule())
    registry.register(FigureCaptionCenterPdfRule())
    registry.register(TableCaptionAbovePdfRule())
    registry.register(TableCaptionCenterPdfRule())
    registry.register(FigureTableSplitAcrossPagesPdfRule())

    # 图像识别侧规则（占位）
    registry.register(OcrFallbackAvailabilityRule())
    return registry


def build_context_status(extras: dict[str, Any]) -> dict[str, Any]:
    # 仅返回轻量状态，避免把大对象直接透出。
    return {
        "docx_parse_error": extras.get("docx_parse_error"),
        "docx_parts_order_count": len(extras.get("docx_parts_order") or []),
        "pdf_path": extras.get("pdf_path"),
        "pdf_extract_error": extras.get("pdf_extract_error"),
        "ocr_error": extras.get("ocr_error"),
    }


def create_default_hybrid_processor() -> HybridProcessor:
    registry = create_default_registry()
    return HybridProcessor(
        registry=registry,
        router=RuleRouter(registry),
        merger=IssueMerger(),
        fix_planner=FixPlanner(),
    )


STATUS_ONLY_RULE_IDS = {
    "ocr.fallback.status",
}
