from model.core.issue import Issue


class IssueMerger:
    SOURCE_PRIORITY = {
        "docx": 3,
        "pdf": 2,
        "ocr": 1,
        "hybrid": 0,
    }

    def merge(self, groups: dict[str, list[Issue]]) -> list[Issue]:
        merged: dict[tuple[str, int | None, str], Issue] = {}

        for _, issues in groups.items():
            for issue in issues:
                key = (issue.rule_id, issue.page, issue.title)
                if key not in merged:
                    merged[key] = issue
                    continue

                current = merged[key]
                old_rank = self.SOURCE_PRIORITY.get(current.source.value, 0)
                new_rank = self.SOURCE_PRIORITY.get(issue.source.value, 0)

                if new_rank > old_rank or issue.confidence > current.confidence:
                    merged[key] = issue

        return list(merged.values())

