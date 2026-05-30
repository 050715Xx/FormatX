"""报告收集器——汇总排版变更与公式统计的结构化数据"""
from dataclasses import dataclass, field
from analyzer.change_tracker import ChangeTracker


@dataclass
class ReportData:
    scene_name: str = ""
    input_file: str = ""
    total_changes: int = 0
    total_failures: int = 0
    changes_by_rule: dict = field(default_factory=dict)
    summary: dict = field(default_factory=dict)
    formula_matched: int = 0
    formula_repaired: int = 0
    low_confidence_items: list[dict] = field(default_factory=list)
    formula_diagnostics: list[dict] = field(default_factory=list)
    validation_issues: list[dict] = field(default_factory=list)


def collect_report(tracker: ChangeTracker,
                   formula_stats=None,
                   scene_name: str = "FormatX",
                   input_file: str = "") -> ReportData:
    report = ReportData(scene_name=scene_name, input_file=input_file)
    rules_seen = {}
    for rec in tracker.records:
        rule = rec.rule_name
        if rule not in rules_seen:
            rules_seen[rule] = []
        rules_seen[rule].append({
            "target": rec.target,
            "type": rec.change_type,
            "before": rec.before,
            "after": rec.after,
            "success": rec.success,
            "failure": rec.failure_reason,
        })
    report.changes_by_rule = rules_seen
    report.total_changes = len(tracker.records)
    report.total_failures = len([r for r in tracker.records if not r.success])
    report.summary = tracker.summary() if hasattr(tracker, 'summary') else {}

    if formula_stats:
        report.formula_matched = getattr(formula_stats, 'matched', 0)
        report.formula_repaired = getattr(formula_stats, 'repaired', 0)
        occurrences = getattr(formula_stats, 'occurrences', [])
        for occ in occurrences:
            if occ.confidence < 0.85:
                report.low_confidence_items.append({
                    "id": occ.formula_id,
                    "text": occ.original_text,
                    "repaired": occ.repaired_text if occ.is_fixed else "",
                    "confidence": occ.confidence,
                    "source": occ.source_type,
                    "index": occ.paragraph_index,
                    "warnings": occ.warnings,
                })
            if occ.is_fixed:
                report.formula_diagnostics.append({
                    "id": occ.formula_id,
                    "original": occ.original_text,
                    "repaired": occ.repaired_text,
                    "confidence": occ.confidence,
                    "index": occ.paragraph_index,
                })

    for rec in tracker.records:
        if not rec.success and rec.failure_reason:
            report.validation_issues.append({
                "target": rec.target,
                "rule": rec.rule_name,
                "reason": rec.failure_reason,
                "before": rec.before,
                "index": rec.paragraph_index,
            })

    return report
