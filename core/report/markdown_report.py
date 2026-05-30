"""Markdown 排版审计报告生成器"""
from core.report.collector import ReportData


def _escape_md_cell(value, *, limit: int | None = None) -> str:
    text = "" if value is None else str(value)
    text = text.replace("|", "\\|").replace("\n", "<br>")
    if limit is not None:
        text = text[:limit]
    return text


def generate_markdown_report(report: ReportData, output_path: str) -> None:
    lines = ["# FormatX 排版审计报告\n"]
    lines.append(f"- **输入文件**: `{report.input_file}`")
    lines.append(f"- **总修改数**: {report.total_changes} 项")
    if report.total_failures > 0:
        lines.append(f"- **潜在问题**: {report.total_failures} 项")
    if report.formula_matched > 0:
        lines.append(f"- **公式解析**: {report.formula_matched} 个")
    if report.formula_repaired > 0:
        lines.append(f"- **ICU 抢救**: {report.formula_repaired} 个")

    if report.low_confidence_items:
        lines.append("\n## 纸上风骨 · 气韵微瑕\n")
        lines.append("| 编号 | 卷段 | 身世 | 气韵 | 执念 | 新生 |")
        lines.append("|---|---|---|---|---|---|")
        for item in report.low_confidence_items[:20]:
            src = {"plain_text":"网页复制","ole_equation":"岁月遗痕","ocr_fragment":"光影拓印"}.get(item['source'], item['source'])
            lines.append(
                f"| {item['id']} | 卷之{item['index']}段 | {src} "
                f"| {item['confidence']*100:.0f}% "
                f"| `{_escape_md_cell(item['text'], limit=28)}` "
                f"| `{_escape_md_cell(item['repaired'] or '—', limit=28)}` |")
        lines.append("")

    if report.formula_diagnostics:
        lines.append("\n## 炉火重铸 · ICU 抢救纪\n")
        for item in report.formula_diagnostics[:20]:
            lines.append(f"- **卷之{item['index']}段**: 旧影 `{_escape_md_cell(item['original'], limit=30)}` → 新生 `{_escape_md_cell(item['repaired'], limit=30)}`（气韵 {item['confidence']*100:.0f}%）")
        lines.append("")

    if report.validation_issues:
        lines.append("\n## 岁时无惊 · 微澜偶起\n")
        for issue in report.validation_issues[:30]:
            lines.append(f"- **{issue['rule']}**（{issue['target']}）: {issue['reason']}")
        lines.append("")

    lines.append("\n## 细微之处 · 修改详情\n")

    for rule_name, changes in report.changes_by_rule.items():
        lines.append(f"### {rule_name} ({len(changes)} 次)\n")
        lines.append("| 位置 | 类型 | 修改前 | 修改后 | 状态 |")
        lines.append("|---|---|---|---|---|")
        for c in changes:
            status = "成功" if c["success"] else f"警告: {c['failure'] or ''}"
            lines.append(
                f"| {_escape_md_cell(c['target'])} "
                f"| {_escape_md_cell(c['type'])} "
                f"| `{_escape_md_cell(c['before'], limit=30)}` "
                f"| `{_escape_md_cell(c['after'], limit=30)}` "
                f"| {status} |")
        lines.append("")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
