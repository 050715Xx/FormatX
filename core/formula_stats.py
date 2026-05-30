"""公式身份追踪与全局统计模型"""
from dataclasses import dataclass, field


@dataclass
class FormulaOccurrence:
    """公式身份证：完整记录一个公式的身世、坐标与状态"""
    formula_id: str
    original_text: str
    repaired_text: str = ""
    source_type: str = "plain_text"
    paragraph_index: int = -1
    confidence: float = 1.0
    is_fixed: bool = False
    warnings: list[str] = field(default_factory=list)


@dataclass
class FormulaRuleStats:
    """全局公式运行统计看板"""
    matched: int = 0
    converted: int = 0
    repaired: int = 0
    high_confidence: int = 0
    medium_confidence: int = 0
    low_confidence: int = 0
    occurrences: list[FormulaOccurrence] = field(default_factory=list)

    def record_occurrence(self, occ: FormulaOccurrence):
        self.matched += 1
        self.occurrences.append(occ)
        if occ.confidence >= 0.85:
            self.high_confidence += 1
        elif occ.confidence >= 0.60:
            self.medium_confidence += 1
        else:
            self.low_confidence += 1
        if occ.is_fixed:
            self.repaired += 1
