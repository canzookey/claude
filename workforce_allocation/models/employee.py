"""従業員モデル - スキル・資格・所属地域を管理する"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import IntEnum
from typing import Optional


class SkillLevel(IntEnum):
    """スキル習熟度レベル"""

    BEGINNER = 1  # 初級：基本作業が可能
    INTERMEDIATE = 2  # 中級：独力で標準作業が可能
    ADVANCED = 3  # 上級：複雑な案件に対応可能
    EXPERT = 4  # エキスパート：指導・設計が可能


@dataclass
class Skill:
    """スキル情報"""

    name: str
    level: SkillLevel
    certified: bool = False  # 資格保有フラグ
    years_experience: float = 0.0

    @property
    def weighted_score(self) -> float:
        """スキルの重み付きスコアを計算する。

        資格保有と経験年数をスキルレベルに加味して総合スコアを算出する。
        """
        base = float(self.level)
        cert_bonus = 1.5 if self.certified else 1.0
        exp_bonus = min(self.years_experience / 10.0, 0.5)  # 最大0.5ポイント加算
        return base * cert_bonus + exp_bonus


@dataclass
class Employee:
    """従業員"""

    employee_id: str
    name: str
    department: str
    home_region_id: str  # 所属地域
    skills: list[Skill] = field(default_factory=list)
    max_weekly_hours: float = 40.0
    is_active: bool = True
    role: str = "field_engineer"  # field_engineer, senior_engineer, manager
    travel_capable: bool = True  # 出張可能フラグ
    certifications: list[str] = field(default_factory=list)
    notes: Optional[str] = None

    def has_skill(self, skill_name: str, min_level: SkillLevel = SkillLevel.BEGINNER) -> bool:
        """指定スキルを必要レベル以上で保有しているか判定する"""
        for s in self.skills:
            if s.name == skill_name and s.level >= min_level:
                return True
        return False

    def get_skill(self, skill_name: str) -> Optional[Skill]:
        """指定スキルを取得する"""
        for s in self.skills:
            if s.name == skill_name:
                return s
        return None

    def skill_match_score(self, required_skills: dict[str, SkillLevel]) -> float:
        """要求スキルセットとのマッチ度を計算する。

        Args:
            required_skills: スキル名→必要レベルの辞書

        Returns:
            0.0〜1.0のマッチスコア。1.0は全スキル要件を満たしていることを示す。
            要件を超えるスキルレベルはボーナスとして加算される。
        """
        if not required_skills:
            return 1.0

        total_score = 0.0
        max_possible = 0.0

        for skill_name, required_level in required_skills.items():
            max_possible += float(required_level)
            emp_skill = self.get_skill(skill_name)
            if emp_skill is None:
                continue
            actual = min(float(emp_skill.level), float(required_level) * 1.2)
            total_score += actual

        if max_possible == 0:
            return 1.0
        return min(total_score / max_possible, 1.0)

    def meets_requirements(self, required_skills: dict[str, SkillLevel]) -> bool:
        """全ての要求スキルを満たしているか判定する"""
        for skill_name, required_level in required_skills.items():
            if not self.has_skill(skill_name, required_level):
                return False
        return True
