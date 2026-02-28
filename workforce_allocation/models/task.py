"""タスクモデル - フィールドサービス業務を表現する"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from enum import IntEnum, auto
from typing import Optional

from workforce_allocation.models.employee import SkillLevel


class TaskPriority(IntEnum):
    """タスク優先度"""

    LOW = 1
    MEDIUM = 2
    HIGH = 3
    CRITICAL = 4


class TaskStatus(IntEnum):
    """タスクステータス"""

    UNASSIGNED = auto()
    ASSIGNED = auto()
    IN_PROGRESS = auto()
    COMPLETED = auto()
    CANCELLED = auto()


@dataclass
class Task:
    """フィールドサービスタスク"""

    task_id: str
    title: str
    description: str
    region_id: str  # 対象地域
    priority: TaskPriority = TaskPriority.MEDIUM
    status: TaskStatus = TaskStatus.UNASSIGNED
    required_skills: dict[str, SkillLevel] = field(default_factory=dict)
    estimated_hours: float = 4.0
    scheduled_date: Optional[date] = None
    deadline: Optional[date] = None
    assigned_employee_id: Optional[str] = None
    customer_name: str = ""
    customer_id: str = ""
    task_type: str = "maintenance"  # maintenance, installation, repair, consultation
    requires_onsite: bool = True
    created_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None

    def __post_init__(self) -> None:
        if self.created_at is None:
            self.created_at = datetime.now()

    @property
    def is_overdue(self) -> bool:
        """期限超過チェック"""
        if self.deadline is None:
            return False
        if self.status in (TaskStatus.COMPLETED, TaskStatus.CANCELLED):
            return False
        return date.today() > self.deadline

    @property
    def days_until_deadline(self) -> Optional[int]:
        """期限までの残日数"""
        if self.deadline is None:
            return None
        return (self.deadline - date.today()).days

    @property
    def urgency_score(self) -> float:
        """緊急度スコアを計算する。

        優先度と期限を考慮した緊急度を0〜10のスコアで返す。
        """
        base = float(self.priority) * 2.0

        if self.deadline is not None:
            days_left = self.days_until_deadline
            if days_left is not None:
                if days_left < 0:
                    base += 3.0  # 期限超過ペナルティ
                elif days_left <= 1:
                    base += 2.0
                elif days_left <= 3:
                    base += 1.0

        return min(base, 10.0)

    def assign(self, employee_id: str, scheduled_date: Optional[date] = None) -> None:
        """タスクに従業員を割り当てる"""
        self.assigned_employee_id = employee_id
        self.status = TaskStatus.ASSIGNED
        if scheduled_date is not None:
            self.scheduled_date = scheduled_date

    def complete(self) -> None:
        """タスクを完了にする"""
        self.status = TaskStatus.COMPLETED
        self.completed_at = datetime.now()

    def cancel(self) -> None:
        """タスクをキャンセルする"""
        self.status = TaskStatus.CANCELLED
