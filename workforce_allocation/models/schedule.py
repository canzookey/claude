"""スケジュールモデル - 従業員の稼働スケジュール管理"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Optional


@dataclass
class TimeSlot:
    """タイムスロット（1日単位の稼働枠）"""

    slot_date: date
    employee_id: str
    available_hours: float = 8.0
    allocated_hours: float = 0.0
    task_ids: list[str] = field(default_factory=list)
    is_day_off: bool = False
    note: Optional[str] = None

    @property
    def remaining_hours(self) -> float:
        """残りの利用可能時間"""
        if self.is_day_off:
            return 0.0
        return max(0.0, self.available_hours - self.allocated_hours)

    @property
    def utilization_rate(self) -> float:
        """稼働率（0.0〜1.0）"""
        if self.available_hours <= 0 or self.is_day_off:
            return 0.0
        return min(self.allocated_hours / self.available_hours, 1.0)

    def can_allocate(self, hours: float) -> bool:
        """指定時間のタスクを割り当て可能か判定する"""
        return not self.is_day_off and self.remaining_hours >= hours

    def allocate(self, task_id: str, hours: float) -> bool:
        """タスクを割り当てる。成功したらTrueを返す。"""
        if not self.can_allocate(hours):
            return False
        self.allocated_hours += hours
        self.task_ids.append(task_id)
        return True

    def deallocate(self, task_id: str, hours: float) -> None:
        """タスクの割り当てを解除する"""
        if task_id in self.task_ids:
            self.task_ids.remove(task_id)
            self.allocated_hours = max(0.0, self.allocated_hours - hours)


class Schedule:
    """従業員のスケジュール管理"""

    def __init__(self, employee_id: str) -> None:
        self.employee_id = employee_id
        self._slots: dict[date, TimeSlot] = {}

    def get_slot(self, target_date: date) -> TimeSlot:
        """指定日のタイムスロットを取得する（なければ作成）"""
        if target_date not in self._slots:
            is_weekend = target_date.weekday() >= 5
            self._slots[target_date] = TimeSlot(
                slot_date=target_date,
                employee_id=self.employee_id,
                is_day_off=is_weekend,
            )
        return self._slots[target_date]

    def get_available_dates(
        self,
        start_date: date,
        end_date: date,
        min_hours: float = 1.0,
    ) -> list[date]:
        """指定期間内で空き時間のある日を取得する"""
        available = []
        current = start_date
        while current <= end_date:
            slot = self.get_slot(current)
            if slot.remaining_hours >= min_hours:
                available.append(current)
            current += timedelta(days=1)
        return available

    def get_weekly_allocated_hours(self, week_start: date) -> float:
        """指定週の合計割り当て時間を取得する"""
        total = 0.0
        for i in range(7):
            d = week_start + timedelta(days=i)
            slot = self.get_slot(d)
            total += slot.allocated_hours
        return total

    def get_weekly_utilization(self, week_start: date) -> float:
        """指定週の平均稼働率を取得する"""
        total_available = 0.0
        total_allocated = 0.0
        for i in range(7):
            d = week_start + timedelta(days=i)
            slot = self.get_slot(d)
            if not slot.is_day_off:
                total_available += slot.available_hours
                total_allocated += slot.allocated_hours
        if total_available == 0:
            return 0.0
        return total_allocated / total_available

    def set_day_off(self, target_date: date, note: Optional[str] = None) -> None:
        """休日を設定する"""
        slot = self.get_slot(target_date)
        slot.is_day_off = True
        slot.note = note

    def allocate_task(self, target_date: date, task_id: str, hours: float) -> bool:
        """指定日にタスクを割り当てる"""
        slot = self.get_slot(target_date)
        return slot.allocate(task_id, hours)

    def deallocate_task(self, target_date: date, task_id: str, hours: float) -> None:
        """タスク割り当てを解除する"""
        slot = self.get_slot(target_date)
        slot.deallocate(task_id, hours)
