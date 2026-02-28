"""スケジュール管理サービス - 従業員の稼働管理と可用性チェック"""

from __future__ import annotations

from datetime import date, timedelta

from workforce_allocation.models.employee import Employee
from workforce_allocation.models.schedule import Schedule
from workforce_allocation.models.task import Task, TaskStatus


class ScheduleManager:
    """スケジュール管理サービス

    従業員の稼働スケジュールを一元管理し、
    可用性チェック・稼働率計算・休日管理を行う。
    """

    def __init__(self) -> None:
        self._schedules: dict[str, Schedule] = {}

    @property
    def schedules(self) -> dict[str, Schedule]:
        return self._schedules

    def get_or_create_schedule(self, employee_id: str) -> Schedule:
        """従業員のスケジュールを取得（なければ作成）する"""
        if employee_id not in self._schedules:
            self._schedules[employee_id] = Schedule(employee_id)
        return self._schedules[employee_id]

    def check_availability(
        self,
        employee_id: str,
        target_date: date,
        required_hours: float,
    ) -> bool:
        """指定日に必要時間の空きがあるか確認する"""
        schedule = self.get_or_create_schedule(employee_id)
        slot = schedule.get_slot(target_date)
        return slot.can_allocate(required_hours)

    def find_available_employees(
        self,
        employees: list[Employee],
        target_date: date,
        required_hours: float,
    ) -> list[Employee]:
        """指定日に空きのある従業員一覧を取得する"""
        available = []
        for emp in employees:
            if not emp.is_active:
                continue
            if self.check_availability(emp.employee_id, target_date, required_hours):
                available.append(emp)
        return available

    def find_earliest_available_date(
        self,
        employee_id: str,
        start_date: date,
        required_hours: float,
        max_days_ahead: int = 30,
    ) -> date | None:
        """最短で利用可能な日付を探す"""
        schedule = self.get_or_create_schedule(employee_id)
        available_dates = schedule.get_available_dates(
            start_date,
            start_date + timedelta(days=max_days_ahead),
            required_hours,
        )
        return available_dates[0] if available_dates else None

    def assign_task(
        self,
        employee_id: str,
        task: Task,
        target_date: date,
    ) -> bool:
        """タスクをスケジュールに割り当てる"""
        schedule = self.get_or_create_schedule(employee_id)
        success = schedule.allocate_task(target_date, task.task_id, task.estimated_hours)
        if success:
            task.assign(employee_id, target_date)
        return success

    def unassign_task(
        self,
        employee_id: str,
        task: Task,
        target_date: date,
    ) -> None:
        """タスクのスケジュール割り当てを解除する"""
        schedule = self.get_or_create_schedule(employee_id)
        schedule.deallocate_task(target_date, task.task_id, task.estimated_hours)
        task.status = TaskStatus.UNASSIGNED
        task.assigned_employee_id = None

    def set_employee_day_off(
        self,
        employee_id: str,
        target_date: date,
        note: str | None = None,
    ) -> None:
        """従業員の休日を設定する"""
        schedule = self.get_or_create_schedule(employee_id)
        schedule.set_day_off(target_date, note)

    def set_employee_days_off(
        self,
        employee_id: str,
        start_date: date,
        end_date: date,
        note: str | None = None,
    ) -> None:
        """従業員の連続休暇を設定する"""
        current = start_date
        while current <= end_date:
            self.set_employee_day_off(employee_id, current, note)
            current += timedelta(days=1)

    def get_employee_utilization(
        self,
        employee_id: str,
        start_date: date,
        end_date: date,
    ) -> float:
        """指定期間の従業員稼働率を計算する"""
        schedule = self.get_or_create_schedule(employee_id)
        total_available = 0.0
        total_allocated = 0.0
        current = start_date
        while current <= end_date:
            slot = schedule.get_slot(current)
            if not slot.is_day_off:
                total_available += slot.available_hours
                total_allocated += slot.allocated_hours
            current += timedelta(days=1)
        if total_available == 0:
            return 0.0
        return total_allocated / total_available

    def get_team_utilization(
        self,
        employee_ids: list[str],
        start_date: date,
        end_date: date,
    ) -> dict[str, float]:
        """チーム全体の稼働率を取得する"""
        return {
            eid: self.get_employee_utilization(eid, start_date, end_date)
            for eid in employee_ids
        }

    def get_daily_summary(
        self,
        employee_ids: list[str],
        target_date: date,
    ) -> dict[str, dict]:
        """指定日の全従業員のスケジュールサマリーを取得する"""
        summary = {}
        for eid in employee_ids:
            schedule = self.get_or_create_schedule(eid)
            slot = schedule.get_slot(target_date)
            summary[eid] = {
                "available_hours": slot.available_hours,
                "allocated_hours": slot.allocated_hours,
                "remaining_hours": slot.remaining_hours,
                "utilization_rate": slot.utilization_rate,
                "task_count": len(slot.task_ids),
                "is_day_off": slot.is_day_off,
            }
        return summary
