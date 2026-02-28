"""配置エンジン - スキルマッチング・地域最適化に基づく要員配置"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Optional

from workforce_allocation.models.employee import Employee, SkillLevel
from workforce_allocation.models.region import RegionMap
from workforce_allocation.models.schedule import Schedule
from workforce_allocation.models.task import Task, TaskStatus


@dataclass
class AllocationCandidate:
    """配置候補 - 各候補者のスコアリング結果を保持する"""

    employee: Employee
    task: Task
    skill_score: float  # スキルマッチスコア (0.0〜1.0)
    region_score: float  # 地域適合スコア (0.0〜1.0)
    availability_score: float  # 空き状況スコア (0.0〜1.0)
    total_score: float = 0.0
    travel_time_minutes: float = 0.0
    scheduled_date: Optional[date] = None

    def __post_init__(self) -> None:
        if self.total_score == 0.0:
            self.total_score = self.calculate_total_score()

    def calculate_total_score(
        self,
        skill_weight: float = 0.4,
        region_weight: float = 0.3,
        availability_weight: float = 0.3,
    ) -> float:
        """加重スコアを計算する"""
        return (
            self.skill_score * skill_weight
            + self.region_score * region_weight
            + self.availability_score * availability_weight
        )


@dataclass
class AllocationResult:
    """配置結果"""

    task: Task
    assigned_employee: Optional[Employee]
    candidate: Optional[AllocationCandidate]
    success: bool
    reason: str = ""


class AllocationEngine:
    """要員配置エンジン

    スキルマッチング、地域最適化、稼働率バランスを総合的に評価して
    最適な要員配置を決定する。
    """

    def __init__(
        self,
        employees: list[Employee],
        region_map: RegionMap,
        schedules: dict[str, Schedule],
        skill_weight: float = 0.4,
        region_weight: float = 0.3,
        availability_weight: float = 0.3,
    ) -> None:
        self.employees = employees
        self.region_map = region_map
        self.schedules = schedules
        self.skill_weight = skill_weight
        self.region_weight = region_weight
        self.availability_weight = availability_weight

    def _get_active_employees(self) -> list[Employee]:
        """アクティブな従業員一覧を取得する"""
        return [e for e in self.employees if e.is_active]

    def _calculate_skill_score(self, employee: Employee, task: Task) -> float:
        """スキルマッチスコアを計算する"""
        return employee.skill_match_score(task.required_skills)

    def _calculate_region_score(self, employee: Employee, task: Task) -> float:
        """地域適合スコアを計算する。

        同一地域：1.0、隣接地域：0.7、その他は移動時間に基づいて減衰。
        """
        if employee.home_region_id == task.region_id:
            return 1.0

        region = self.region_map.get_region(employee.home_region_id)
        if region and region.is_adjacent(task.region_id):
            return 0.7

        travel_time = self.region_map.get_travel_time(
            employee.home_region_id, task.region_id
        )

        if not employee.travel_capable:
            if travel_time > 60:
                return 0.0

        # 移動時間に基づくスコア減衰（2時間以上で大幅減少）
        if travel_time <= 30:
            return 0.9
        elif travel_time <= 60:
            return 0.7
        elif travel_time <= 120:
            return 0.4
        else:
            return max(0.1, 1.0 - travel_time / 300.0)

    def _calculate_availability_score(
        self,
        employee: Employee,
        task: Task,
        target_date: date,
    ) -> float:
        """空き状況スコアを計算する"""
        schedule = self.schedules.get(employee.employee_id)
        if schedule is None:
            return 1.0

        slot = schedule.get_slot(target_date)
        if slot.is_day_off:
            return 0.0
        if not slot.can_allocate(task.estimated_hours):
            return 0.0

        # 残り時間が多いほど高スコア
        remaining_ratio = slot.remaining_hours / slot.available_hours
        return remaining_ratio

    def evaluate_candidate(
        self,
        employee: Employee,
        task: Task,
        target_date: date,
    ) -> Optional[AllocationCandidate]:
        """候補者を評価し、スコアリングする"""
        # スキル要件チェック
        skill_score = self._calculate_skill_score(employee, task)
        if skill_score == 0.0 and task.required_skills:
            return None

        region_score = self._calculate_region_score(employee, task)
        if region_score == 0.0:
            return None

        availability_score = self._calculate_availability_score(
            employee, task, target_date
        )
        if availability_score == 0.0:
            return None

        travel_time = self.region_map.get_travel_time(
            employee.home_region_id, task.region_id
        )

        candidate = AllocationCandidate(
            employee=employee,
            task=task,
            skill_score=skill_score,
            region_score=region_score,
            availability_score=availability_score,
            travel_time_minutes=travel_time,
            scheduled_date=target_date,
        )
        candidate.total_score = candidate.calculate_total_score(
            self.skill_weight, self.region_weight, self.availability_weight
        )
        return candidate

    def find_candidates(
        self,
        task: Task,
        target_date: date,
        min_score: float = 0.3,
    ) -> list[AllocationCandidate]:
        """タスクに対する候補者一覧をスコア順で取得する"""
        candidates = []
        for employee in self._get_active_employees():
            candidate = self.evaluate_candidate(employee, task, target_date)
            if candidate and candidate.total_score >= min_score:
                candidates.append(candidate)

        candidates.sort(key=lambda c: c.total_score, reverse=True)
        return candidates

    def allocate_task(
        self,
        task: Task,
        target_date: date,
        min_score: float = 0.3,
    ) -> AllocationResult:
        """タスクに最適な要員を割り当てる"""
        if task.status not in (TaskStatus.UNASSIGNED,):
            return AllocationResult(
                task=task,
                assigned_employee=None,
                candidate=None,
                success=False,
                reason=f"タスクのステータスが不正です: {task.status.name}",
            )

        candidates = self.find_candidates(task, target_date, min_score)

        if not candidates:
            return AllocationResult(
                task=task,
                assigned_employee=None,
                candidate=None,
                success=False,
                reason="要件を満たす候補者が見つかりませんでした",
            )

        best = candidates[0]

        # スケジュールに割り当て
        schedule = self.schedules.get(best.employee.employee_id)
        if schedule is None:
            schedule = Schedule(best.employee.employee_id)
            self.schedules[best.employee.employee_id] = schedule

        allocated = schedule.allocate_task(
            target_date, task.task_id, task.estimated_hours
        )
        if not allocated:
            return AllocationResult(
                task=task,
                assigned_employee=None,
                candidate=best,
                success=False,
                reason="スケジュールへの割り当てに失敗しました",
            )

        task.assign(best.employee.employee_id, target_date)

        return AllocationResult(
            task=task,
            assigned_employee=best.employee,
            candidate=best,
            success=True,
            reason="配置完了",
        )

    def allocate_batch(
        self,
        tasks: list[Task],
        target_date: date,
        min_score: float = 0.3,
    ) -> list[AllocationResult]:
        """複数タスクを一括配置する（優先度・緊急度順）"""
        # 緊急度の高いタスクから配置
        sorted_tasks = sorted(tasks, key=lambda t: t.urgency_score, reverse=True)
        results = []
        for task in sorted_tasks:
            result = self.allocate_task(task, target_date, min_score)
            results.append(result)
        return results

    def suggest_reallocation(
        self,
        task: Task,
        target_date: date,
    ) -> list[AllocationCandidate]:
        """再配置の候補を提案する（現在の割り当てを考慮しない）"""
        candidates = []
        for employee in self._get_active_employees():
            skill_score = self._calculate_skill_score(employee, task)
            region_score = self._calculate_region_score(employee, task)

            if skill_score == 0.0 and task.required_skills:
                continue

            travel_time = self.region_map.get_travel_time(
                employee.home_region_id, task.region_id
            )

            candidate = AllocationCandidate(
                employee=employee,
                task=task,
                skill_score=skill_score,
                region_score=region_score,
                availability_score=0.5,  # 再配置のため固定値
                travel_time_minutes=travel_time,
                scheduled_date=target_date,
            )
            candidate.total_score = candidate.calculate_total_score(
                self.skill_weight, self.region_weight, self.availability_weight
            )
            candidates.append(candidate)

        candidates.sort(key=lambda c: c.total_score, reverse=True)
        return candidates
