"""レポーティングサービス - 配置状況・稼働率の分析レポート生成"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

from workforce_allocation.models.employee import Employee
from workforce_allocation.models.task import Task, TaskPriority, TaskStatus
from workforce_allocation.services.schedule_manager import ScheduleManager


@dataclass
class UtilizationReport:
    """稼働率レポート"""

    period_start: date
    period_end: date
    employee_utilizations: dict[str, float] = field(default_factory=dict)
    team_average: float = 0.0
    overloaded_employees: list[str] = field(default_factory=list)  # 稼働率80%超
    underutilized_employees: list[str] = field(default_factory=list)  # 稼働率30%未満


@dataclass
class AllocationSummary:
    """配置サマリー"""

    total_tasks: int = 0
    assigned_tasks: int = 0
    unassigned_tasks: int = 0
    completed_tasks: int = 0
    overdue_tasks: int = 0
    assignment_rate: float = 0.0
    tasks_by_priority: dict[str, int] = field(default_factory=dict)
    tasks_by_region: dict[str, int] = field(default_factory=dict)
    tasks_by_type: dict[str, int] = field(default_factory=dict)


@dataclass
class RegionWorkloadReport:
    """地域別ワークロードレポート"""

    region_id: str
    region_name: str
    total_tasks: int = 0
    assigned_tasks: int = 0
    total_hours: float = 0.0
    assigned_employees: int = 0


class ReportingService:
    """レポーティングサービス

    配置状況と稼働率に関する各種レポートを生成する。
    """

    def __init__(
        self,
        employees: list[Employee],
        tasks: list[Task],
        schedule_manager: ScheduleManager,
    ) -> None:
        self.employees = employees
        self.tasks = tasks
        self.schedule_manager = schedule_manager

    def generate_utilization_report(
        self,
        start_date: date,
        end_date: date,
        overload_threshold: float = 0.8,
        underutil_threshold: float = 0.3,
    ) -> UtilizationReport:
        """稼働率レポートを生成する"""
        report = UtilizationReport(period_start=start_date, period_end=end_date)

        active_employees = [e for e in self.employees if e.is_active]
        total_util = 0.0

        for emp in active_employees:
            util = self.schedule_manager.get_employee_utilization(
                emp.employee_id, start_date, end_date
            )
            report.employee_utilizations[emp.employee_id] = util
            total_util += util

            if util > overload_threshold:
                report.overloaded_employees.append(emp.employee_id)
            elif util < underutil_threshold:
                report.underutilized_employees.append(emp.employee_id)

        if active_employees:
            report.team_average = total_util / len(active_employees)

        return report

    def generate_allocation_summary(self) -> AllocationSummary:
        """配置サマリーを生成する"""
        summary = AllocationSummary()
        summary.total_tasks = len(self.tasks)

        for task in self.tasks:
            # ステータス別カウント
            if task.status == TaskStatus.UNASSIGNED:
                summary.unassigned_tasks += 1
            elif task.status == TaskStatus.COMPLETED:
                summary.completed_tasks += 1
            else:
                summary.assigned_tasks += 1

            if task.is_overdue:
                summary.overdue_tasks += 1

            # 優先度別
            priority_name = task.priority.name
            summary.tasks_by_priority[priority_name] = (
                summary.tasks_by_priority.get(priority_name, 0) + 1
            )

            # 地域別
            summary.tasks_by_region[task.region_id] = (
                summary.tasks_by_region.get(task.region_id, 0) + 1
            )

            # タイプ別
            summary.tasks_by_type[task.task_type] = (
                summary.tasks_by_type.get(task.task_type, 0) + 1
            )

        active_count = summary.assigned_tasks + summary.completed_tasks
        if summary.total_tasks > 0:
            summary.assignment_rate = active_count / summary.total_tasks

        return summary

    def generate_region_workload(self) -> list[RegionWorkloadReport]:
        """地域別ワークロードレポートを生成する"""
        region_data: dict[str, RegionWorkloadReport] = {}

        for task in self.tasks:
            if task.status == TaskStatus.CANCELLED:
                continue

            if task.region_id not in region_data:
                region_data[task.region_id] = RegionWorkloadReport(
                    region_id=task.region_id,
                    region_name=task.region_id,  # 地域名は別途設定
                )

            report = region_data[task.region_id]
            report.total_tasks += 1
            report.total_hours += task.estimated_hours
            if task.status != TaskStatus.UNASSIGNED:
                report.assigned_tasks += 1

        # 各地域の配置済み従業員数を集計
        for emp in self.employees:
            if emp.is_active and emp.home_region_id in region_data:
                region_data[emp.home_region_id].assigned_employees += 1

        return sorted(
            region_data.values(),
            key=lambda r: r.total_tasks,
            reverse=True,
        )

    def generate_employee_workload_summary(
        self,
        target_date: date,
    ) -> list[dict]:
        """従業員別ワークロードサマリーを生成する"""
        results = []
        active_employees = [e for e in self.employees if e.is_active]

        for emp in active_employees:
            emp_tasks = [
                t
                for t in self.tasks
                if t.assigned_employee_id == emp.employee_id
                and t.status not in (TaskStatus.COMPLETED, TaskStatus.CANCELLED)
            ]

            week_start = target_date - timedelta(days=target_date.weekday())
            weekly_util = self.schedule_manager.get_employee_utilization(
                emp.employee_id, week_start, week_start + timedelta(days=6)
            )

            results.append(
                {
                    "employee_id": emp.employee_id,
                    "name": emp.name,
                    "department": emp.department,
                    "region": emp.home_region_id,
                    "active_tasks": len(emp_tasks),
                    "weekly_utilization": round(weekly_util * 100, 1),
                    "skills_count": len(emp.skills),
                    "role": emp.role,
                }
            )

        results.sort(key=lambda x: x["weekly_utilization"], reverse=True)
        return results

    def identify_bottlenecks(self) -> dict:
        """ボトルネックを特定する"""
        unassigned_by_skill: dict[str, int] = {}
        unassigned_by_region: dict[str, int] = {}
        high_priority_unassigned = 0

        for task in self.tasks:
            if task.status != TaskStatus.UNASSIGNED:
                continue

            if task.priority >= TaskPriority.HIGH:
                high_priority_unassigned += 1

            for skill_name in task.required_skills:
                unassigned_by_skill[skill_name] = (
                    unassigned_by_skill.get(skill_name, 0) + 1
                )

            unassigned_by_region[task.region_id] = (
                unassigned_by_region.get(task.region_id, 0) + 1
            )

        # スキル不足の特定
        skill_shortage = {}
        for skill_name, count in unassigned_by_skill.items():
            employees_with_skill = sum(
                1 for e in self.employees if e.is_active and e.has_skill(skill_name)
            )
            if employees_with_skill < count:
                skill_shortage[skill_name] = {
                    "demand": count,
                    "supply": employees_with_skill,
                    "gap": count - employees_with_skill,
                }

        return {
            "high_priority_unassigned": high_priority_unassigned,
            "skill_shortage": skill_shortage,
            "unassigned_by_region": unassigned_by_region,
            "total_unassigned": sum(
                1
                for t in self.tasks
                if t.status == TaskStatus.UNASSIGNED
            ),
        }
