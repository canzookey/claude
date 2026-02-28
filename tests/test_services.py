"""サービス層のテスト"""

import pytest
from datetime import date, timedelta

from workforce_allocation.models.employee import Employee, Skill, SkillLevel
from workforce_allocation.models.region import Region, RegionMap
from workforce_allocation.models.task import Task, TaskPriority, TaskStatus
from workforce_allocation.models.schedule import Schedule
from workforce_allocation.services.allocation_engine import AllocationEngine
from workforce_allocation.services.schedule_manager import ScheduleManager
from workforce_allocation.services.reporting import ReportingService


def _build_test_data():
    """テスト用データを構築する"""
    # 地域マップ
    region_map = RegionMap()
    tokyo = Region(
        region_id="tokyo",
        name="東京",
        prefecture="東京都",
        adjacent_region_ids=["kanagawa", "saitama"],
    )
    kanagawa = Region(
        region_id="kanagawa",
        name="神奈川",
        prefecture="神奈川県",
        adjacent_region_ids=["tokyo"],
    )
    osaka = Region(
        region_id="osaka",
        name="大阪",
        prefecture="大阪府",
    )
    region_map.add_region(tokyo)
    region_map.add_region(kanagawa)
    region_map.add_region(osaka)
    region_map.set_distance("tokyo", "kanagawa", 30.0, 45.0)
    region_map.set_distance("tokyo", "osaka", 500.0, 150.0)
    region_map.set_distance("kanagawa", "osaka", 480.0, 140.0)

    # 従業員
    employees = [
        Employee(
            employee_id="EMP001",
            name="山田太郎",
            department="フィールドサービス",
            home_region_id="tokyo",
            skills=[
                Skill(name="サーバー", level=SkillLevel.ADVANCED, certified=True),
                Skill(name="ネットワーク", level=SkillLevel.INTERMEDIATE),
            ],
            role="senior_engineer",
        ),
        Employee(
            employee_id="EMP002",
            name="鈴木花子",
            department="フィールドサービス",
            home_region_id="kanagawa",
            skills=[
                Skill(name="ネットワーク", level=SkillLevel.EXPERT, certified=True),
                Skill(name="セキュリティ", level=SkillLevel.ADVANCED),
            ],
            role="field_engineer",
        ),
        Employee(
            employee_id="EMP003",
            name="佐藤一郎",
            department="フィールドサービス",
            home_region_id="osaka",
            skills=[
                Skill(name="サーバー", level=SkillLevel.INTERMEDIATE),
                Skill(name="データベース", level=SkillLevel.ADVANCED),
            ],
            role="field_engineer",
        ),
    ]

    return region_map, employees


class TestAllocationEngine:
    def _make_engine(self):
        region_map, employees = _build_test_data()
        schedules = {emp.employee_id: Schedule(emp.employee_id) for emp in employees}
        engine = AllocationEngine(employees, region_map, schedules)
        return engine, employees

    def test_find_candidates_skill_match(self):
        engine, employees = self._make_engine()
        task = Task(
            task_id="T001",
            title="サーバー保守",
            description="定期メンテナンス",
            region_id="tokyo",
            required_skills={"サーバー": SkillLevel.INTERMEDIATE},
            estimated_hours=4.0,
        )
        target = date(2026, 3, 2)  # 月曜日
        candidates = engine.find_candidates(task, target)
        assert len(candidates) > 0
        # 東京在住の山田がトップ候補になるべき
        assert candidates[0].employee.employee_id == "EMP001"

    def test_find_candidates_no_match(self):
        engine, _ = self._make_engine()
        task = Task(
            task_id="T002",
            title="特殊作業",
            description="特殊なスキルが必要",
            region_id="tokyo",
            required_skills={"量子コンピューティング": SkillLevel.EXPERT},
            estimated_hours=4.0,
        )
        target = date(2026, 3, 2)
        candidates = engine.find_candidates(task, target)
        assert len(candidates) == 0

    def test_allocate_task_success(self):
        engine, _ = self._make_engine()
        task = Task(
            task_id="T001",
            title="ネットワーク点検",
            description="定期点検",
            region_id="kanagawa",
            required_skills={"ネットワーク": SkillLevel.INTERMEDIATE},
            estimated_hours=4.0,
        )
        target = date(2026, 3, 2)
        result = engine.allocate_task(task, target)
        assert result.success
        assert result.assigned_employee is not None
        # 神奈川在住のネットワークエキスパートの鈴木がベスト
        assert result.assigned_employee.employee_id == "EMP002"
        assert task.status == TaskStatus.ASSIGNED

    def test_allocate_task_already_assigned(self):
        engine, _ = self._make_engine()
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
            status=TaskStatus.ASSIGNED,
        )
        target = date(2026, 3, 2)
        result = engine.allocate_task(task, target)
        assert not result.success

    def test_allocate_task_weekend(self):
        engine, _ = self._make_engine()
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
            estimated_hours=4.0,
        )
        target = date(2026, 3, 1)  # 日曜日
        result = engine.allocate_task(task, target)
        assert not result.success

    def test_allocate_batch(self):
        engine, _ = self._make_engine()
        tasks = [
            Task(
                task_id="T001",
                title="緊急修理",
                description="テスト",
                region_id="tokyo",
                priority=TaskPriority.CRITICAL,
                estimated_hours=3.0,
            ),
            Task(
                task_id="T002",
                title="定期点検",
                description="テスト",
                region_id="kanagawa",
                priority=TaskPriority.LOW,
                estimated_hours=2.0,
            ),
        ]
        target = date(2026, 3, 2)
        results = engine.allocate_batch(tasks, target)
        assert len(results) == 2
        # 緊急タスクが先に配置される
        assert results[0].task.task_id == "T001"

    def test_region_score_same_region(self):
        engine, employees = self._make_engine()
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
        )
        score = engine._calculate_region_score(employees[0], task)
        assert score == 1.0

    def test_region_score_adjacent_region(self):
        engine, employees = self._make_engine()
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="kanagawa",
        )
        score = engine._calculate_region_score(employees[0], task)  # 東京→神奈川
        assert score == 0.7

    def test_suggest_reallocation(self):
        engine, _ = self._make_engine()
        task = Task(
            task_id="T001",
            title="サーバー保守",
            description="テスト",
            region_id="tokyo",
            required_skills={"サーバー": SkillLevel.BEGINNER},
        )
        target = date(2026, 3, 2)
        suggestions = engine.suggest_reallocation(task, target)
        assert len(suggestions) > 0


class TestScheduleManager:
    def test_check_availability(self):
        sm = ScheduleManager()
        # 平日は空きあり
        assert sm.check_availability("EMP001", date(2026, 3, 2), 4.0)
        # 休日は空きなし
        assert not sm.check_availability("EMP001", date(2026, 3, 1), 4.0)

    def test_find_available_employees(self):
        _, employees = _build_test_data()
        sm = ScheduleManager()
        available = sm.find_available_employees(employees, date(2026, 3, 2), 4.0)
        assert len(available) == 3

    def test_assign_and_unassign_task(self):
        sm = ScheduleManager()
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
            estimated_hours=4.0,
        )
        target = date(2026, 3, 2)

        # 割り当て
        assert sm.assign_task("EMP001", task, target)
        assert task.status == TaskStatus.ASSIGNED
        assert not sm.check_availability("EMP001", target, 5.0)

        # 解除
        sm.unassign_task("EMP001", task, target)
        assert task.status == TaskStatus.UNASSIGNED
        assert sm.check_availability("EMP001", target, 8.0)

    def test_set_days_off(self):
        sm = ScheduleManager()
        start = date(2026, 3, 2)
        end = date(2026, 3, 4)
        sm.set_employee_days_off("EMP001", start, end, "夏季休暇")
        for i in range(3):
            d = start + timedelta(days=i)
            assert not sm.check_availability("EMP001", d, 1.0)

    def test_get_daily_summary(self):
        sm = ScheduleManager()
        target = date(2026, 3, 2)
        summary = sm.get_daily_summary(["EMP001", "EMP002"], target)
        assert "EMP001" in summary
        assert "EMP002" in summary
        assert summary["EMP001"]["available_hours"] == 8.0

    def test_find_earliest_available_date(self):
        sm = ScheduleManager()
        # 月曜日を休みにする
        sm.set_employee_day_off("EMP001", date(2026, 3, 2))
        earliest = sm.find_earliest_available_date(
            "EMP001", date(2026, 3, 2), 4.0
        )
        assert earliest == date(2026, 3, 3)  # 火曜日

    def test_get_team_utilization(self):
        sm = ScheduleManager()
        target = date(2026, 3, 2)
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
            estimated_hours=4.0,
        )
        sm.assign_task("EMP001", task, target)
        utils = sm.get_team_utilization(
            ["EMP001", "EMP002"],
            target,
            target,
        )
        assert utils["EMP001"] == 0.5
        assert utils["EMP002"] == 0.0


class TestReportingService:
    def _make_service(self):
        _, employees = _build_test_data()
        sm = ScheduleManager()
        tasks = [
            Task(
                task_id="T001",
                title="サーバー保守",
                description="テスト",
                region_id="tokyo",
                priority=TaskPriority.HIGH,
                task_type="maintenance",
                estimated_hours=4.0,
            ),
            Task(
                task_id="T002",
                title="ネットワーク設定",
                description="テスト",
                region_id="kanagawa",
                priority=TaskPriority.MEDIUM,
                task_type="installation",
                estimated_hours=6.0,
            ),
            Task(
                task_id="T003",
                title="DB移行",
                description="テスト",
                region_id="osaka",
                priority=TaskPriority.CRITICAL,
                task_type="repair",
                estimated_hours=8.0,
                deadline=date.today() - timedelta(days=1),
            ),
        ]
        # T001を割り当て済みにする
        target = date(2026, 3, 2)
        sm.assign_task("EMP001", tasks[0], target)

        return ReportingService(employees, tasks, sm), tasks, employees, sm

    def test_allocation_summary(self):
        service, tasks, _, _ = self._make_service()
        summary = service.generate_allocation_summary()
        assert summary.total_tasks == 3
        assert summary.assigned_tasks == 1
        assert summary.unassigned_tasks == 2
        assert summary.overdue_tasks == 1
        assert "HIGH" in summary.tasks_by_priority
        assert "tokyo" in summary.tasks_by_region

    def test_utilization_report(self):
        service, _, _, _ = self._make_service()
        start = date(2026, 3, 2)
        end = date(2026, 3, 6)
        report = service.generate_utilization_report(start, end)
        assert report.period_start == start
        assert len(report.employee_utilizations) == 3

    def test_region_workload(self):
        service, _, _, _ = self._make_service()
        reports = service.generate_region_workload()
        assert len(reports) == 3  # tokyo, kanagawa, osaka
        # 各地域に1タスクずつ
        for r in reports:
            assert r.total_tasks == 1

    def test_employee_workload_summary(self):
        service, _, _, _ = self._make_service()
        summaries = service.generate_employee_workload_summary(date(2026, 3, 2))
        assert len(summaries) == 3
        # 山田が最も高い稼働率
        assert summaries[0]["employee_id"] == "EMP001"

    def test_identify_bottlenecks(self):
        service, _, _, _ = self._make_service()
        bottlenecks = service.identify_bottlenecks()
        assert bottlenecks["total_unassigned"] == 2
        assert bottlenecks["high_priority_unassigned"] >= 1
