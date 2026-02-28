"""データモデルのテスト"""

import pytest
from datetime import date, timedelta

from workforce_allocation.models.employee import Employee, Skill, SkillLevel
from workforce_allocation.models.region import Region, RegionMap
from workforce_allocation.models.task import Task, TaskPriority, TaskStatus
from workforce_allocation.models.schedule import Schedule, TimeSlot


# --- Employee テスト ---


class TestSkill:
    def test_weighted_score_basic(self):
        skill = Skill(name="ネットワーク", level=SkillLevel.INTERMEDIATE)
        assert skill.weighted_score == 2.0

    def test_weighted_score_with_certification(self):
        skill = Skill(name="ネットワーク", level=SkillLevel.INTERMEDIATE, certified=True)
        assert skill.weighted_score == 3.0  # 2 * 1.5

    def test_weighted_score_with_experience(self):
        skill = Skill(
            name="ネットワーク",
            level=SkillLevel.ADVANCED,
            years_experience=10.0,
        )
        assert skill.weighted_score == 3.5  # 3 * 1.0 + 0.5


class TestEmployee:
    def _make_employee(self, **kwargs):
        defaults = {
            "employee_id": "EMP001",
            "name": "山田太郎",
            "department": "フィールドサービス",
            "home_region_id": "tokyo",
        }
        defaults.update(kwargs)
        return Employee(**defaults)

    def test_has_skill(self):
        emp = self._make_employee(
            skills=[Skill(name="サーバー", level=SkillLevel.ADVANCED)]
        )
        assert emp.has_skill("サーバー", SkillLevel.BEGINNER)
        assert emp.has_skill("サーバー", SkillLevel.ADVANCED)
        assert not emp.has_skill("サーバー", SkillLevel.EXPERT)
        assert not emp.has_skill("データベース")

    def test_get_skill(self):
        skill = Skill(name="サーバー", level=SkillLevel.ADVANCED)
        emp = self._make_employee(skills=[skill])
        assert emp.get_skill("サーバー") is skill
        assert emp.get_skill("データベース") is None

    def test_skill_match_score_full_match(self):
        emp = self._make_employee(
            skills=[
                Skill(name="サーバー", level=SkillLevel.ADVANCED),
                Skill(name="ネットワーク", level=SkillLevel.INTERMEDIATE),
            ]
        )
        required = {
            "サーバー": SkillLevel.INTERMEDIATE,
            "ネットワーク": SkillLevel.BEGINNER,
        }
        score = emp.skill_match_score(required)
        assert score == 1.0

    def test_skill_match_score_no_skills(self):
        emp = self._make_employee(skills=[])
        required = {"サーバー": SkillLevel.INTERMEDIATE}
        score = emp.skill_match_score(required)
        assert score == 0.0

    def test_skill_match_score_empty_requirements(self):
        emp = self._make_employee()
        score = emp.skill_match_score({})
        assert score == 1.0

    def test_meets_requirements(self):
        emp = self._make_employee(
            skills=[
                Skill(name="サーバー", level=SkillLevel.ADVANCED),
                Skill(name="ネットワーク", level=SkillLevel.INTERMEDIATE),
            ]
        )
        assert emp.meets_requirements(
            {"サーバー": SkillLevel.INTERMEDIATE}
        )
        assert not emp.meets_requirements(
            {"データベース": SkillLevel.BEGINNER}
        )


# --- Region テスト ---


class TestRegionMap:
    def _make_region_map(self):
        rm = RegionMap()
        tokyo = Region(
            region_id="tokyo",
            name="東京",
            prefecture="東京都",
            adjacent_region_ids=["kanagawa", "saitama", "chiba"],
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
        rm.add_region(tokyo)
        rm.add_region(kanagawa)
        rm.add_region(osaka)
        rm.set_distance("tokyo", "kanagawa", 30.0, 45.0)
        rm.set_distance("tokyo", "osaka", 500.0, 150.0)
        return rm

    def test_add_and_get_region(self):
        rm = self._make_region_map()
        region = rm.get_region("tokyo")
        assert region is not None
        assert region.name == "東京"
        assert rm.get_region("unknown") is None

    def test_distance_bidirectional(self):
        rm = self._make_region_map()
        d1 = rm.get_distance("tokyo", "kanagawa")
        d2 = rm.get_distance("kanagawa", "tokyo")
        assert d1 is not None
        assert d2 is not None
        assert d1.distance_km == d2.distance_km

    def test_same_region_distance(self):
        rm = self._make_region_map()
        d = rm.get_distance("tokyo", "tokyo")
        assert d is not None
        assert d.distance_km == 0.0

    def test_travel_time(self):
        rm = self._make_region_map()
        assert rm.get_travel_time("tokyo", "kanagawa") == 45.0
        assert rm.get_travel_time("tokyo", "unknown") == 120.0  # デフォルト

    def test_nearby_regions(self):
        rm = self._make_region_map()
        nearby = rm.get_nearby_regions("tokyo", max_travel_minutes=60.0)
        assert "kanagawa" in nearby
        assert "osaka" not in nearby

    def test_adjacent(self):
        rm = self._make_region_map()
        tokyo = rm.get_region("tokyo")
        assert tokyo.is_adjacent("kanagawa")
        assert not tokyo.is_adjacent("osaka")


# --- Task テスト ---


class TestTask:
    def test_default_status(self):
        task = Task(
            task_id="T001",
            title="サーバー保守",
            description="定期メンテナンス",
            region_id="tokyo",
        )
        assert task.status == TaskStatus.UNASSIGNED

    def test_assign(self):
        task = Task(
            task_id="T001",
            title="サーバー保守",
            description="定期メンテナンス",
            region_id="tokyo",
        )
        task.assign("EMP001", date(2026, 3, 1))
        assert task.status == TaskStatus.ASSIGNED
        assert task.assigned_employee_id == "EMP001"
        assert task.scheduled_date == date(2026, 3, 1)

    def test_complete(self):
        task = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
        )
        task.complete()
        assert task.status == TaskStatus.COMPLETED
        assert task.completed_at is not None

    def test_urgency_score_critical_overdue(self):
        task = Task(
            task_id="T001",
            title="緊急修理",
            description="テスト",
            region_id="tokyo",
            priority=TaskPriority.CRITICAL,
            deadline=date.today() - timedelta(days=1),
        )
        assert task.urgency_score == 10.0  # min(4*2 + 3, 10)

    def test_urgency_score_low_no_deadline(self):
        task = Task(
            task_id="T001",
            title="定期チェック",
            description="テスト",
            region_id="tokyo",
            priority=TaskPriority.LOW,
        )
        assert task.urgency_score == 2.0  # 1*2

    def test_is_overdue(self):
        past = Task(
            task_id="T001",
            title="テスト",
            description="テスト",
            region_id="tokyo",
            deadline=date.today() - timedelta(days=1),
        )
        assert past.is_overdue

        future = Task(
            task_id="T002",
            title="テスト",
            description="テスト",
            region_id="tokyo",
            deadline=date.today() + timedelta(days=5),
        )
        assert not future.is_overdue


# --- Schedule テスト ---


class TestTimeSlot:
    def test_remaining_hours(self):
        slot = TimeSlot(
            slot_date=date(2026, 3, 2),
            employee_id="EMP001",
            available_hours=8.0,
            allocated_hours=3.0,
        )
        assert slot.remaining_hours == 5.0

    def test_day_off(self):
        slot = TimeSlot(
            slot_date=date(2026, 3, 1),
            employee_id="EMP001",
            is_day_off=True,
        )
        assert slot.remaining_hours == 0.0
        assert not slot.can_allocate(1.0)

    def test_allocate_and_deallocate(self):
        slot = TimeSlot(
            slot_date=date(2026, 3, 2),
            employee_id="EMP001",
        )
        assert slot.allocate("T001", 4.0)
        assert slot.allocated_hours == 4.0
        assert "T001" in slot.task_ids

        slot.deallocate("T001", 4.0)
        assert slot.allocated_hours == 0.0
        assert "T001" not in slot.task_ids

    def test_allocate_exceeds_capacity(self):
        slot = TimeSlot(
            slot_date=date(2026, 3, 2),
            employee_id="EMP001",
            available_hours=8.0,
            allocated_hours=6.0,
        )
        assert not slot.can_allocate(4.0)
        assert not slot.allocate("T001", 4.0)

    def test_utilization_rate(self):
        slot = TimeSlot(
            slot_date=date(2026, 3, 2),
            employee_id="EMP001",
            available_hours=8.0,
            allocated_hours=4.0,
        )
        assert slot.utilization_rate == 0.5


class TestSchedule:
    def test_get_slot_auto_create(self):
        schedule = Schedule("EMP001")
        # 平日
        slot = schedule.get_slot(date(2026, 3, 2))  # 月曜日
        assert not slot.is_day_off

    def test_weekend_auto_day_off(self):
        schedule = Schedule("EMP001")
        slot = schedule.get_slot(date(2026, 3, 1))  # 日曜日
        assert slot.is_day_off

    def test_available_dates(self):
        schedule = Schedule("EMP001")
        start = date(2026, 3, 2)  # 月曜日
        end = date(2026, 3, 6)  # 金曜日
        available = schedule.get_available_dates(start, end, 4.0)
        assert len(available) == 5  # 平日5日間

    def test_allocate_task(self):
        schedule = Schedule("EMP001")
        target = date(2026, 3, 2)  # 月曜日
        assert schedule.allocate_task(target, "T001", 4.0)
        slot = schedule.get_slot(target)
        assert slot.allocated_hours == 4.0

    def test_set_day_off(self):
        schedule = Schedule("EMP001")
        target = date(2026, 3, 3)  # 火曜日
        schedule.set_day_off(target, "有給休暇")
        slot = schedule.get_slot(target)
        assert slot.is_day_off
        assert slot.note == "有給休暇"
