"""
要員配置システム デモスクリプト

フィールドサービス事業向けの要員配置を実行し、
配置結果・稼働率レポート・ボトルネック分析を表示する。
"""

from datetime import date, timedelta

from workforce_allocation.models.schedule import Schedule
from workforce_allocation.services.allocation_engine import AllocationEngine
from workforce_allocation.services.schedule_manager import ScheduleManager
from workforce_allocation.services.reporting import ReportingService
from workforce_allocation.utils.sample_data import (
    create_sample_employees,
    create_sample_region_map,
    create_sample_tasks,
)


def main() -> None:
    print("=" * 70)
    print("  要員配置システム (Workforce Allocation System)")
    print("  フィールドサービス事業展開 デモ")
    print("=" * 70)
    print()

    # --- データ準備 ---
    region_map = create_sample_region_map()
    employees = create_sample_employees()
    # 月曜日を対象日とする
    today = date.today()
    days_until_monday = (7 - today.weekday()) % 7
    if days_until_monday == 0 and today.weekday() != 0:
        days_until_monday = 7
    target_date = today + timedelta(days=days_until_monday) if today.weekday() != 0 else today
    tasks = create_sample_tasks(target_date)

    schedule_manager = ScheduleManager()

    # --- 配置エンジン初期化 ---
    engine = AllocationEngine(
        employees=employees,
        region_map=region_map,
        schedules=schedule_manager.schedules,
        skill_weight=0.4,
        region_weight=0.3,
        availability_weight=0.3,
    )

    # スケジュールの初期化
    for emp in employees:
        schedule_manager.get_or_create_schedule(emp.employee_id)

    # --- 従業員一覧 ---
    print("【登録従業員一覧】")
    print("-" * 70)
    print(f"{'ID':<8} {'氏名':<12} {'所属地域':<8} {'役割':<18} {'スキル数'}")
    print("-" * 70)
    for emp in employees:
        region = region_map.get_region(emp.home_region_id)
        region_name = region.name if region else emp.home_region_id
        print(
            f"{emp.employee_id:<8} {emp.name:<12} {region_name:<8} "
            f"{emp.role:<18} {len(emp.skills)}"
        )
    print()

    # --- タスク一覧 ---
    print("【配置対象タスク一覧】")
    print("-" * 70)
    print(f"{'ID':<10} {'タイトル':<24} {'地域':<6} {'優先度':<10} {'工数(h)'}")
    print("-" * 70)
    for task in tasks:
        region = region_map.get_region(task.region_id)
        region_name = region.name if region else task.region_id
        print(
            f"{task.task_id:<10} {task.title:<24} {region_name:<6} "
            f"{task.priority.name:<10} {task.estimated_hours}"
        )
    print()

    # --- 一括配置実行 ---
    print("【配置実行結果】")
    print("=" * 70)
    results = engine.allocate_batch(tasks, target_date)

    success_count = 0
    fail_count = 0

    for result in results:
        status = "OK" if result.success else "NG"
        if result.success:
            success_count += 1
            emp = result.assigned_employee
            candidate = result.candidate
            print(
                f"  [{status}] {result.task.title}"
            )
            print(
                f"       → {emp.name} "
                f"(スコア: {candidate.total_score:.2f}, "
                f"スキル: {candidate.skill_score:.2f}, "
                f"地域: {candidate.region_score:.2f}, "
                f"空き: {candidate.availability_score:.2f})"
            )
        else:
            fail_count += 1
            print(f"  [{status}] {result.task.title}")
            print(f"       理由: {result.reason}")
        print()

    print(f"配置結果: {success_count}件成功 / {fail_count}件失敗 / 合計{len(results)}件")
    print()

    # --- レポーティング ---
    reporting = ReportingService(employees, tasks, schedule_manager)

    # 配置サマリー
    summary = reporting.generate_allocation_summary()
    print("【配置サマリー】")
    print("-" * 40)
    print(f"  総タスク数:     {summary.total_tasks}")
    print(f"  配置済み:       {summary.assigned_tasks}")
    print(f"  未配置:         {summary.unassigned_tasks}")
    print(f"  配置率:         {summary.assignment_rate:.0%}")
    print(f"  期限超過:       {summary.overdue_tasks}")
    print()
    print("  優先度別:")
    for priority, count in sorted(summary.tasks_by_priority.items()):
        print(f"    {priority}: {count}件")
    print()
    print("  タスクタイプ別:")
    for ttype, count in sorted(summary.tasks_by_type.items()):
        print(f"    {ttype}: {count}件")
    print()

    # 稼働率レポート
    week_end = target_date + timedelta(days=4)
    util_report = reporting.generate_utilization_report(target_date, week_end)
    print("【週間稼働率レポート】")
    print(f"  対象期間: {util_report.period_start} 〜 {util_report.period_end}")
    print("-" * 50)
    for emp_id, util in sorted(
        util_report.employee_utilizations.items(),
        key=lambda x: x[1],
        reverse=True,
    ):
        emp = next((e for e in employees if e.employee_id == emp_id), None)
        name = emp.name if emp else emp_id
        bar = "█" * int(util * 20) + "░" * (20 - int(util * 20))
        print(f"  {name:<12} [{bar}] {util:.0%}")
    print()
    print(f"  チーム平均稼働率: {util_report.team_average:.0%}")
    if util_report.overloaded_employees:
        names = [
            next((e.name for e in employees if e.employee_id == eid), eid)
            for eid in util_report.overloaded_employees
        ]
        print(f"  ⚠ 高負荷従業員: {', '.join(names)}")
    if util_report.underutilized_employees:
        names = [
            next((e.name for e in employees if e.employee_id == eid), eid)
            for eid in util_report.underutilized_employees
        ]
        print(f"  ⚠ 低稼働従業員: {', '.join(names)}")
    print()

    # 地域別ワークロード
    region_reports = reporting.generate_region_workload()
    print("【地域別ワークロード】")
    print("-" * 50)
    for rr in region_reports:
        region = region_map.get_region(rr.region_id)
        name = region.name if region else rr.region_id
        print(
            f"  {name:<8} タスク: {rr.total_tasks}件 "
            f"(配置済: {rr.assigned_tasks}件) "
            f"工数: {rr.total_hours:.0f}h "
            f"人員: {rr.assigned_employees}名"
        )
    print()

    # ボトルネック分析
    bottlenecks = reporting.identify_bottlenecks()
    print("【ボトルネック分析】")
    print("-" * 50)
    print(f"  未配置タスク: {bottlenecks['total_unassigned']}件")
    print(f"  高優先度未配置: {bottlenecks['high_priority_unassigned']}件")
    if bottlenecks["skill_shortage"]:
        print("  スキル不足:")
        for skill, info in bottlenecks["skill_shortage"].items():
            print(
                f"    {skill}: 需要{info['demand']}件 / "
                f"供給{info['supply']}名 (不足: {info['gap']}名)"
            )
    if bottlenecks["unassigned_by_region"]:
        print("  地域別未配置:")
        for region_id, count in bottlenecks["unassigned_by_region"].items():
            region = region_map.get_region(region_id)
            name = region.name if region else region_id
            print(f"    {name}: {count}件")
    print()
    print("=" * 70)
    print("  デモ完了")
    print("=" * 70)


if __name__ == "__main__":
    main()
