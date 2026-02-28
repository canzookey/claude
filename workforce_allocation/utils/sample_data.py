"""サンプルデータファクトリー - デモ・テスト用のデータを生成する"""

from __future__ import annotations

from datetime import date, timedelta

from workforce_allocation.models.employee import Employee, Skill, SkillLevel
from workforce_allocation.models.region import Region, RegionMap
from workforce_allocation.models.task import Task, TaskPriority


def create_sample_region_map() -> RegionMap:
    """関東・関西を中心としたサンプル地域マップを生成する"""
    rm = RegionMap()

    regions = [
        Region("tokyo", "東京", "東京都", "urban", adjacent_region_ids=["kanagawa", "saitama", "chiba"]),
        Region("kanagawa", "神奈川", "神奈川県", "urban", adjacent_region_ids=["tokyo"]),
        Region("saitama", "埼玉", "埼玉県", "suburban", adjacent_region_ids=["tokyo", "chiba"]),
        Region("chiba", "千葉", "千葉県", "suburban", adjacent_region_ids=["tokyo", "saitama"]),
        Region("osaka", "大阪", "大阪府", "urban", adjacent_region_ids=["kyoto", "hyogo"]),
        Region("kyoto", "京都", "京都府", "urban", adjacent_region_ids=["osaka"]),
        Region("hyogo", "兵庫", "兵庫県", "urban", adjacent_region_ids=["osaka"]),
        Region("nagoya", "名古屋", "愛知県", "urban", adjacent_region_ids=[]),
    ]

    for r in regions:
        rm.add_region(r)

    # 関東エリアの距離
    rm.set_distance("tokyo", "kanagawa", 30, 45)
    rm.set_distance("tokyo", "saitama", 35, 50)
    rm.set_distance("tokyo", "chiba", 40, 55)
    rm.set_distance("kanagawa", "saitama", 60, 80)
    rm.set_distance("kanagawa", "chiba", 70, 90)
    rm.set_distance("saitama", "chiba", 50, 65)

    # 関西エリアの距離
    rm.set_distance("osaka", "kyoto", 45, 30)
    rm.set_distance("osaka", "hyogo", 35, 25)
    rm.set_distance("kyoto", "hyogo", 75, 60)

    # エリア間の距離
    rm.set_distance("tokyo", "osaka", 500, 150)
    rm.set_distance("tokyo", "nagoya", 350, 100)
    rm.set_distance("osaka", "nagoya", 180, 55)
    rm.set_distance("nagoya", "kanagawa", 320, 95)
    rm.set_distance("nagoya", "saitama", 330, 100)

    return rm


def create_sample_employees() -> list[Employee]:
    """サンプル従業員データを生成する"""
    return [
        Employee(
            employee_id="EMP001",
            name="山田太郎",
            department="フィールドサービス第一課",
            home_region_id="tokyo",
            skills=[
                Skill("サーバー構築", SkillLevel.EXPERT, certified=True, years_experience=12.0),
                Skill("ネットワーク", SkillLevel.ADVANCED, years_experience=8.0),
                Skill("クラウド(AWS)", SkillLevel.ADVANCED, certified=True, years_experience=5.0),
            ],
            role="senior_engineer",
            certifications=["AWS SAP", "LPIC-3"],
        ),
        Employee(
            employee_id="EMP002",
            name="鈴木花子",
            department="フィールドサービス第一課",
            home_region_id="kanagawa",
            skills=[
                Skill("ネットワーク", SkillLevel.EXPERT, certified=True, years_experience=10.0),
                Skill("セキュリティ", SkillLevel.ADVANCED, certified=True, years_experience=7.0),
                Skill("ファイアウォール", SkillLevel.EXPERT, years_experience=9.0),
            ],
            role="senior_engineer",
            certifications=["CCNP", "情報処理安全確保支援士"],
        ),
        Employee(
            employee_id="EMP003",
            name="佐藤一郎",
            department="フィールドサービス第二課",
            home_region_id="osaka",
            skills=[
                Skill("データベース", SkillLevel.EXPERT, certified=True, years_experience=15.0),
                Skill("サーバー構築", SkillLevel.INTERMEDIATE, years_experience=4.0),
                Skill("バックアップ", SkillLevel.ADVANCED, years_experience=8.0),
            ],
            role="senior_engineer",
            certifications=["Oracle Master Gold", "LPIC-2"],
        ),
        Employee(
            employee_id="EMP004",
            name="田中美咲",
            department="フィールドサービス第一課",
            home_region_id="tokyo",
            skills=[
                Skill("クラウド(AWS)", SkillLevel.EXPERT, certified=True, years_experience=6.0),
                Skill("クラウド(Azure)", SkillLevel.ADVANCED, certified=True, years_experience=4.0),
                Skill("コンテナ", SkillLevel.ADVANCED, years_experience=3.0),
            ],
            role="field_engineer",
            certifications=["AWS SAA", "AZ-104"],
        ),
        Employee(
            employee_id="EMP005",
            name="高橋健二",
            department="フィールドサービス第二課",
            home_region_id="saitama",
            skills=[
                Skill("サーバー構築", SkillLevel.INTERMEDIATE, years_experience=3.0),
                Skill("ネットワーク", SkillLevel.INTERMEDIATE, years_experience=3.0),
                Skill("ヘルプデスク", SkillLevel.ADVANCED, years_experience=5.0),
            ],
            role="field_engineer",
        ),
        Employee(
            employee_id="EMP006",
            name="伊藤直美",
            department="フィールドサービス第二課",
            home_region_id="osaka",
            skills=[
                Skill("ネットワーク", SkillLevel.ADVANCED, years_experience=6.0),
                Skill("VoIP", SkillLevel.EXPERT, certified=True, years_experience=8.0),
                Skill("セキュリティ", SkillLevel.INTERMEDIATE, years_experience=3.0),
            ],
            role="field_engineer",
            certifications=["CCNA"],
        ),
        Employee(
            employee_id="EMP007",
            name="渡辺修平",
            department="フィールドサービス第一課",
            home_region_id="nagoya",
            skills=[
                Skill("サーバー構築", SkillLevel.ADVANCED, years_experience=7.0),
                Skill("ストレージ", SkillLevel.EXPERT, certified=True, years_experience=10.0),
                Skill("バックアップ", SkillLevel.ADVANCED, years_experience=6.0),
            ],
            role="field_engineer",
            certifications=["NetApp NCDA"],
        ),
        Employee(
            employee_id="EMP008",
            name="小林あゆみ",
            department="フィールドサービス第一課",
            home_region_id="chiba",
            skills=[
                Skill("ヘルプデスク", SkillLevel.EXPERT, years_experience=8.0),
                Skill("PC設定", SkillLevel.EXPERT, years_experience=10.0),
                Skill("ネットワーク", SkillLevel.BEGINNER, years_experience=2.0),
            ],
            role="field_engineer",
        ),
    ]


def create_sample_tasks(base_date: date | None = None) -> list[Task]:
    """サンプルタスクデータを生成する"""
    if base_date is None:
        base_date = date.today()

    return [
        Task(
            task_id="TASK001",
            title="本社サーバールーム定期点検",
            description="四半期ごとのサーバールーム設備点検。温度・湿度確認、ケーブル点検を含む。",
            region_id="tokyo",
            priority=TaskPriority.MEDIUM,
            required_skills={"サーバー構築": SkillLevel.INTERMEDIATE},
            estimated_hours=4.0,
            deadline=base_date + timedelta(days=7),
            customer_name="本社IT部門",
            customer_id="CUS001",
            task_type="maintenance",
        ),
        Task(
            task_id="TASK002",
            title="新規拠点ネットワーク構築",
            description="横浜新拠点のネットワーク環境構築。スイッチ、AP、ファイアウォール設定。",
            region_id="kanagawa",
            priority=TaskPriority.HIGH,
            required_skills={
                "ネットワーク": SkillLevel.ADVANCED,
                "ファイアウォール": SkillLevel.INTERMEDIATE,
            },
            estimated_hours=8.0,
            deadline=base_date + timedelta(days=5),
            customer_name="ABCコンサルティング",
            customer_id="CUS002",
            task_type="installation",
        ),
        Task(
            task_id="TASK003",
            title="データベース障害復旧",
            description="本番DBサーバーのディスク障害。データ整合性確認とフェイルオーバー対応。",
            region_id="osaka",
            priority=TaskPriority.CRITICAL,
            required_skills={"データベース": SkillLevel.ADVANCED},
            estimated_hours=6.0,
            deadline=base_date + timedelta(days=1),
            customer_name="XYZ製造",
            customer_id="CUS003",
            task_type="repair",
        ),
        Task(
            task_id="TASK004",
            title="クラウド移行コンサルティング",
            description="オンプレミスからAWSへの移行計画策定。現状評価と移行戦略の提案。",
            region_id="tokyo",
            priority=TaskPriority.MEDIUM,
            required_skills={"クラウド(AWS)": SkillLevel.ADVANCED},
            estimated_hours=4.0,
            deadline=base_date + timedelta(days=14),
            customer_name="DEFサービス",
            customer_id="CUS004",
            task_type="consultation",
        ),
        Task(
            task_id="TASK005",
            title="VoIPシステム導入",
            description="IP電話システムの設計・導入。既存PBXからの切り替え作業。",
            region_id="osaka",
            priority=TaskPriority.HIGH,
            required_skills={"VoIP": SkillLevel.ADVANCED},
            estimated_hours=8.0,
            deadline=base_date + timedelta(days=10),
            customer_name="GHI商事",
            customer_id="CUS005",
            task_type="installation",
        ),
        Task(
            task_id="TASK006",
            title="セキュリティ監査対応",
            description="年次セキュリティ監査のための設定確認・ログ収集・レポート作成。",
            region_id="tokyo",
            priority=TaskPriority.HIGH,
            required_skills={"セキュリティ": SkillLevel.ADVANCED},
            estimated_hours=6.0,
            deadline=base_date + timedelta(days=3),
            customer_name="本社コンプライアンス部",
            customer_id="CUS001",
            task_type="maintenance",
        ),
        Task(
            task_id="TASK007",
            title="ストレージ増設作業",
            description="名古屋DCのストレージ筐体増設。既存RAIDグループの拡張。",
            region_id="nagoya",
            priority=TaskPriority.MEDIUM,
            required_skills={
                "ストレージ": SkillLevel.ADVANCED,
                "サーバー構築": SkillLevel.INTERMEDIATE,
            },
            estimated_hours=6.0,
            deadline=base_date + timedelta(days=7),
            customer_name="名古屋DC運用チーム",
            customer_id="CUS006",
            task_type="installation",
        ),
        Task(
            task_id="TASK008",
            title="PC大量キッティング",
            description="新入社員向けPC50台のセットアップ。OSインストール、ドメイン参加、ソフト設定。",
            region_id="chiba",
            priority=TaskPriority.MEDIUM,
            required_skills={"PC設定": SkillLevel.INTERMEDIATE},
            estimated_hours=8.0,
            deadline=base_date + timedelta(days=14),
            customer_name="JKL物流",
            customer_id="CUS007",
            task_type="installation",
        ),
    ]
