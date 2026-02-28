"""地域モデル - サービスエリアと距離管理"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Region:
    """サービス地域"""

    region_id: str
    name: str
    prefecture: str  # 都道府県
    area_type: str = "urban"  # urban, suburban, rural
    latitude: float = 0.0
    longitude: float = 0.0
    adjacent_region_ids: list[str] = field(default_factory=list)

    def is_adjacent(self, other_region_id: str) -> bool:
        """隣接地域かどうか判定する"""
        return other_region_id in self.adjacent_region_ids


@dataclass
class RegionDistance:
    """地域間の距離・移動時間情報"""

    from_region_id: str
    to_region_id: str
    distance_km: float
    travel_time_minutes: float

    @property
    def travel_time_hours(self) -> float:
        return self.travel_time_minutes / 60.0


class RegionMap:
    """地域マップ - 地域間の距離・移動時間を管理する"""

    def __init__(self) -> None:
        self._regions: dict[str, Region] = {}
        self._distances: dict[tuple[str, str], RegionDistance] = {}

    def add_region(self, region: Region) -> None:
        self._regions[region.region_id] = region

    def get_region(self, region_id: str) -> Region | None:
        return self._regions.get(region_id)

    @property
    def regions(self) -> list[Region]:
        return list(self._regions.values())

    def set_distance(
        self,
        from_id: str,
        to_id: str,
        distance_km: float,
        travel_time_minutes: float,
    ) -> None:
        """地域間の距離を設定する（双方向）"""
        d = RegionDistance(from_id, to_id, distance_km, travel_time_minutes)
        self._distances[(from_id, to_id)] = d
        d_rev = RegionDistance(to_id, from_id, distance_km, travel_time_minutes)
        self._distances[(to_id, from_id)] = d_rev

    def get_distance(self, from_id: str, to_id: str) -> RegionDistance | None:
        if from_id == to_id:
            return RegionDistance(from_id, to_id, 0.0, 0.0)
        return self._distances.get((from_id, to_id))

    def get_travel_time(self, from_id: str, to_id: str) -> float:
        """移動時間（分）を取得する。不明な場合はデフォルト値を返す。"""
        dist = self.get_distance(from_id, to_id)
        if dist is not None:
            return dist.travel_time_minutes
        return 120.0  # デフォルト: 2時間

    def get_nearby_regions(self, region_id: str, max_travel_minutes: float = 60.0) -> list[str]:
        """指定時間内でアクセス可能な近隣地域を取得する"""
        nearby = []
        for (from_id, to_id), dist in self._distances.items():
            if from_id == region_id and dist.travel_time_minutes <= max_travel_minutes:
                nearby.append(to_id)
        return nearby
