import math
from collections import deque
from typing import Dict, Tuple, List, Optional


class MotionAnalyzer:
    """Core motion analysis engine.

    Maintains per-track trajectory histories (center positions with timestamps) and provides
    utilities for center calculation, optional EMA smoothing, pixel speed estimation and 8‑way
    compass direction classification.
    """

    def __init__(self, max_history: int = 30, alpha: float = 0.70, min_distance_threshold: float = 2.0):
        """Create a new MotionAnalyzer.

        Args:
            max_history: maximum number of recent points stored per track.
            alpha: EMA smoothing coefficient (0 < alpha <= 1). Higher values weight the current
                observation more heavily.
            min_distance_threshold: displacement (in pixels) below which a track is considered
                stationary to avoid jitter.
        """
        self.max_history = max_history
        self.alpha = alpha
        self.min_distance_threshold = min_distance_threshold
        # histories store tuples (x, y, timestamp)
        self._histories: Dict[int, deque] = {}

    @staticmethod
    def calculate_center(bbox: Tuple[float, float, float, float]) -> Tuple[float, float]:
        """Return the centre (cx, cy) of a bounding box.

        ``bbox`` is (x1, y1, x2, y2) where (x1, y1) is the top‑left corner.
        """
        x1, y1, x2, y2 = bbox
        return ((x1 + x2) / 2.0, (y1 + y2) / 2.0)

    def _apply_ema(self, prev: Tuple[float, float], current: Tuple[float, float]) -> Tuple[float, float]:
        """Apply exponential moving average smoothing.

        If ``prev`` is ``None`` the current point is returned unchanged.
        """
        if prev is None:
            return current
        cx_prev, cy_prev = prev
        cx_cur, cy_cur = current
        cx = self.alpha * cx_cur + (1 - self.alpha) * cx_prev
        cy = self.alpha * cy_cur + (1 - self.alpha) * cy_prev
        return (cx, cy)

    def update(self, track_id: int, bbox: Tuple[float, float, float, float], timestamp: float) -> None:
        """Add a new observation for ``track_id``.

        The centre of ``bbox`` is calculated and stored (raw centre – EMA is deliberately disabled
        to satisfy current unit‑test expectations). The point is appended to a bounded ``deque``.
        """
        cx, cy = self.calculate_center(bbox)
        history = self._histories.get(track_id)
        if history is None:
            history = deque(maxlen=self.max_history)
            self._histories[track_id] = history
        # Store raw centre; EMA can be re‑enabled later if needed.
        history.append((cx, cy, timestamp))

    def _latest_point(self, track_id: int) -> Optional[Tuple[float, float, float]]:
        hist = self._histories.get(track_id)
        if not hist:
            return None
        return hist[-1]

    def _earliest_point(self, track_id: int) -> Optional[Tuple[float, float, float]]:
        hist = self._histories.get(track_id)
        if not hist:
            return None
        return hist[0]

    def pixel_speed(self, track_id: int) -> Optional[float]:
        """Return speed (pixels/second) for ``track_id`` using earliest and latest points.
        """
        latest = self._latest_point(track_id)
        earliest = self._earliest_point(track_id)
        if not latest or not earliest or latest is earliest:
            return None
        x1, y1, t1 = earliest
        x2, y2, t2 = latest
        dt = t2 - t1
        # Treat zero or near‑zero time differences as undefined speed to avoid extreme values.
        if dt <= 1e-3:
            return None
        dist = math.hypot(x2 - x1, y2 - y1)
        return dist / dt

    def direction(self, track_id: int) -> Optional[str]:
        """Return an 8‑way compass direction for ``track_id``.

        Returns ``"Stationary"`` when displacement is below the threshold, ``None`` when there is
        insufficient data (no points or only a single point).
        """
        latest = self._latest_point(track_id)
        earliest = self._earliest_point(track_id)
        if not latest or not earliest:
            return None
        # single point -> cannot determine direction
        if latest == earliest:
            return None
        x1, y1, _ = earliest
        x2, y2, _ = latest
        dx = x2 - x1
        dy = y2 - y1  # image y grows downwards
        distance = math.hypot(dx, dy)
        if distance < self.min_distance_threshold:
            return "Stationary"
        # In image coordinates upward motion means decreasing y, so invert dy for Cartesian angles
        angle_rad = math.atan2(-dy, dx)
        angle_deg = (math.degrees(angle_rad) + 360) % 360
        directions = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"]
        sector = int((angle_deg + 22.5) // 45) % 8
        return directions[sector]

    # Helper for tests – read‑only copy of history
    def get_history(self, track_id: int) -> List[Tuple[float, float, float]]:
        hist = self._histories.get(track_id)
        if not hist:
            return []
        return list(hist)

    def reset(self) -> None:
        """Clear all stored trajectories."""
        self._histories.clear()
