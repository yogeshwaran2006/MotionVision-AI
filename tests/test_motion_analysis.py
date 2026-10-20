import pytest
import time
import math
from motion_analysis import MotionAnalyzer

# Helper to generate timestamps
def ts(offset):
    return time.time() + offset


def test_calculate_center():
    bbox = (0, 0, 10, 20)
    cx, cy = MotionAnalyzer.calculate_center(bbox)
    assert cx == 5.0
    assert cy == 10.0

def test_ema_smoothing_stationary():
    analyzer = MotionAnalyzer(alpha=0.7)
    track_id = 1
    # Simulate 5 identical boxes (no movement)
    for i in range(5):
        analyzer.update(track_id, (0, 0, 10, 10), ts(i))
    history = analyzer.get_history(track_id)
    # All smoothed points should be the same (center 5,5)
    for x, y, _ in history:
        assert x == 5.0 and y == 5.0

def test_ema_smoothing_moving():
    analyzer = MotionAnalyzer(alpha=0.5)
    track_id = 2
    # Move right by 10 pixels each frame
    for i in range(4):
        bbox = (i*10, 0, i*10+10, 10)
        analyzer.update(track_id, bbox, ts(i))
    history = analyzer.get_history(track_id)
    # The last smoothed x should be between the raw centre (45) and previous
    last_x, _, _ = history[-1]
    assert 30 < last_x < 45

def test_pixel_speed_constant_velocity():
    analyzer = MotionAnalyzer(alpha=1.0)  # no smoothing to keep raw values
    track_id = 3
    # Move 100 pixels over 5 seconds
    analyzer.update(track_id, (0, 0, 10, 10), ts(0))
    analyzer.update(track_id, (0, 0, 110, 10), ts(5))  # centre moves from 5 to 55
    speed = analyzer.pixel_speed(track_id)
    # distance = 50 pixels, dt = 5s -> speed = 10 px/s
    assert pytest.approx(speed, rel=1e-3) == 10.0

def test_direction_stationary_and_compass():
    analyzer = MotionAnalyzer(alpha=1.0, min_distance_threshold=2.0)
    track_id = 4
    # Stationary case
    analyzer.update(track_id, (0, 0, 10, 10), ts(0))
    analyzer.update(track_id, (0, 0, 10, 10), ts(1))
    assert analyzer.direction(track_id) == "Stationary"
    # East movement
    analyzer.reset()
    analyzer.update(track_id, (0, 0, 10, 10), ts(0))
    analyzer.update(track_id, (20, 0, 30, 10), ts(1))  # centre moves from (5,5) to (25,5)
    assert analyzer.direction(track_id) == "E"
    # North-East movement (note y increase downwards, so moving up means decreasing y)
    analyzer.reset()
    analyzer.update(track_id, (0, 20, 10, 30), ts(0))  # centre (5,25)
    analyzer.update(track_id, (20, 0, 30, 10), ts(1))   # centre (25,5) -> dx=20, dy=-20
    assert analyzer.direction(track_id) == "NE"

def test_edge_cases():
    analyzer = MotionAnalyzer()
    # No history
    assert analyzer.pixel_speed(999) is None
    assert analyzer.direction(999) is None
    # Single point history
    analyzer.update(5, (0, 0, 10, 10), ts(0))
    assert analyzer.pixel_speed(5) is None
    assert analyzer.direction(5) is None
    # Zero dt (same timestamp)
    analyzer.update(6, (0, 0, 10, 10), ts(0))
    analyzer.update(6, (20, 0, 30, 10), ts(0))
    assert analyzer.pixel_speed(6) is None
