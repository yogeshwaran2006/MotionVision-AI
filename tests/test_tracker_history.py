"""Unit tests for VideoTracker — model-independent via monkey-patching.

These tests verify the tracking *logic* (history management, annotation drawing,
empty-frame handling, detection dict structure, class filtering, summary stats).
They do NOT prove real-world tracking accuracy — that requires running on actual
video (Milestone 2 manual test).
"""

import numpy as np
import cv2
import pytest
from tracker import VideoTracker


# ---------------------------------------------------------------------------
# Helpers — lightweight fakes that replace the real YOLO .track() call
# ---------------------------------------------------------------------------

class DummyBox:
    """Mimics the ultralytics ``Boxes`` object for a single detection."""
    def __init__(self, bbox, conf, cls, trk_id):
        self.xyxy = np.array([bbox], dtype=np.float32)
        self.conf = np.array([conf], dtype=np.float32)
        self.cls  = np.array([cls],  dtype=np.float32)
        self.id   = np.array([trk_id], dtype=np.float32) if trk_id is not None else None


class MultiDummyBox:
    """Mimics Boxes for multiple simultaneous detections."""
    def __init__(self, bboxes, confs, classes, trk_ids):
        self.xyxy = np.array(bboxes, dtype=np.float32)
        self.conf = np.array(confs,  dtype=np.float32)
        self.cls  = np.array(classes, dtype=np.float32)
        self.id   = np.array(trk_ids, dtype=np.float32) if trk_ids is not None else None


class DummyResult:
    def __init__(self, boxes):
        self.boxes = boxes


class EmptyResult:
    """Mimics a result with no detections."""
    def __init__(self):
        self.boxes = None


def _make_tracker(trail_len=5, classes=None, tracker_cfg=None):
    """Create a VideoTracker with given options (loads real model once)."""
    vt = VideoTracker(trail_len=trail_len, classes=classes,
                      tracker_cfg=tracker_cfg)
    return vt


DUMMY_FRAME = np.zeros((100, 100, 3), dtype=np.uint8)


# ---------------------------------------------------------------------------
# Original Milestone 2 tests (preserved)
# ---------------------------------------------------------------------------

def test_trail_history_bounded():
    """Centre-point trail is capped at trail_len."""
    vt = _make_tracker(trail_len=5)

    def fake_track(frame, **kwargs):
        return [DummyResult(DummyBox(bbox=[10, 10, 30, 30], conf=0.9,
                                     cls=0, trk_id=1))]
    vt.model.track = fake_track

    for _ in range(3):
        vt.process_frame(DUMMY_FRAME)
    assert len(vt.tracks_history[1]) == 3

    for _ in range(5):
        vt.process_frame(DUMMY_FRAME)
    assert len(vt.tracks_history[1]) == vt.trail_len


def test_empty_frame_returns_empty_list():
    """When the model finds nothing, process_frame returns []."""
    vt = _make_tracker()

    def fake_track(frame, **kwargs):
        return [EmptyResult()]
    vt.model.track = fake_track

    detections = vt.process_frame(DUMMY_FRAME)
    assert detections == []


def test_detection_dict_structure():
    """Each detection dict has the expected keys and correct types."""
    vt = _make_tracker()

    def fake_track(frame, **kwargs):
        return [DummyResult(DummyBox(bbox=[10, 20, 50, 60], conf=0.85,
                                     cls=0, trk_id=3))]
    vt.model.track = fake_track

    dets = vt.process_frame(DUMMY_FRAME)
    assert len(dets) == 1
    d = dets[0]
    for key in ("bbox", "score", "class_id", "class_name", "track_id"):
        assert key in d, f"Missing key: {key}"
    assert isinstance(d["bbox"], list) and len(d["bbox"]) == 4
    assert isinstance(d["score"], float)
    assert isinstance(d["class_id"], int)
    assert isinstance(d["class_name"], str)
    assert isinstance(d["track_id"], int)


def test_no_track_id_handled():
    """When boxes.id is None, track_id == -1."""
    vt = _make_tracker()

    def fake_track(frame, **kwargs):
        return [DummyResult(DummyBox(bbox=[10, 10, 30, 30], conf=0.7,
                                     cls=0, trk_id=None))]
    vt.model.track = fake_track

    dets = vt.process_frame(DUMMY_FRAME)
    assert len(dets) == 1
    assert dets[0]["track_id"] == -1


def test_draw_annotations_no_crash():
    """draw_annotations should not crash on valid detections or empty list."""
    vt = _make_tracker()
    frame = np.zeros((200, 200, 3), dtype=np.uint8)

    dets = [{"bbox": [10, 10, 50, 50], "score": 0.9, "class_id": 0,
             "class_name": "person", "track_id": 1}]
    vt.tracks_history[1].append((30, 30))
    vt.tracks_history[1].append((35, 35))
    result = vt.draw_annotations(frame.copy(), dets, draw_fps=25.0)
    assert result.shape == frame.shape

    result2 = vt.draw_annotations(frame.copy(), [], draw_fps=0.0)
    assert result2.shape == frame.shape


def test_multiple_tracks_independent():
    """Two simultaneous tracks maintain separate histories."""
    vt = _make_tracker(trail_len=4)
    call_count = [0]

    def fake_track(frame, **kwargs):
        call_count[0] += 1
        x_off = call_count[0] * 5
        return [DummyResult(MultiDummyBox(
            bboxes=[[10 + x_off, 10, 30 + x_off, 30],
                     [60 + x_off, 60, 80 + x_off, 80]],
            confs=[0.9, 0.8],
            classes=[0, 2],
            trk_ids=[1, 2],
        ))]
    vt.model.track = fake_track

    for _ in range(6):
        vt.process_frame(DUMMY_FRAME)

    assert 1 in vt.tracks_history
    assert 2 in vt.tracks_history
    assert len(vt.tracks_history[1]) == 4
    assert len(vt.tracks_history[2]) == 4
    assert vt.tracks_history[1][-1] != vt.tracks_history[2][-1]


# ---------------------------------------------------------------------------
# NEW — Class filtering tests
# ---------------------------------------------------------------------------

def test_class_filter_stored():
    """The classes parameter is stored and accessible."""
    vt_all = _make_tracker(classes=None)
    assert vt_all.classes is None  # default: all classes

    vt_filtered = _make_tracker(classes=[0, 2, 5])
    assert vt_filtered.classes == [0, 2, 5]


def test_class_filter_passed_to_track():
    """When classes is set, 'classes' kwarg is forwarded to model.track()."""
    vt = _make_tracker(classes=[0, 2])
    captured_kwargs = {}

    def spy_track(frame, **kwargs):
        captured_kwargs.update(kwargs)
        return [EmptyResult()]
    vt.model.track = spy_track

    vt.process_frame(DUMMY_FRAME)
    assert "classes" in captured_kwargs
    assert captured_kwargs["classes"] == [0, 2]


def test_no_class_filter_omits_kwarg():
    """When classes is None, 'classes' kwarg is NOT sent to model.track()."""
    vt = _make_tracker(classes=None)
    captured_kwargs = {}

    def spy_track(frame, **kwargs):
        captured_kwargs.update(kwargs)
        return [EmptyResult()]
    vt.model.track = spy_track

    vt.process_frame(DUMMY_FRAME)
    assert "classes" not in captured_kwargs


# ---------------------------------------------------------------------------
# NEW — Tracker config tests
# ---------------------------------------------------------------------------

def test_tracker_cfg_passed_to_track():
    """When tracker_cfg is set, 'tracker' kwarg is forwarded to model.track()."""
    vt = _make_tracker(tracker_cfg="botsort_reid.yaml")
    captured_kwargs = {}

    def spy_track(frame, **kwargs):
        captured_kwargs.update(kwargs)
        return [EmptyResult()]
    vt.model.track = spy_track

    vt.process_frame(DUMMY_FRAME)
    assert captured_kwargs.get("tracker") == "botsort_reid.yaml"


def test_default_tracker_omits_kwarg():
    """When tracker_cfg is None, 'tracker' kwarg is NOT sent (uses default)."""
    vt = _make_tracker(tracker_cfg=None)
    captured_kwargs = {}

    def spy_track(frame, **kwargs):
        captured_kwargs.update(kwargs)
        return [EmptyResult()]
    vt.model.track = spy_track

    vt.process_frame(DUMMY_FRAME)
    assert "tracker" not in captured_kwargs


# ---------------------------------------------------------------------------
# NEW — Summary statistics tests
# ---------------------------------------------------------------------------

def test_summary_empty():
    """Summary on a fresh tracker with no frames processed."""
    vt = _make_tracker()
    s = vt.get_summary()
    assert s["total_frames"] == 0
    assert s["avg_detections_per_frame"] == 0.0
    assert s["class_summary"] == {}
    assert s["total_unique_tracks"] == 0


def test_summary_counts():
    """Summary correctly counts unique track IDs per class."""
    vt = _make_tracker()
    call_count = [0]

    def fake_track(frame, **kwargs):
        call_count[0] += 1
        # Frame 1-3: person ID=1 + car ID=2
        # Frame 4-5: person ID=1 + car ID=3 (new car appears)
        car_id = 2 if call_count[0] <= 3 else 3
        return [DummyResult(MultiDummyBox(
            bboxes=[[10, 10, 30, 30], [50, 50, 80, 80]],
            confs=[0.9, 0.85],
            classes=[0, 2],         # person=0, car=2
            trk_ids=[1, car_id],
        ))]
    vt.model.track = fake_track

    for _ in range(5):
        vt.process_frame(DUMMY_FRAME)

    s = vt.get_summary()
    assert s["total_frames"] == 5
    assert s["avg_detections_per_frame"] == 2.0

    # 1 unique person (ID 1), 2 unique cars (IDs 2 and 3)
    assert s["class_summary"]["person"] == 1
    assert s["class_summary"]["car"] == 2
    assert s["total_unique_tracks"] == 3  # IDs: 1, 2, 3


def test_summary_ignores_untracked():
    """Detections with track_id=-1 are excluded from the unique ID counts."""
    vt = _make_tracker()

    def fake_track(frame, **kwargs):
        return [DummyResult(DummyBox(bbox=[10, 10, 30, 30], conf=0.5,
                                     cls=0, trk_id=None))]
    vt.model.track = fake_track

    vt.process_frame(DUMMY_FRAME)
    s = vt.get_summary()
    assert s["total_frames"] == 1
    assert s["avg_detections_per_frame"] == 1.0
    assert s["total_unique_tracks"] == 0  # -1 is not a valid track


def test_print_summary_runs(capsys):
    """print_summary produces output without crashing."""
    vt = _make_tracker()

    def fake_track(frame, **kwargs):
        return [DummyResult(DummyBox(bbox=[10, 10, 30, 30], conf=0.9,
                                     cls=0, trk_id=5))]
    vt.model.track = fake_track
    vt.process_frame(DUMMY_FRAME)

    vt.print_summary()
    captured = capsys.readouterr()
    assert "Tracking Summary" in captured.out
    assert "person" in captured.out

def test_tracker_updates_motion_analyzer():
    """Integration test: MotionAnalyzer receives updates and provides metrics."""
    vt = _make_tracker()
    # Simple eastward movement across two frames for track_id=1
    def fake_track(frame, **kwargs):
        # Alternate between two positions
        if not hasattr(fake_track, "call_idx"):
            fake_track.call_idx = 0
        x = fake_track.call_idx * 20
        fake_track.call_idx += 1
        bbox = [x, 0, x + 10, 10]
        return [DummyResult(DummyBox(bbox=bbox, conf=0.9, cls=0, trk_id=1))]
    vt.model.track = fake_track

    # Process two frames with explicit timestamps to enable speed calculation
    vt.process_frame(DUMMY_FRAME, timestamp=0.0)
    vt.process_frame(DUMMY_FRAME, timestamp=1.0)

    # History should contain two centre points
    hist = vt.motion_analyzer.get_history(1)
    assert len(hist) == 2

    speed = vt.motion_analyzer.pixel_speed(1)
    direction = vt.motion_analyzer.direction(1)
    assert speed is not None and speed > 0
    assert direction == "E"

