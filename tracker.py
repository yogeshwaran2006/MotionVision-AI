import os
import cv2
import time
import torch
import numpy as np
from collections import defaultdict, deque
from ultralytics import YOLO
from motion_analysis import MotionAnalyzer


class VideoTracker:
    """Video-based multi-object tracker using Ultralytics YOLO + ByteTrack/BoT-SORT.

    * Loads the YOLO model once (weights ``yolov8n.pt`` by default).
    * Detects ALL 80 COCO classes by default.  Pass ``classes=[0, 2, 5]`` to
      restrict to specific class IDs (e.g. person, car, bus).
    * Provides ``process_frame`` -> list of dicts with ``bbox``, ``score``,
      ``class_id``, ``class_name``, ``track_id``.
    * Maintains a short centre-point trail for each track (configurable length).
    * ``draw_annotations`` draws boxes, class name, confidence and track ID,
      plus the motion trail.
    * ``run_video`` handles VideoCapture, optional VideoWriter, FPS measurement,
      UI display, and prints an end-of-video summary.

    Parameters
    ----------
    model_path : str
        Name or path to a YOLO model.  Default ``"yolov8n.pt"`` (tiny, fast).
    conf : float
        Minimum confidence score for detections.  Default 0.15 — deliberately
        low so that small or partially-occluded objects are not lost.  ByteTrack
        already uses a two-stage threshold internally (``track_high_thresh`` /
        ``track_low_thresh``) to separate high- and low-confidence detections,
        so feeding more candidates to the tracker is usually beneficial.
    iou : float
        IoU threshold for NMS inside the tracker.
    trail_len : int
        Maximum number of centre points kept per track for the motion trail.
    device : str or None
        ``"cuda"`` or ``"cpu"``.  Auto-detected via ``torch.cuda.is_available()``
        when ``None``.
    classes : list[int] or None
        If given, only detect these COCO class IDs.  ``None`` means ALL classes.
    tracker_cfg : str or None
        Path to a tracker YAML config file (e.g. ``"botsort_reid.yaml"``).
        When ``None`` the Ultralytics default (ByteTrack) is used.
    """

    def __init__(self, model_path: str = "yolov8n.pt", conf: float = 0.15,
                 iou: float = 0.7, trail_len: int = 15, device: str = None,
                 classes: list = None, tracker_cfg: str = None):
        # Detect device — respect user-specified or auto-detect via torch
        if device is None:
            device = "cuda" if torch.cuda.is_available() else "cpu"
        self.device = device

        # Load YOLO model once
        self.model = YOLO(model_path)
        self.model.to(self.device)

        self.conf = conf
        self.iou = iou
        self.trail_len = trail_len
        self.classes = classes          # None → all classes
        self.tracker_cfg = tracker_cfg  # None → Ultralytics default (ByteTrack)

        # Mapping of track_id -> deque of centre points (x, y)
        self.tracks_history = defaultdict(lambda: deque(maxlen=self.trail_len))

        # Class-id -> name mapping (available immediately after model load)
        self.class_names = self.model.names

        # Motion analyzer instance
        self.motion_analyzer = MotionAnalyzer(max_history=30)

        # ----- Accumulation state for end-of-video summary -----
        # Per-class set of unique track IDs seen across all frames
        self._class_track_ids = defaultdict(set)    # class_name -> {track_id, ...}
        # Per-frame detection count (for average detections/frame)
        self._per_frame_det_counts = []

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _add_center(self, track_id: int, center: tuple):
        """Append a centre point to the bounded trail for *track_id*."""
        self.tracks_history[track_id].append(center)

    def _accumulate(self, detections: list):
        """Update the running summary counters after each frame."""
        self._per_frame_det_counts.append(len(detections))
        for det in detections:
            tid = det["track_id"]
            if tid >= 0:                             # ignore untracked (-1)
                self._class_track_ids[det["class_name"]].add(tid)

    # ------------------------------------------------------------------
    # Public API — process a single OpenCV frame (np.ndarray)
    # ------------------------------------------------------------------
    def process_frame(self, frame, frame_idx: int = None, timestamp: float = None):
        """Run detection + tracking on a single frame.

        Returns a list of dicts, one per detected object, containing:
            * ``bbox``       — [x1, y1, x2, y2] (int)
            * ``score``      — confidence (float)
            * ``class_id``   — integer class index
            * ``class_name`` — readable name from the model
            * ``track_id``   — stable tracking identifier (-1 if unassigned)
            * ``speed``      — pixels/second (optional, may be None)
            * ``direction``  — compass direction string (optional, may be None)
        The method also updates the internal trail history and summary stats.
        """
        # Build keyword args for model.track()
        # GPU: device 0 + FP16 half precision (Tensor Cores). CPU: FP32 fallback.
        use_gpu = str(self.device).startswith("cuda") and torch.cuda.is_available()
        track_kwargs = dict(
            conf=self.conf,
            iou=self.iou,
            persist=True,
            verbose=False,
            device=0 if use_gpu else "cpu",
            half=use_gpu,
        )
        if self.classes is not None:
            track_kwargs["classes"] = self.classes
        if self.tracker_cfg is not None:
            track_kwargs["tracker"] = self.tracker_cfg

        results = self.model.track(frame, **track_kwargs)

        detections = []

        # Helper to convert torch tensors or numpy arrays to plain numpy
        def _to_numpy(t):
            return t.cpu().numpy() if hasattr(t, "cpu") else t

        for r in results:
            if not r.boxes:
                continue
            boxes     = _to_numpy(r.boxes.xyxy).astype(int).tolist()
            scores    = _to_numpy(r.boxes.conf).tolist()
            class_ids = _to_numpy(r.boxes.cls).astype(int).tolist()
            track_ids = (
                _to_numpy(r.boxes.id).astype(int).tolist()
                if r.boxes.id is not None
                else [None] * len(boxes)
            )

            for bbox, score, cls_id, trk_id in zip(boxes, scores, class_ids, track_ids):
                cx = int((bbox[0] + bbox[2]) / 2)
                cy = int((bbox[1] + bbox[3]) / 2)
                if trk_id is not None:
                    self._add_center(trk_id, (cx, cy))
                # Update motion analyzer
                if trk_id is not None:
                    # Compute timestamp if not provided
                    ts = timestamp
                    if ts is None and frame_idx is not None and hasattr(self, 'fps'):
                        ts = frame_idx / self.fps
                    if ts is None:
                        ts = 0.0
                    self.motion_analyzer.update(trk_id, bbox, ts)
                    speed = self.motion_analyzer.pixel_speed(trk_id)
                    direction = self.motion_analyzer.direction(trk_id)
                else:
                    speed = None
                    direction = None
                detections.append({
                    "bbox":       bbox,
                    "score":      float(score),
                    "class_id":   int(cls_id),
                    "class_name": self.class_names.get(cls_id, f"id_{cls_id}"),
                    "track_id":   int(trk_id) if trk_id is not None else -1,
                    "speed":      speed,
                    "direction":  direction,
                })

        # Update summary stats
        self._accumulate(detections)
        return detections

    # ------------------------------------------------------------------
    # Draw bounding boxes, labels and trail on the supplied frame (in-place)
    # ------------------------------------------------------------------
    def draw_annotations(self, frame, detections, draw_fps: float = None):
        for det in detections:
            x1, y1, x2, y2 = det["bbox"]
            # Base label with class, confidence and ID
            base = f"{det['class_name']} {det['score']:.2f} ID:{det['track_id']}"
            # Append motion info if available
            if det.get('speed') is not None and det.get('direction') is not None:
                motion_part = f" | {int(det['speed'])} px/s {det['direction']}"
            else:
                motion_part = ""
            label = base + motion_part
            # Colour based on track_id for visual consistency (simple hashing)
            tid = det["track_id"]
            colour = ((tid * 37) % 256, (tid * 17) % 256, (tid * 29) % 256)
            cv2.rectangle(frame, (x1, y1), (x2, y2), colour, 2)
            cv2.putText(frame, label, (x1, max(y1 - 5, 0)),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, colour, 2)
            # Draw trail if we have history for this track
            if tid in self.tracks_history and len(self.tracks_history[tid]) >= 2:
                pts = list(self.tracks_history[tid])
                cv2.polylines(frame, [np.array(pts, dtype=np.int32)],
                              isClosed=False, color=colour, thickness=2)
        if draw_fps is not None:
            cv2.putText(frame, f"FPS: {draw_fps:.2f}", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 255, 0), 2)
        return frame

    # ------------------------------------------------------------------
    # Summary of detections & unique track IDs across the whole video
    # ------------------------------------------------------------------
    def get_summary(self):
        """Return a dict summarising what was seen across all processed frames.

        Keys:
            * ``total_frames``          — number of frames processed
            * ``avg_detections_per_frame`` — mean detections per frame
            * ``class_summary``         — dict of class_name -> unique_track_count
            * ``total_unique_tracks``   — total unique track IDs across all classes
        """
        total_frames = len(self._per_frame_det_counts)
        avg_dets = (
            sum(self._per_frame_det_counts) / total_frames
            if total_frames > 0 else 0.0
        )
        class_summary = {
            cls: len(ids) for cls, ids in sorted(self._class_track_ids.items())
        }
        total_unique = len(set().union(*self._class_track_ids.values())) if self._class_track_ids else 0
        return {
            "total_frames": total_frames,
            "avg_detections_per_frame": round(avg_dets, 2),
            "class_summary": class_summary,
            "total_unique_tracks": total_unique,
        }

    def print_summary(self):
        """Pretty-print the end-of-video detection & tracking summary."""
        s = self.get_summary()
        print("\n" + "=" * 60)
        print("  MotionVision AI — Tracking Summary")
        print("=" * 60)
        print(f"  Frames processed          : {s['total_frames']}")
        print(f"  Avg detections / frame    : {s['avg_detections_per_frame']}")
        print(f"  Total unique tracked IDs  : {s['total_unique_tracks']}")
        print("-" * 60)
        print(f"  {'Class':<20} {'Unique Track IDs':>18}")
        print("-" * 60)
        for cls_name, count in s["class_summary"].items():
            print(f"  {cls_name:<20} {count:>18}")
        print("=" * 60 + "\n")

    # ------------------------------------------------------------------
    # High-level driver for a video file (or webcam if path is 0)
    # ------------------------------------------------------------------
    def run_video(self, input_path, output_path: str = None,
                  show: bool = True, max_frames: int = None):
        """Process a video end-to-end: detect, track, annotate, display/save.

        Parameters
        ----------
        input_path : str or int
            File path or ``0`` for webcam.
        output_path : str or None
            If given, write annotated video to this path.
        show : bool
            Display a live OpenCV window (press *q* to quit early).
        max_frames : int or None
            Stop after this many frames (useful for quick benchmarks).

        Returns
        -------
        float
            Average processing FPS.
        """
        cap = cv2.VideoCapture(input_path)
        if not cap.isOpened():
            raise RuntimeError(f"Cannot open video source: {input_path}")

        width  = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        fps_input = cap.get(cv2.CAP_PROP_FPS) or 30.0

        writer = None
        if output_path:
            fourcc = cv2.VideoWriter_fourcc(*"mp4v")
            writer = cv2.VideoWriter(output_path, fourcc, fps_input,
                                     (width, height))

        frame_count = 0
        start_time = time.time()

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            detections = self.process_frame(frame)
            frame_count += 1

            # Instantaneous FPS from elapsed time
            elapsed = time.time() - start_time
            fps = frame_count / elapsed if elapsed > 0 else 0.0

            annotated = self.draw_annotations(frame, detections, draw_fps=fps)

            if writer:
                writer.write(annotated)
            if show:
                cv2.imshow("MotionVision Tracking", annotated)
                if cv2.waitKey(1) & 0xFF == ord("q"):
                    break
            if max_frames and frame_count >= max_frames:
                break

        # --- Clean up resources ---
        cap.release()
        if writer:
            writer.release()
        if show:
            cv2.destroyAllWindows()

        total_time = time.time() - start_time
        avg_fps = frame_count / total_time if total_time > 0 else 0.0

        print(f"\nProcessed {frame_count} frames in {total_time:.2f}s "
              f"— average FPS: {avg_fps:.2f}")

        # Print the detection/tracking summary
        self.print_summary()

        return avg_fps
