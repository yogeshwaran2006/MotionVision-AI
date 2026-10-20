"""MotionVision AI — Live hand / finger arithmetic.

Detects up to two hands with MediaPipe, counts raised fingers (0-5) per hand and
shows a top banner with all four operations (+, -, x, /) between the Left (L)
and Right (R) hand counts simultaneously.

Backends (picked automatically)
-------------------------------
1. ``mediapipe.solutions.hands``  — legacy API (mediapipe < 0.10.2x).
2. ``mediapipe.tasks`` HandLandmarker — used when ``mp.solutions`` was removed
   (newer mediapipe wheels). The ``hand_landmarker.task`` model is downloaded
   once into ``./models/``.

If mediapipe is missing (or the model cannot be fetched) the analyzer reports
``available == False`` and every method becomes a harmless no-op, so the rest of
the tracking pipeline keeps working.

Typical use
-----------
    hands = HandArithmeticAnalyzer(mirrored=True)   # selfie-flipped frames
    infos = hands.process(frame)                    # BGR frame in
    hands.draw(frame, infos)                        # skeletons, boxes, banner
"""

from __future__ import annotations

import os
import time
import urllib.request
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

import cv2

MODEL_URL = ("https://storage.googleapis.com/mediapipe-models/hand_landmarker/"
             "hand_landmarker/float16/1/hand_landmarker.task")
MODEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "models", "hand_landmarker.task")

# Landmark indices (MediaPipe hand model)
WRIST = 0
THUMB_MCP, THUMB_IP, THUMB_TIP = 2, 3, 4
INDEX_MCP, PINKY_MCP = 5, 17
# (tip, pip) pairs for index, middle, ring, pinky
FINGER_TIP_PIP = [(8, 6), (12, 10), (16, 14), (20, 18)]

HAND_CONNECTIONS = [
    (0, 1), (1, 2), (2, 3), (3, 4),
    (0, 5), (5, 6), (6, 7), (7, 8),
    (5, 9), (9, 10), (10, 11), (11, 12),
    (9, 13), (13, 14), (14, 15), (15, 16),
    (13, 17), (17, 18), (18, 19), (19, 20),
    (0, 17),
]

OPERATORS = {
    "+": "Sum",
    "-": "Diff",
    "*": "Prod",
    "/": "Quot",
}


@dataclass
class HandInfo:
    """One detected hand."""
    label: str                                   # "Left" / "Right" (anatomical)
    fingers: int                                 # 0-5
    bbox: Tuple[int, int, int, int]              # x1, y1, x2, y2 (pixels)
    points: List[Tuple[int, int]] = field(default_factory=list)  # 21 px points


def _dist(a, b) -> float:
    return ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5


def count_fingers(pts: List[Tuple[float, float]], label: str,
                  mirrored: bool = True) -> int:
    """Count raised fingers (0-5) from 21 hand landmarks (normalised x, y).

    * Index / middle / ring / pinky: raised when ``tip.y < pip.y`` (image y grows
      downward, so the tip is above its PIP joint). Assumes an upright hand.
    * Thumb: it folds sideways, so compare ``tip.x`` with ``ip.x``. On a
      mirrored (selfie) frame with the palm facing the camera, an open Right
      thumb points to image-left (``tip.x < ip.x``) and an open Left thumb to
      image-right (``tip.x > ip.x``). For un-mirrored frames the rule flips.
    """
    count = 0
    for tip, pip in FINGER_TIP_PIP:
        if pts[tip][1] < pts[pip][1]:
            count += 1

    thumb_opens_left = (label == "Right") == mirrored
    tip_x, ip_x = pts[THUMB_TIP][0], pts[THUMB_IP][0]
    if (tip_x < ip_x) if thumb_opens_left else (tip_x > ip_x):
        count += 1
    return count


def compute(op: str, x: int, y: int) -> Optional[float]:
    """Apply ``op`` to (x, y). Returns None for division by zero."""
    if op == "+":
        return x + y
    if op == "-":
        return x - y
    if op == "*":
        return x * y
    if op == "/":
        return None if y == 0 else x / y
    raise ValueError(f"Unknown operator: {op}")


def _fmt(value: Optional[float]) -> str:
    if value is None:
        return "undefined"
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.2f}"


class HandArithmeticAnalyzer:
    """Detect hands, count fingers and render an arithmetic HUD."""

    def __init__(self, max_hands: int = 2, min_detection_confidence: float = 0.7,
                 min_tracking_confidence: float = 0.6,
                 mirrored: bool = True, model_path: str = MODEL_PATH):
        """
        Parameters
        ----------
        mirrored : bool
            Set True if frames passed to :meth:`process` are selfie-flipped.
            MediaPipe assumes mirrored input when assigning Left/Right, so for
            raw (un-flipped) webcam frames the labels are swapped internally.
        """
        self.max_hands = max_hands
        self.min_det = min_detection_confidence
        self.min_track = min_tracking_confidence
        self.mirrored = mirrored
        self.model_path = model_path
        self.operator = "+"

        self.available = False
        self.unavailable_reason = ""
        self._backend = None        # "solutions" | "tasks"
        self._hands = None          # legacy solver
        self._landmarker = None     # tasks landmarker
        self._mp = None
        self._last_ts_ms = 0
        self._t0 = time.monotonic()
        self._init_backend()

    # ------------------------------------------------------------------ setup
    def _init_backend(self) -> None:
        try:
            import mediapipe as mp
        except Exception as exc:                       # not installed / broken
            self.unavailable_reason = f"mediapipe not available ({exc})"
            return
        self._mp = mp

        # 1) Legacy solutions API
        try:
            solutions = getattr(mp, "solutions", None)
            if solutions is not None and hasattr(solutions, "hands"):
                self._hands = solutions.hands.Hands(
                    static_image_mode=False,
                    max_num_hands=self.max_hands,
                    min_detection_confidence=self.min_det,
                    min_tracking_confidence=self.min_track,
                )
                self._backend = "solutions"
                self.available = True
                return
        except Exception as exc:
            self.unavailable_reason = f"mp.solutions.hands failed ({exc})"

        # 2) Tasks API fallback
        try:
            from mediapipe.tasks import python as mp_python
            from mediapipe.tasks.python import vision

            if not os.path.exists(self.model_path):
                os.makedirs(os.path.dirname(self.model_path), exist_ok=True)
                print(f"[hands] Downloading hand_landmarker model to {self.model_path} ...")
                urllib.request.urlretrieve(MODEL_URL, self.model_path)

            options = vision.HandLandmarkerOptions(
                base_options=mp_python.BaseOptions(model_asset_path=self.model_path),
                running_mode=vision.RunningMode.VIDEO,
                num_hands=self.max_hands,
                min_hand_detection_confidence=self.min_det,
                min_hand_presence_confidence=0.5,
                min_tracking_confidence=self.min_track,
            )
            self._landmarker = vision.HandLandmarker.create_from_options(options)
            self._backend = "tasks"
            self.available = True
            self.unavailable_reason = ""
        except Exception as exc:
            self.unavailable_reason = f"hand landmarker unavailable ({exc})"
            self.available = False

    # ------------------------------------------------------------------ input
    def set_operator(self, op: str) -> bool:
        """Switch the primary operator. Returns True if ``op`` is valid."""
        if op in OPERATORS:
            self.operator = op
            return True
        return False

    def _anatomical_label(self, raw: str) -> str:
        if self.mirrored:
            return raw
        return "Left" if raw == "Right" else "Right"

    # ---------------------------------------------------------------- process
    def process_frame(self, frame_bgr) -> List[HandInfo]:
        """Same as :meth:`process` (mirrors ``VideoTracker.process_frame``)."""
        return self.process(frame_bgr)

    def process(self, frame_bgr) -> List[HandInfo]:
        """Detect hands in a BGR frame and return a list of :class:`HandInfo`."""
        if not self.available:
            return []
        h, w = frame_bgr.shape[:2]
        rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)

        raw_hands = []   # list of (raw_label, [(x_norm, y_norm) * 21])
        try:
            if self._backend == "solutions":
                res = self._hands.process(rgb)
                if res.multi_hand_landmarks:
                    for lm, hd in zip(res.multi_hand_landmarks, res.multi_handedness):
                        raw_hands.append((hd.classification[0].label,
                                          [(p.x, p.y) for p in lm.landmark]))
            else:
                mp = self._mp
                image = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)
                ts = int((time.monotonic() - self._t0) * 1000)
                ts = max(ts, self._last_ts_ms + 1)   # must be strictly increasing
                self._last_ts_ms = ts
                res = self._landmarker.detect_for_video(image, ts)
                for lms, handed in zip(res.hand_landmarks, res.handedness):
                    raw_hands.append((handed[0].category_name,
                                      [(p.x, p.y) for p in lms]))
        except Exception as exc:
            print(f"[hands] detection error, disabling hand mode: {exc}")
            self.available = False
            self.unavailable_reason = str(exc)
            return []

        infos: List[HandInfo] = []
        for raw_label, norm in raw_hands:
            label = self._anatomical_label(raw_label)
            fingers = count_fingers(norm, label, mirrored=self.mirrored)
            px = [(int(x * w), int(y * h)) for x, y in norm]
            xs = [p[0] for p in px]
            ys = [p[1] for p in px]
            pad = 15
            bbox = (max(min(xs) - pad, 0), max(min(ys) - pad, 0),
                    min(max(xs) + pad, w - 1), min(max(ys) + pad, h - 1))
            infos.append(HandInfo(label=label, fingers=fingers, bbox=bbox, points=px))

        # Two hands reported with the same label -> disambiguate by position.
        if len(infos) == 2 and infos[0].label == infos[1].label:
            infos.sort(key=lambda i: i.bbox[0])
            left_x_first = self.mirrored   # mirrored: true Left is on image-left
            infos[0].label = "Left" if left_x_first else "Right"
            infos[1].label = "Right" if left_x_first else "Left"
        return infos

    # ------------------------------------------------------------------- draw
    def draw(self, frame, hands: List[HandInfo]):
        """Draw landmarks, palm boxes and the arithmetic HUD onto ``frame``."""
        if not self.available:
            return frame

        green = (0, 255, 0)
        font = cv2.FONT_HERSHEY_SIMPLEX
        for hand in hands:
            for a, b in HAND_CONNECTIONS:
                cv2.line(frame, hand.points[a], hand.points[b], green, 2)
            for p in hand.points:
                cv2.circle(frame, p, 4, (255, 255, 255), -1)
                cv2.circle(frame, p, 4, green, 1)
            x1, y1, x2, y2 = hand.bbox
            cv2.rectangle(frame, (x1, y1), (x2, y2), green, 2)

            # "Left: N" / "Right: N" tag above the palm
            text = f"{hand.label}: {hand.fingers}"
            (tw, th), _ = cv2.getTextSize(text, font, 0.8, 2)
            ty = max(y1 - 10, th + 10)
            cv2.rectangle(frame, (x1, ty - th - 8), (x1 + tw + 10, ty + 6), green, -1)
            cv2.putText(frame, text, (x1 + 5, ty), font, 0.8, (0, 0, 0), 2)

        if len(hands) == 1:
            h = hands[0]
            self._draw_banner(frame, f"Palm Detected: {h.label} (Number: {h.fingers})")
        elif len(hands) >= 2:
            by_label = {h.label: h.fingers for h in hands}
            L = by_label.get("Left", hands[0].fingers)
            R = by_label.get("Right", hands[1].fingers)
            self._draw_banner(frame, self.arithmetic_text(L, R))
        return frame

    @staticmethod
    def arithmetic_text(L: int, R: int) -> str:
        """All four operations on (L, R) as a single banner string."""
        div = f"{L / R:.2f}" if R != 0 else "N/A"
        return (f"Left: {L} | Right: {R}   ==>   [+] {L + R}   |   [-] {L - R}   |   "
                f"[x] {L * R}   |   [/] {div}")

    @staticmethod
    def _draw_banner(frame, text: str) -> None:
        """High-contrast full-width banner across the top of the frame."""
        font = cv2.FONT_HERSHEY_SIMPLEX
        fw = frame.shape[1]
        # Shrink the font until the line fits the frame width
        scale, thick = 1.0, 2
        (tw, th), _ = cv2.getTextSize(text, font, scale, thick)
        while tw > fw - 40 and scale > 0.4:
            scale -= 0.05
            (tw, th), _ = cv2.getTextSize(text, font, scale, thick)

        banner_h = th + 30
        overlay = frame.copy()
        cv2.rectangle(overlay, (0, 0), (fw, banner_h), (0, 0, 0), -1)
        cv2.addWeighted(overlay, 0.8, frame, 0.2, 0, frame)
        cv2.line(frame, (0, banner_h), (fw, banner_h), (0, 255, 0), 2)
        cv2.putText(frame, text, ((fw - tw) // 2, (banner_h + th) // 2),
                    font, scale, (0, 255, 255), thick)

    # ---------------------------------------------------------------- cleanup
    def close(self) -> None:
        for obj in (self._hands, self._landmarker):
            try:
                if obj is not None:
                    obj.close()
            except Exception:
                pass
        self._hands = None
        self._landmarker = None
        self.available = False


# Short alias used by run_tracking.py
HandArithmetic = HandArithmeticAnalyzer
