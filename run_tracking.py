"""MotionVision AI — Run YOLO + tracker on a video file (or webcam).

Usage examples
--------------
# Default (yolo11x FP16 on GPU, ByteTrack, all classes, conf=0.35):
  python run_tracking.py sample_video.mp4

# Save annotated output & limit frames:
  python run_tracking.py sample_video.mp4 --output tracked.mp4 --max_frames 200

# Use BoT-SORT with ReID instead of ByteTrack:
  python run_tracking.py sample_video.mp4 --tracker botsort_reid.yaml

# Live webcam: mirrored, fullscreen, YOLO + hand/finger arithmetic
# (q / ESC quit, f toggle fullscreen):
  python run_tracking.py 0 --tracker bytetrack.yaml

# Filter to people and cars only:
  python run_tracking.py sample_video.mp4 --classes 0 2

# Use a faster (less accurate) model:
  python run_tracking.py sample_video.mp4 --model yolov8s.pt

# Headless mode (no display window, just save output):
  python run_tracking.py sample_video.mp4 --output tracked.mp4 --no-show
"""

import argparse
import os
import time
import cv2
from tracker import VideoTracker
from hand_arithmetic import HandArithmetic

WINDOW_NAME = "MotionVision AI - Live"
ESC_KEY = 27


def main():
    parser = argparse.ArgumentParser(
        description="MotionVision AI — YOLO + ByteTrack/BoT-SORT video tracking.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )

    # ----- Input / Output -----
    parser.add_argument("input",
                        help='Path to input video file, or camera index (e.g. "0") for webcam.')
    parser.add_argument("--output",
                        help="Path to save annotated output video (optional).",
                        default=None)

    # ----- Model & detection -----
    parser.add_argument("--model", default="yolo11x.pt",
                        help="YOLO model weights (default: yolo11x.pt, auto-downloaded). "
                             "Use yolov8s.pt / yolov8n.pt for more speed.")
    parser.add_argument("--conf", type=float, default=0.35,
                        help="Detection confidence threshold (default: 0.35). "
                             "Lower catches more objects but adds noise.")
    parser.add_argument("--iou", type=float, default=0.7,
                        help="IoU threshold for NMS (default: 0.7).")
    parser.add_argument("--classes", type=int, nargs="+", default=None,
                        help="COCO class IDs to detect (default: all 80). "
                             "Example: --classes 0 2 5  (person, car, bus).")

    # ----- Tracker -----
    parser.add_argument("--tracker", default=None,
                        help="Path to a tracker YAML config. "
                             "Default: ByteTrack (Ultralytics built-in). "
                             "Use botsort_reid.yaml for BoT-SORT with ReID.")

    # ----- Display & limits -----
    parser.add_argument("--trail", type=int, default=15,
                        help="Length of motion trail in frames (default: 15).")
    parser.add_argument("--max_frames", type=int, default=None,
                        help="Stop after this many frames (for quick tests).")
    parser.add_argument("--no-show", action="store_true",
                        help="Disable the live display window (headless mode).")
    parser.add_argument("--hands", action="store_true",
                        help="Enable hand/finger arithmetic (on by default for webcam).")
    parser.add_argument("--no-hands", action="store_true",
                        help="Disable hand/finger arithmetic for webcam input.")

    args = parser.parse_args()

    # Resolve input — camera index (int) or file path (str)
    try:
        source_idx = int(args.input)
        is_webcam = True
    except ValueError:
        is_webcam = False

    if is_webcam:
        cap = cv2.VideoCapture(source_idx, cv2.CAP_DSHOW)
        if not cap.isOpened():
            cap = cv2.VideoCapture(source_idx)
    else:
        if not os.path.exists(args.input):
            raise FileNotFoundError(f"Video file not found: {args.input}")
        cap = cv2.VideoCapture(args.input)

    if not cap.isOpened():
        raise RuntimeError(f"Cannot open video source: {args.input}")

    # Webcam: request 1280x720 (must be set before reading the actual size)
    if is_webcam:
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

    # Output video writer (optional)
    writer = None
    if args.output:
        width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        fps_input = cap.get(cv2.CAP_PROP_FPS) or 30.0
        fourcc = cv2.VideoWriter_fourcc(*"mp4v")
        writer = cv2.VideoWriter(args.output, fourcc, fps_input, (width, height))

    # Hand / finger arithmetic (default on for webcam, or via --hands)
    use_hands = args.hands or (is_webcam and not args.no_hands)
    hands = None
    if use_hands:
        hands = HandArithmetic(
            max_hands=2,
            min_detection_confidence=0.7,
            min_tracking_confidence=0.6,
            mirrored=is_webcam,   # webcam frames are flipped below
        )
        if not hands.available:
            print(f"[warning] Hand arithmetic disabled: {hands.unavailable_reason}")
            hands = None

    # --- Print configuration ---
    tracker_label = args.tracker if args.tracker else "ByteTrack (default)"
    classes_label = (
        ", ".join(str(c) for c in args.classes) if args.classes
        else "all 80 COCO classes"
    )
    print("=" * 60)
    print("  MotionVision AI — Tracking Configuration")
    print("=" * 60)
    print(f"  Source          : {'Webcam index ' + str(source_idx) if is_webcam else args.input}")
    print(f"  Model           : {args.model}")
    print(f"  Tracker         : {tracker_label}")
    print(f"  Confidence      : {args.conf}")
    print(f"  IoU             : {args.iou}")
    print(f"  Classes         : {classes_label}")
    print(f"  Trail length    : {args.trail}")
    print(f"  Max frames      : {args.max_frames or 'unlimited'}")
    print(f"  Output          : {args.output or 'none (live display only)'}")
    print(f"  Mirrored        : {'yes' if is_webcam else 'no'}")
    print(f"  Hand arithmetic : {'on' if hands else 'off'}")
    print("  Keys            : q / ESC quit | f toggle fullscreen")
    print("=" * 60 + "\n")

    tracker = VideoTracker(
        model_path=args.model,
        conf=args.conf,
        iou=args.iou,
        trail_len=args.trail,
        classes=args.classes,
        tracker_cfg=args.tracker,
    )

    frame_count = 0
    start_time = time.time()
    fullscreen = is_webcam

    if not args.no_show:
        cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
        if fullscreen:
            cv2.setWindowProperty(WINDOW_NAME, cv2.WND_PROP_FULLSCREEN,
                                  cv2.WINDOW_FULLSCREEN)

    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                break
            if is_webcam:
                frame = cv2.flip(frame, 1)   # selfie / mirror view

            # --- Pass 1: YOLOv8 object tracking (phone, bottle, person, ...) ---
            elapsed = time.time() - start_time
            detections = tracker.process_frame(frame, frame_idx=frame_count, timestamp=elapsed)
            frame_count += 1

            # --- Pass 2: hand landmarks (read before any overlay is drawn) ---
            hand_infos = hands.process_frame(frame) if hands else []

            # Boxes, IDs, trails and per-object motion (px/s + direction)
            annotated_frame = tracker.draw_annotations(frame, detections)

            # Hand skeletons, palm boxes + arithmetic banner
            if hands:
                hands.draw(annotated_frame, hand_infos)

            # --- Pass 3: global stats (drawn last so the banner never hides it) ---
            fps = frame_count / elapsed if elapsed > 0 else 0.0
            stats = f"FPS: {fps:.1f} | Objects: {len(detections)}"
            if hands:
                stats += f" | Hands: {len(hand_infos)}"
            h = annotated_frame.shape[0]
            cv2.putText(annotated_frame, stats, (10, h - 15),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 4)
            cv2.putText(annotated_frame, stats, (10, h - 15),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)

            if writer is not None:
                writer.write(annotated_frame)

            if not args.no_show:
                cv2.imshow(WINDOW_NAME, annotated_frame)
                key = cv2.waitKey(1) & 0xFF
                if key in (ord('q'), ESC_KEY):
                    break
                elif key == ord('f'):
                    fullscreen = not fullscreen
                    cv2.setWindowProperty(
                        WINDOW_NAME, cv2.WND_PROP_FULLSCREEN,
                        cv2.WINDOW_FULLSCREEN if fullscreen else cv2.WINDOW_NORMAL)

            if args.max_frames and frame_count >= args.max_frames:
                break
    finally:
        cap.release()
        if writer is not None:
            writer.release()
        if hands:
            hands.close()
        cv2.destroyAllWindows()

    total_time = time.time() - start_time
    avg_fps = frame_count / total_time if total_time > 0 else 0.0

    print(f"\nProcessed {frame_count} frames in {total_time:.2f}s "
          f"— average FPS: {avg_fps:.2f}")

    tracker.print_summary()


if __name__ == "__main__":
    main()
