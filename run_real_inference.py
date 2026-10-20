import os
import json
import cv2
from detector import YOLODetector

# Path to real‑world image (downloaded in previous step)
image_path = 'sample_real.jpg'
if not os.path.exists(image_path):
    raise FileNotFoundError(f"{image_path} not found. Please provide a real image.")

# Initialise detector (will load yolov8n.pt if not already cached)
detector = YOLODetector()

# Run inference using the underlying Ultralytics model to keep the Result objects
results = detector.model(image_path)

# Get class‑name mapping from the model
class_names = detector.model.names  # dict: id -> name

# Prepare a printable list of detections
all_detections = []
for r in results:
    if r.boxes:
        boxes = r.boxes.xyxy.cpu().numpy().tolist()
        scores = r.boxes.conf.cpu().numpy().tolist()
        classes = r.boxes.cls.cpu().numpy().astype(int).tolist()
    else:
        boxes, scores, classes = [], [], []
    for box, score, cls_id in zip(boxes, scores, classes):
        name = class_names.get(cls_id, f"id_{cls_id}")
        all_detections.append({
            "class_id": cls_id,
            "class_name": name,
            "score": float(score),
            "box": box
        })

print("Detections (class, score, box):")
print(json.dumps(all_detections, indent=2))

# Create an annotated image using the Result.plot() helper (draws boxes & labels)
# results[0] is the first (and only) image result
annotated_img = results[0].plot()
annotated_path = 'sample_real_annotated.jpg'
cv2.imwrite(annotated_path, annotated_img)
print(f"Annotated image saved to: {os.path.abspath(annotated_path)}")
