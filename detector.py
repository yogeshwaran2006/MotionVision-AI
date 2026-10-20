import os
from pathlib import Path
from ultralytics import YOLO

class YOLODetector:
    """Simple wrapper around Ultralytics YOLO model for basic detection.

    Parameters
    ----------
    model_path: str, default "yolov8n.pt"
        Name or path to a YOLOv8 model. The small "n" model will be downloaded
        automatically on first use.
    """
    def __init__(self, model_path: str = "yolov8n.pt"):
        self.model_path = model_path
        self.model = YOLO(self.model_path)

    def detect(self, image_path: str):
        """Run inference on a single image and return detections.
        Returns a list of dicts with ``boxes``, ``scores`` and ``classes``.
        """
        if not Path(image_path).exists():
            raise FileNotFoundError(f"Image not found: {image_path}")
        results = self.model(image_path)
        detections = []
        for r in results:
            if r.boxes:
                boxes = r.boxes.xyxy.cpu().numpy().tolist()
                scores = r.boxes.conf.cpu().numpy().tolist()
                classes = r.boxes.cls.cpu().numpy().astype(int).tolist()
            else:
                boxes, scores, classes = [], [], []
            detections.append({"boxes": boxes, "scores": scores, "classes": classes})
        return detections
