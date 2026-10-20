import json, os
from detector import YOLODetector

image_path = 'synthetic.png'
if not os.path.exists(image_path):
    raise FileNotFoundError(f"{image_path} not found")

detector = YOLODetector()
results = detector.detect(image_path)
print('Detections:', json.dumps(results, indent=2))
