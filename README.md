# MotionVision AI — Real-Time Multi-Object Detection, Tracking, Motion Analytics & Hand Arithmetic

![Python](https://img.shields.io/badge/Python-3.14-blue)
![PyTorch](https://img.shields.io/badge/PyTorch-CUDA%20FP16-ee4c2c)
![YOLO11x](https://img.shields.io/badge/Ultralytics-YOLO11x-00FFFF)
![MediaPipe](https://img.shields.io/badge/MediaPipe-Hands-00A86B)
![Tests](https://img.shields.io/badge/Tests-22%2F22%20Passing-brightgreen)
![License](https://img.shields.io/badge/License-MIT-green)

**Academic Institution:** SIMATS Engineering — Saveetha Institute of Medical and Technical Sciences, Chennai  
**Course:** Machine Learning / Computer Vision Capstone Project  
**Team Members:**  
- **Yogeshwaran B (192412540)** — *Module 1: Deep Learning Detection & Multi-Object Tracking Engine*  
- **Harine M (192412115)** — *Module 2: Kinematic Motion Analytics, Hand Landmark Arithmetic & Live Full-Screen HUD*  

---

## 1. Project Overview
**MotionVision AI** is an end-to-end, GPU-accelerated computer vision platform designed for real-time visual perception, kinematic analysis, and touchless interaction. The system integrates:
1. **80-Class Object Detection:** Powered by Ultralytics **YOLO11x** (56.9M parameters) using FP16 Tensor Core acceleration to eliminate small-object misclassifications (such as phones being mistaken for suitcases).
2. **Persistent Multi-Object Tracking:** Temporal identity association using **ByteTrack** (Kalman filtering + two-stage IoU) and **BoT-SORT** (deep appearance ReID).
3. **Kinematic Motion Analytics:** Centroid extraction, Exponential Moving Average (EMA) trajectory smoothing, Euclidean speed computation (`px/s`), and 8-way compass heading classification with stationary jitter suppression.
4. **Interactive Hand & Finger Arithmetic:** Google **MediaPipe Hands** (21 3D landmarks) tracking Left and Right palms, counting raised fingers (0–5 per hand), and displaying real-time simultaneous four-operation arithmetic (`+`, `-`, `*`, `/`) on a mirrored full-screen OpenCV HUD.

---

## 2. Two-Module System Architecture

```text
┌────────────────────────────────────────────────────────────────────────────┐
│              INPUT LAYER: 1080p USB SMARTPHONE CAMERA / VIDEO              │
│            OpenCV DirectShow (CAP_DSHOW) + Horizontal Mirror Flip          │
└─────────────────────────────────────┬──────────────────────────────────────┘
                                      │
                                      ▼
┌────────────────────────────────────────────────────────────────────────────┐
│    MODULE 1: DEEP LEARNING DETECTION & TRACKING ENGINE (Yogeshwaran B)     │
│  • YOLO11x Detector (56.9M Params, C3k2 + C2PSA Attention, FP16 Tensor)    │
│  • 80-Class COCO Object Detection (Conf >= 0.35, IoU = 0.70)               │
│  • Multi-Object Trackers: ByteTrack (bytetrack.yaml) & BoT-SORT (ReID)     │
│  • Bounded Track History Management (Per-ID Deque Queues)                  │
└─────────────────────────────────────┬──────────────────────────────────────┘
                                      │
                                      ▼
┌────────────────────────────────────────────────────────────────────────────┐
│    MODULE 2: KINEMATIC MOTION ANALYTICS & HAND ARITHMETIC (Harine M)       │
│  • MotionAnalyzer: Centroid Extraction + EMA Position Smoothing (alpha=0.7)│
│  • Kinematics: Time-Normalized Pixel Speed (px/s) + 8-Way Compass Heading  │
│  • MediaPipe Hands: 21 3D Landmarks, Left/Right Palm & Finger Counter (0-5)│
│  • Live Arithmetic Engine: Simultaneous [+], [-], [x], [/] HUD Banner      │
│  • Native OpenCV Full-Screen Annotated Display (Zero-Disk Clutter)         │
└────────────────────────────────────────────────────────────────────────────┘
