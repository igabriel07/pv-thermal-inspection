import os
import threading
from typing import Any, Dict, List

import cv2
import numpy as np
import torch
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from ultralytics import YOLO
from ultralytics.nn.tasks import DetectionModel
from ultralytics.nn.modules.conv import Conv
from ultralytics.nn.modules.block import C2f
from torch.nn.modules.activation import SiLU, LeakyReLU
from torch.nn.modules.batchnorm import BatchNorm2d
from torch.nn.modules.container import Sequential
from torch.nn.modules.container import ModuleList
from torch.nn.modules.conv import Conv2d
from torch.nn.modules.dropout import Dropout
from torch.nn.modules.linear import Linear
from torch.nn.modules.padding import ZeroPad2d
from torch.nn.modules.pooling import MaxPool2d
from torch.nn.modules.upsampling import Upsample

app = FastAPI(title="FastAPI + React Demo")


def _parse_cors_allow_origins(value: str | None) -> List[str]:
    if not value:
        return ["http://localhost:5173", "http://127.0.0.1:5173"]
    value = value.strip()
    if not value:
        return ["http://localhost:5173", "http://127.0.0.1:5173"]
    if value == "*":
        return ["*"]
    return [part.strip() for part in value.split(",") if part.strip()]

# Trust the local checkpoint: force weights_only=False for torch.load to avoid
# safe-global allowlisting issues with PyTorch 2.6+.
_torch_load = torch.load


def _torch_load_allow_pickle(*args, **kwargs):
    if "weights_only" not in kwargs:
        kwargs["weights_only"] = False
    return _torch_load(*args, **kwargs)


torch.load = _torch_load_allow_pickle

MODEL_PATH = os.getenv(
    "FAULT_MODEL_PATH",
    os.path.join(os.path.dirname(__file__), "models", "best.pt"),
)
IMGSZ_FAULT = int(os.getenv("FAULT_IMGSZ", "1280"))
CONF_FAULT = float(os.getenv("FAULT_CONF", "0.20"))
IOU_FAULT = float(os.getenv("FAULT_IOU", "0.50"))

_model_lock = threading.Lock()
_model: YOLO | None = None

app.add_middleware(
    CORSMiddleware,
    allow_origins=_parse_cors_allow_origins(os.getenv("CORS_ALLOW_ORIGINS")),
    allow_credentials=os.getenv("CORS_ALLOW_ORIGINS", "").strip() != "*",
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/hello")
def hello():
    return {"message": "Γεια σου! FastAPI backend is running."}


@app.on_event("startup")
def load_model_on_startup():
    try:
        get_model()
    except Exception as exc:
        print(f"[ERROR] Failed to load model: {exc}")


def get_model() -> YOLO:
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                _model = YOLO(MODEL_PATH)
    return _model


def to_gray_3ch(image: np.ndarray) -> np.ndarray:
    if image.ndim == 2:
        gray = image
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return cv2.merge([gray, gray, gray])


def detect_faults(img: np.ndarray) -> List[Dict[str, Any]]:
    model = get_model()
    result = model.predict(
        source=img,
        imgsz=IMGSZ_FAULT,
        conf=CONF_FAULT,
        iou=IOU_FAULT,
        verbose=False,
    )[0]

    if result.boxes is None or result.boxes.xyxy is None:
        return []

    xyxy = result.boxes.xyxy.cpu().numpy()
    confs = result.boxes.conf.cpu().numpy()
    classes = result.boxes.cls.cpu().numpy() if result.boxes.cls is not None else np.zeros(len(xyxy))
    h, w = img.shape[:2]

    labels: List[Dict[str, Any]] = []
    for (x1, y1, x2, y2), conf, cls in zip(xyxy, confs, classes):
        x1f, y1f, x2f, y2f = float(x1), float(y1), float(x2), float(y2)
        cx = ((x1f + x2f) / 2.0) / w
        cy = ((y1f + y2f) / 2.0) / h
        bw = (x2f - x1f) / w
        bh = (y2f - y1f) / h
        labels.append(
            {
                "classId": int(cls),
                "conf": float(conf),
                "x": round(cx, 6),
                "y": round(cy, 6),
                "w": round(bw, 6),
                "h": round(bh, 6),
            }
        )

    return labels


@app.post("/api/scan/detect")
async def scan_detect(file: UploadFile = File(...)):
    try:
        raw = await file.read()
        if not raw:
            raise HTTPException(status_code=400, detail="Empty file")

        data = np.frombuffer(raw, dtype=np.uint8)
        decoded = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if decoded is None:
            raise HTTPException(status_code=400, detail="Invalid image")

        gray = to_gray_3ch(decoded)
        labels = detect_faults(gray)
        return {
            "imageName": file.filename,
            "hasFaults": len(labels) > 0,
            "labels": labels,
        }
    except HTTPException:
        raise
    except Exception as exc:
        print(f"[ERROR] Scan failed: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))
