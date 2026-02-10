from pathlib import Path
import cv2
import numpy as np
from ultralytics import YOLO

# ================== PATHS ==================
FAULT_MODEL_A = r"F:\toposol_1.0\fine_tuning_1.0\weights\best.pt"

SOURCE = r"F:\toposol_1.0\original-images-yes-no\park_4\tmp"
OUT_DIR = r"F:\toposol_1.0\outputs_faults_only_1"
# ==========================================

# ============== INFERENCE PARAMS ==========
IMGSZ_FAULT = 1280
CONF_FAULT = 0.20
IOU_FAULT = 0.50
# ==========================================

IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}

# Κόκκινο faults (BGR)
FAULT_COLOR = (0, 0, 255)


def list_images(src: Path):
    if src.is_file():
        return [src]
    imgs = []
    for p in src.rglob("*"):
        if p.is_file() and p.suffix.lower() in IMG_EXTS:
            imgs.append(p)
    return sorted(imgs)


def detect_faults(model: YOLO, img_path: Path):
    r = model.predict(
        source=str(img_path),
        imgsz=IMGSZ_FAULT,
        conf=CONF_FAULT,
        iou=IOU_FAULT,
        verbose=False
    )[0]

    out = []
    if r.boxes is None or r.boxes.xyxy is None:
        return out

    xyxy = r.boxes.xyxy.cpu().numpy()
    confs = r.boxes.conf.cpu().numpy()

    for (x1, y1, x2, y2), cf in zip(xyxy, confs):
        out.append({"xyxy": [float(x1), float(y1), float(x2), float(y2)], "conf": float(cf)})
    return out


def draw_fault_boxes(img_bgr, boxes, color=FAULT_COLOR):
    out = img_bgr.copy()
    for b in boxes:
        x1, y1, x2, y2 = map(int, b["xyxy"])
        cv2.rectangle(out, (x1, y1), (x2, y2), color, 2)
    return out


def main():
    src = Path(SOURCE)
    images = list_images(src)
    if not images:
        raise RuntimeError(f"No images found in: {src}")

    out_root = Path(OUT_DIR)
    out_root.mkdir(parents=True, exist_ok=True)

    fault_a = YOLO(FAULT_MODEL_A)

    print(f"Found {len(images)} images. Saving outputs to: {out_root}")

    for img_path in images:
        img = cv2.imread(str(img_path), cv2.IMREAD_COLOR)
        if img is None:
            print(f"[WARN] Cannot read: {img_path}")
            continue

        boxes_a = detect_faults(fault_a, img_path)

        out_img = draw_fault_boxes(img, boxes_a)
        cv2.imwrite(str(out_root / f"{img_path.stem}.jpg"), out_img)

        print(f"{img_path.name}: detections={len(boxes_a)}")

    print("\n✔ Done.")
    print("Output folder:", out_root.resolve())


if __name__ == "__main__":
    main()
