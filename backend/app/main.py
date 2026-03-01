import base64
import json
import os
import shutil
import subprocess
import sys
import threading
import tempfile
from pathlib import Path
from typing import Any, Dict, List

# Ensure the project root is on sys.path so absolute imports like `backend.*`
# work regardless of the current working directory used to launch uvicorn.
_ROOT = Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import cv2
import numpy as np
import torch
from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
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


def create_app() -> FastAPI:
    app = FastAPI(title="Thermal Fault Detection API")

    app.add_middleware(
        CORSMiddleware,
        allow_origins=_parse_cors_allow_origins(os.getenv("CORS_ALLOW_ORIGINS")),
        allow_credentials=os.getenv("CORS_ALLOW_ORIGINS", "").strip() != "*",
        allow_methods=["*"],
        allow_headers=["*"],
    )

    @app.get("/api/health")
    def health():
        versions: Dict[str, str] = {
            "python": sys.version.split()[0],
        }

        try:
            import fastapi

            v = getattr(fastapi, "__version__", "")
            if isinstance(v, str) and v:
                versions["fastapi"] = v
        except Exception:
            pass

        for name, mod in (
            ("numpy", np),
            ("opencv", cv2),
            ("torch", torch),
        ):
            v = getattr(mod, "__version__", "")
            if isinstance(v, str) and v:
                versions[name] = v

        try:
            import ultralytics

            v = getattr(ultralytics, "__version__", "")
            if isinstance(v, str) and v:
                versions["ultralytics"] = v
        except Exception:
            pass

        return {"status": "ok", "message": "FastAPI backend is running.", "versions": versions}

    @app.on_event("startup")
    def load_model_on_startup():
        # For faster local development, allow skipping the heavy model preload.
        # Set `PRELOAD_MODEL=0` to avoid loading YOLO at startup.
        preload = os.getenv("PRELOAD_MODEL", "1").strip().lower() not in {"0", "false", "no", "off"}
        if not preload:
            print("[INFO] PRELOAD_MODEL=0; skipping model preload.")
            return
        try:
            get_model()
            # Best-effort: also preload the type model if present.
            get_type_model()
        except Exception as exc:
            print(f"[ERROR] Failed to load model: {exc}")

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

    async def _probe_tiff_metadata(file: UploadFile, qr: bool) -> Dict[str, Any]:
        raw = await file.read()
        if not raw:
            raise HTTPException(status_code=400, detail="Empty file")

        filename = file.filename or "uploaded.tiff"
        lowered = filename.lower()
        if not (lowered.endswith(".tif") or lowered.endswith(".tiff")):
            raise HTTPException(status_code=400, detail="File must be .tif or .tiff")

        from backend.check_for_metadata_tiff import (
            google_maps_link,
            make_qr_png,
            probe_tiff,
            to_json_dict,
        )  # local import (heavy deps)

        suffix = ".tiff" if lowered.endswith(".tiff") else ".tif"
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(raw)
                tmp_path = tmp.name

            result = await run_in_threadpool(probe_tiff, tmp_path)
            report: Dict[str, Any] = to_json_dict(result)
            report["file"] = filename

            if qr:
                geo: Dict[str, Any] = {}
                if isinstance(getattr(result, "categories", None), dict):
                    geo = result.categories.get("geolocation", {}) or {}
                lat = geo.get("latitude") if isinstance(geo, dict) else None
                lon = geo.get("longitude") if isinstance(geo, dict) else None

                maps_link = None
                qr_b64 = None
                qr_err = None
                tmp_png_path = None

                try:
                    if lat is not None and lon is not None:
                        maps_link = google_maps_link(float(lat), float(lon))
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_png:
                            tmp_png_path = tmp_png.name
                        saved_path, qr_err = make_qr_png(maps_link, tmp_png_path)
                        if saved_path:
                            with open(saved_path, "rb") as f:
                                qr_b64 = base64.b64encode(f.read()).decode("ascii")
                    else:
                        qr_err = "No latitude/longitude found; QR not generated."
                except Exception as exc:
                    qr_err = f"QR generation exception: {exc}"
                finally:
                    if tmp_png_path:
                        try:
                            os.unlink(tmp_png_path)
                        except Exception:
                            pass

                report["maps_link"] = maps_link
                report["qr_error"] = qr_err
                if qr_b64:
                    report["qr_png_base64"] = qr_b64

                if isinstance(report.get("categories"), dict):
                    report["categories"].setdefault("maps", {})
                    if isinstance(report["categories"].get("maps"), dict):
                        report["categories"]["maps"]["google_maps_link"] = maps_link
                        report["categories"]["maps"]["qr_error"] = qr_err

            return report
        finally:
            if tmp_path:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

    @app.post("/api/metadata/tiff")
    async def metadata_tiff(
        file: UploadFile = File(...),
        qr: bool = Query(False, description="Generate a Google Maps QR code when GPS is present"),
    ):
        """Extract metadata from a .tif/.tiff file and return a JSON report.

        The client is responsible for saving the JSON to disk (e.g. thermal/faults/metadata).
        """
        try:
            return await _probe_tiff_metadata(file, qr)
        except HTTPException:
            raise
        except Exception as exc:
            print(f"[ERROR] Metadata probe failed: {exc}")
            raise HTTPException(status_code=500, detail=str(exc))

    @app.post("/api/metadata/probe")
    async def metadata_probe(
        file: UploadFile = File(...),
        qr: bool = Query(False, description="Generate a Google Maps QR code when GPS is present"),
    ):
        """Alias for the TIFF metadata probe (preferred endpoint name for the UI)."""
        try:
            return await _probe_tiff_metadata(file, qr)
        except HTTPException:
            raise
        except Exception as exc:
            print(f"[ERROR] Metadata probe failed: {exc}")
            raise HTTPException(status_code=500, detail=str(exc))

    def _find_soffice_executable() -> str | None:
        """Return the path to LibreOffice's `soffice` executable, or None if not found."""

        # Allow explicit configuration.
        for env_name in ("SOFFICE_PATH", "LIBREOFFICE_PATH"):
            value = os.getenv(env_name)
            if not value:
                continue
            value = value.strip().strip('"')
            if not value:
                continue
            p = Path(value)

            # Accept either the full path to soffice(.exe) or a LibreOffice install folder.
            if p.is_dir():
                candidate = p / "program" / ("soffice.exe" if os.name == "nt" else "soffice")
                if candidate.exists():
                    return str(candidate)
            elif p.exists():
                return str(p)

        # PATH lookup.
        for name in ("soffice", "soffice.exe"):
            found = shutil.which(name)
            if found:
                return found

        # Common Windows install locations.
        if os.name == "nt":
            common = [
                Path("C:/Program Files/LibreOffice/program/soffice.exe"),
                Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
            ]
            for c in common:
                if c.exists():
                    return str(c)

        return None

    def _convert_docx_bytes_to_pdf_bytes(docx_bytes: bytes, soffice_path: str, timeout_s: int = 90) -> bytes:
        """Convert DOCX bytes to PDF bytes using LibreOffice headless."""
        with tempfile.TemporaryDirectory(prefix="docx2pdf-") as tmpdir:
            tmp = Path(tmpdir)
            in_path = tmp / "input.docx"
            out_path = tmp / "input.pdf"

            in_path.write_bytes(docx_bytes)

            cmd = [
                soffice_path,
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--norestore",
                "--invisible",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp),
                str(in_path),
            ]

            proc = subprocess.run(
                cmd,
                cwd=str(tmp),
                capture_output=True,
                text=True,
                timeout=timeout_s,
            )

            if proc.returncode != 0 or not out_path.exists():
                stdout = (proc.stdout or "").strip()
                stderr = (proc.stderr or "").strip()
                msg = "LibreOffice conversion failed."
                if stderr:
                    msg += f" stderr: {stderr[:2000]}"
                elif stdout:
                    msg += f" stdout: {stdout[:2000]}"
                raise RuntimeError(msg)

            return out_path.read_bytes()

    @app.post("/api/convert/docx-to-pdf")
    async def convert_docx_to_pdf(file: UploadFile = File(...)):
        """Convert an uploaded .docx file to PDF using LibreOffice headless."""

        raw = await file.read()
        if not raw:
            raise HTTPException(status_code=400, detail="Empty file")

        filename = (file.filename or "").strip() or "report.docx"
        if not filename.lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="File must be .docx")

        # Soft limit to avoid accidental huge uploads.
        if len(raw) > 25 * 1024 * 1024:
            raise HTTPException(status_code=413, detail="File too large")

        soffice = _find_soffice_executable()
        if not soffice:
            raise HTTPException(
                status_code=503,
                detail=(
                    "LibreOffice (soffice) not found. Install LibreOffice, or set SOFFICE_PATH / LIBREOFFICE_PATH. "
                    "Example SOFFICE_PATH on Windows: C:/Program Files/LibreOffice/program/soffice.exe"
                ),
            )

        try:
            pdf_bytes = await run_in_threadpool(_convert_docx_bytes_to_pdf_bytes, raw, soffice)
        except subprocess.TimeoutExpired:
            raise HTTPException(status_code=504, detail="LibreOffice conversion timed out")
        except Exception as exc:
            print(f"[ERROR] DOCX->PDF conversion failed: {exc}")
            raise HTTPException(status_code=500, detail=str(exc))

        return Response(content=pdf_bytes, media_type="application/pdf")

    async def _read_uploaded_tiff_to_tmp_path(file: UploadFile) -> tuple[str, str]:
        raw = await file.read()
        if not raw:
            raise HTTPException(status_code=400, detail="Empty file")

        filename = file.filename or "uploaded.tiff"
        lowered = filename.lower()
        if not (lowered.endswith(".tif") or lowered.endswith(".tiff")):
            raise HTTPException(status_code=400, detail="File must be .tif or .tiff")

        suffix = ".tiff" if lowered.endswith(".tiff") else ".tif"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(raw)
            tmp_path = tmp.name
        return tmp_path, filename

    @app.post("/api/temperatures/csv")
    async def temperatures_csv(
        file: UploadFile = File(...),
        mode: str = Query("wide", description="CSV format: wide | long | long_with_xy"),
        sample: int = Query(1, ge=1, le=1000, description="Keep every Nth pixel"),
        nan_empty: bool = Query(True, description="Write NaN/Inf as empty cells"),
    ):
        """Convert a thermal TIFF to a per-pixel CSV using tiff_to_pixel_temperatures_csv.py helpers.

        Returns the CSV bytes. The client is responsible for saving the CSV.
        """
        tmp_tiff = None
        tmp_csv = None
        try:
            tmp_tiff, filename = await _read_uploaded_tiff_to_tmp_path(file)

            from backend.tiff_to_pixel_temperatures_csv import read_array, write_long_csv, write_wide_csv

            arr = await run_in_threadpool(read_array, tmp_tiff)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp:
                tmp_csv = tmp.name

            def _write():
                if mode == "wide":
                    write_wide_csv(arr, tmp_csv, sample_step=sample, nan_as_empty=nan_empty)
                    return
                if mode == "long_with_xy":
                    write_long_csv(arr, tmp_csv, include_xy=True, sample_step=sample, nan_as_empty=nan_empty)
                    return
                # default to long
                write_long_csv(arr, tmp_csv, include_xy=False, sample_step=sample, nan_as_empty=nan_empty)

            await run_in_threadpool(_write)

            with open(tmp_csv, "rb") as f:
                content = f.read()

            # Provide a useful download name hint (client may ignore)
            safe_name = os.path.basename(filename)
            out_name = f"{os.path.splitext(safe_name)[0]}.pixel_temps.{mode}.csv"
            return Response(
                content=content,
                media_type="text/csv; charset=utf-8",
                headers={"Content-Disposition": f"attachment; filename={out_name}"},
            )
        except HTTPException:
            raise
        except Exception as exc:
            print(f"[ERROR] Temperature CSV export failed: {exc}")
            raise HTTPException(status_code=500, detail=str(exc))
        finally:
            for p in [tmp_tiff, tmp_csv]:
                if p:
                    try:
                        os.unlink(p)
                    except Exception:
                        pass

    def _finite_stats(values: np.ndarray) -> dict[str, float | None]:
        if values.size == 0:
            return {"mean": None, "min": None, "max": None}
        finite = values[np.isfinite(values)]
        if finite.size == 0:
            return {"mean": None, "min": None, "max": None}
        return {
            "mean": float(np.nanmean(finite)),
            "min": float(np.nanmin(finite)),
            "max": float(np.nanmax(finite)),
        }

    def _label_bbox_px(label: dict[str, Any], width: int, height: int) -> tuple[int, int, int, int, bool]:
        # Supports both normalized (0..1) and absolute pixel coords.
        x = float(label.get("x", 0) or 0)
        y = float(label.get("y", 0) or 0)
        w = float(label.get("w", 0) or 0)
        h = float(label.get("h", 0) or 0)

        is_normalized = (x <= 1 and y <= 1 and w <= 1 and h <= 1)
        if is_normalized:
            cx = x * width
            cy = y * height
            bw = w * width
            bh = h * height
        else:
            cx = x
            cy = y
            bw = w
            bh = h

        x1 = int(round(cx - bw / 2.0))
        x2 = int(round(cx + bw / 2.0))
        y1 = int(round(cy - bh / 2.0))
        y2 = int(round(cy + bh / 2.0))

        x1 = max(0, min(width, x1))
        x2 = max(0, min(width, x2))
        y1 = max(0, min(height, y1))
        y2 = max(0, min(height, y2))

        if x2 <= x1:
            x2 = min(width, x1 + 1)
        if y2 <= y1:
            y2 = min(height, y1 + 1)

        return x1, y1, x2, y2, is_normalized

    def _ellipse_masks_for_outer_slice(
        x1o: int,
        y1o: int,
        x2o: int,
        y2o: int,
        cx: float,
        cy: float,
        a_inner: float,
        b_inner: float,
        a_outer: float,
        b_outer: float,
    ) -> tuple[np.ndarray, np.ndarray]:
        ys = (np.arange(y1o, y2o, dtype=np.float64) - cy)[:, None]
        xs = (np.arange(x1o, x2o, dtype=np.float64) - cx)[None, :]

        def inside(a: float, b: float) -> np.ndarray:
            if a <= 0 or b <= 0:
                return np.zeros((y2o - y1o, x2o - x1o), dtype=bool)
            return ((xs / a) ** 2 + (ys / b) ** 2) <= 1.0

        inner = inside(a_inner, b_inner)
        outer = inside(a_outer, b_outer)
        return inner, outer

    @app.post("/api/temperatures/labels")
    async def temperatures_labels(
        file: UploadFile = File(...),
        labels: str = Form(..., description="JSON-encoded list of label objects"),
        pad_px: int = Query(3, ge=1, le=100, description="Thickness (px) for the outside-edge perimeter band"),
    ):
        """Compute per-label temperature stats from a thermal TIFF.

        For each label returns:
        - outside_edge_mean: avg temperature outside but near the edge (a band around the label)
        - inside_mean: avg temperature inside the label
        - inside_min: min temperature inside the label
        - inside_max: max temperature inside the label
        """
        tmp_tiff = None
        try:
            tmp_tiff, filename = await _read_uploaded_tiff_to_tmp_path(file)

            try:
                labels_list = json.loads(labels)
                if not isinstance(labels_list, list):
                    raise ValueError("labels must be a JSON list")
            except Exception:
                raise HTTPException(status_code=400, detail="Invalid labels JSON")

            from backend.tiff_to_pixel_temperatures_csv import read_array

            arr = await run_in_threadpool(read_array, tmp_tiff)
            if arr.ndim != 2:
                raise HTTPException(status_code=400, detail="Expected a 2D TIFF image")

            a = arr.astype(np.float64, copy=False)
            h, w = a.shape

            results: list[dict[str, Any]] = []
            for i, label in enumerate(labels_list):
                if not isinstance(label, dict):
                    continue

                shape = str(label.get("shape") or "rect")
                source = str(label.get("source") or "auto")

                x1, y1, x2, y2, _ = _label_bbox_px(label, w, h)

                # Outer bbox for edge band
                x1o = max(0, x1 - pad_px)
                y1o = max(0, y1 - pad_px)
                x2o = min(w, x2 + pad_px)
                y2o = min(h, y2 + pad_px)

                inside_vals: np.ndarray
                edge_vals: np.ndarray

                if shape == "ellipse":
                    cx = (x1 + x2) / 2.0
                    cy = (y1 + y2) / 2.0
                    a_inner = max(0.5, (x2 - x1) / 2.0)
                    b_inner = max(0.5, (y2 - y1) / 2.0)
                    a_outer = a_inner + pad_px
                    b_outer = b_inner + pad_px

                    block = a[y1o:y2o, x1o:x2o]
                    inner_mask, outer_mask = _ellipse_masks_for_outer_slice(
                        x1o, y1o, x2o, y2o, cx, cy, a_inner, b_inner, a_outer, b_outer
                    )
                    inside_vals = block[inner_mask]
                    edge_vals = block[outer_mask & (~inner_mask)]
                else:
                    inside_block = a[y1:y2, x1:x2]
                    inside_vals = inside_block.reshape(-1)

                    outer_block = a[y1o:y2o, x1o:x2o]
                    edge_mask = np.ones(outer_block.shape, dtype=bool)
                    iy1 = y1 - y1o
                    iy2 = y2 - y1o
                    ix1 = x1 - x1o
                    ix2 = x2 - x1o
                    edge_mask[iy1:iy2, ix1:ix2] = False
                    edge_vals = outer_block[edge_mask]

                inside_stats = _finite_stats(inside_vals)
                edge_stats = _finite_stats(edge_vals)

                results.append(
                    {
                        "index": i,
                        "shape": shape,
                        "source": source,
                        "bbox_px": {"x1": x1, "y1": y1, "x2": x2, "y2": y2},
                        "counts": {"inside": int(np.isfinite(inside_vals).sum()), "edge": int(np.isfinite(edge_vals).sum())},
                        "outside_edge_mean": edge_stats["mean"],
                        "inside_mean": inside_stats["mean"],
                        "inside_min": inside_stats["min"],
                        "inside_max": inside_stats["max"],
                    }
                )

            return {
                "file": filename,
                "shape": {"width": int(w), "height": int(h)},
                "pad_px": int(pad_px),
                "labels": results,
            }
        except HTTPException:
            raise
        except Exception as exc:
            print(f"[ERROR] Label temperature stats failed: {exc}")
            raise HTTPException(status_code=500, detail=str(exc))
        finally:
            if tmp_tiff:
                try:
                    os.unlink(tmp_tiff)
                except Exception:
                    pass

    return app


_BACKEND_DIR = Path(__file__).resolve().parents[1]

MODEL_PATH = os.getenv("FAULT_MODEL_PATH", str(_BACKEND_DIR / "models" / "best.pt"))
TYPE_MODEL_PATH = os.getenv("FAULT_TYPE_MODEL_PATH", str(_BACKEND_DIR / "models" / "best_8_class.pt"))
IMGSZ_FAULT = int(os.getenv("FAULT_IMGSZ", "1280"))
CONF_FAULT = float(os.getenv("FAULT_CONF", "0.20"))
IOU_FAULT = float(os.getenv("FAULT_IOU", "0.50"))

IMGSZ_TYPE = int(os.getenv("FAULT_TYPE_IMGSZ", "640"))
# Conf used for the *predict* call on crops (low threshold to collect candidates).
CONF_TYPE_PRED = float(os.getenv("FAULT_TYPE_PRED_CONF", "0.15"))
# Minimum confidence required to accept the predicted type.
CONF_TYPE_MIN = float(os.getenv("FAULT_TYPE_CONF", "0.45"))
IOU_TYPE = float(os.getenv("FAULT_TYPE_IOU", "0.50"))
TYPE_PAD_PX = int(os.getenv("FAULT_TYPE_PAD", "12"))

# 8 classes from the type model are 0..7; we add 9 for Unknown when no type is found.
UNKNOWN_TYPE_ID = 9

_model_lock = threading.Lock()
_model: YOLO | None = None

_type_model_lock = threading.Lock()
_type_model: YOLO | None = None


def get_model() -> YOLO:
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                _model = YOLO(MODEL_PATH)
    return _model


def get_type_model() -> YOLO | None:
    """Load the 8-class defect type model.

    Returns None if the model file doesn't exist or cannot be loaded.
    """
    global _type_model
    if _type_model is None:
        with _type_model_lock:
            if _type_model is None:
                try:
                    if not Path(TYPE_MODEL_PATH).exists():
                        print(f"[WARN] Type model not found at: {TYPE_MODEL_PATH}")
                        return None
                    _type_model = YOLO(TYPE_MODEL_PATH)
                except Exception as exc:
                    print(f"[WARN] Failed to load type model: {exc}")
                    return None
    return _type_model


def to_gray_3ch(image: np.ndarray) -> np.ndarray:
    if image.ndim == 2:
        gray = image
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    return cv2.merge([gray, gray, gray])


def detect_faults(img: np.ndarray) -> List[Dict[str, Any]]:
    """Detect defects and classify each detection into a type.

    Pipeline:
      1) Run 1-class detector (best.pt) to find defect boxes.
      2) For each box, crop with padding and run the 8-class type model.
      3) Output `classId` in range 0..9, where 9 means Unknown.
    """
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
    h, w = img.shape[:2]

    # Prepare crops for the type model.
    crops: List[np.ndarray] = []
    crop_to_label_idx: List[int] = []
    padded_boxes: List[tuple[int, int, int, int]] = []

    # Precompute normalized labels + keep pixel bboxes for cropping.
    norm_labels: List[Dict[str, Any]] = []
    for i, ((x1, y1, x2, y2), conf) in enumerate(zip(xyxy, confs)):
        x1f, y1f, x2f, y2f = float(x1), float(y1), float(x2), float(y2)
        cx = ((x1f + x2f) / 2.0) / w
        cy = ((y1f + y2f) / 2.0) / h
        bw = (x2f - x1f) / w
        bh = (y2f - y1f) / h
        norm_labels.append(
            {
                # Will be overridden by the type model below.
                "classId": UNKNOWN_TYPE_ID,
                "conf": float(conf),
                "x": round(cx, 6),
                "y": round(cy, 6),
                "w": round(bw, 6),
                "h": round(bh, 6),
            }
        )

        # Crop region in pixel coords (with padding).
        x1i, y1i, x2i, y2i = int(x1f), int(y1f), int(x2f), int(y2f)
        x1p = max(0, x1i - TYPE_PAD_PX)
        y1p = max(0, y1i - TYPE_PAD_PX)
        x2p = min(w, x2i + TYPE_PAD_PX)
        y2p = min(h, y2i + TYPE_PAD_PX)
        if x2p <= x1p or y2p <= y1p:
            continue
        crop = img[y1p:y2p, x1p:x2p]
        if crop.size == 0:
            continue
        crops.append(crop)
        crop_to_label_idx.append(i)
        padded_boxes.append((x1p, y1p, x2p, y2p))

    type_model = get_type_model()
    if type_model is None or not crops:
        return norm_labels

    try:
        # Batch predict is significantly faster than per-crop predict.
        r2_list = type_model.predict(
            source=crops,
            imgsz=IMGSZ_TYPE,
            conf=CONF_TYPE_PRED,
            iou=IOU_TYPE,
            verbose=False,
        )
    except Exception as exc:
        print(f"[WARN] Type inference failed: {exc}")
        return norm_labels

    for label_idx, r2 in zip(crop_to_label_idx, r2_list):
        try:
            if r2.boxes is None or len(r2.boxes) == 0:
                norm_labels[label_idx]["classId"] = UNKNOWN_TYPE_ID
                continue

            conf_arr = r2.boxes.conf
            cls_arr = r2.boxes.cls
            if conf_arr is None or cls_arr is None or len(conf_arr) == 0:
                norm_labels[label_idx]["classId"] = UNKNOWN_TYPE_ID
                continue

            confs2 = conf_arr.cpu().numpy()
            classes2 = cls_arr.cpu().numpy()
            best_i = int(np.argmax(confs2))
            best_conf = float(confs2[best_i])
            best_cls = int(classes2[best_i])

            if best_conf >= CONF_TYPE_MIN and 0 <= best_cls <= 7:
                norm_labels[label_idx]["classId"] = best_cls
                norm_labels[label_idx]["typeConf"] = round(best_conf, 6)
            else:
                norm_labels[label_idx]["classId"] = UNKNOWN_TYPE_ID
                norm_labels[label_idx]["typeConf"] = round(best_conf, 6)
        except Exception:
            norm_labels[label_idx]["classId"] = UNKNOWN_TYPE_ID

    return norm_labels


app = create_app()
