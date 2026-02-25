import glob
import math
import numpy as np
import tifffile as tiff
import exifread
import rasterio
from rasterio.transform import from_origin
from rasterio.merge import merge
from rasterio.io import MemoryFile
from pyproj import CRS, Transformer
import cv2
import os

INPUT_GLOB = r"F:\project\thermal1_images\tiff\*.tif*"
OUT_TIF = r"F:\project\thermal1_images\quick_mosaic_agl.tiff"

# ΒΑΛΕ ΕΔΩ το ύψος πτήσης ΠΑΝΩ ΑΠΟ ΤΟ ΕΔΑΦΟΣ (AGL)
AGL_ALT_M = 80.0

# ΒΑΛΕ ΕΔΩ (αν δεν ξέρεις, άστο έτσι για αρχή)
HFOV_DEG = 57.0
VFOV_DEG = 44.0

MAX_W = 1200  # για ταχύτητα

def dms_to_deg(dms, ref):
    deg = float(dms[0].num)/dms[0].den
    minute = float(dms[1].num)/dms[1].den
    sec = float(dms[2].num)/dms[2].den
    val = deg + minute/60 + sec/3600
    return -val if ref in ("S","W") else val

def read_latlon(path):
    with open(path, "rb") as fh:
        tags = exifread.process_file(fh, details=False)
    if "GPS GPSLatitude" not in tags or "GPS GPSLongitude" not in tags:
        return None, None
    lat = dms_to_deg(tags["GPS GPSLatitude"].values, str(tags["GPS GPSLatitudeRef"]))
    lon = dms_to_deg(tags["GPS GPSLongitude"].values, str(tags["GPS GPSLongitudeRef"]))
    return lat, lon

def to_uint8_3ch(img):
    img = np.nan_to_num(img)
    if img.ndim > 3:
        img = img[0]
    if img.ndim == 3 and img.shape[2] == 1:
        img = img[:, :, 0]

    if img.dtype != np.uint8:
        lo, hi = np.percentile(img, (2, 98))
        if hi <= lo:
            lo, hi = float(np.min(img)), float(np.max(img))
            if hi <= lo:
                hi = lo + 1.0
        img = np.clip((img - lo) / (hi - lo), 0, 1)
        img8 = (img * 255).astype(np.uint8)
    else:
        img8 = img

    if img8.ndim == 2:
        img8 = cv2.cvtColor(img8, cv2.COLOR_GRAY2BGR)
    elif img8.shape[2] == 4:
        img8 = cv2.cvtColor(img8, cv2.COLOR_BGRA2BGR)

    h, w = img8.shape[:2]
    if w > MAX_W:
        s = MAX_W / w
        img8 = cv2.resize(img8, (int(w*s), int(h*s)), interpolation=cv2.INTER_AREA)

    return img8

def footprint_meters(alt_m, hfov_deg, vfov_deg):
    fw = 2.0 * alt_m * math.tan(math.radians(hfov_deg) / 2.0)
    fh = 2.0 * alt_m * math.tan(math.radians(vfov_deg) / 2.0)
    return fw, fh

def utm_crs_from_lonlat(lon, lat):
    zone = int((lon + 180) / 6) + 1
    epsg = 32600 + zone if lat >= 0 else 32700 + zone
    return CRS.from_epsg(epsg)

paths = sorted(glob.glob(INPUT_GLOB))
if len(paths) < 2:
    raise RuntimeError("Δεν βρέθηκαν αρκετές εικόνες.")

# CRS από την πρώτη εικόνα με GPS
lat0 = lon0 = None
for p in paths:
    lat, lon = read_latlon(p)
    if lat is not None:
        lat0, lon0 = lat, lon
        break
if lat0 is None:
    raise RuntimeError("Δεν βρέθηκαν GPS.")

target_crs = utm_crs_from_lonlat(lon0, lat0)
to_utm = Transformer.from_crs(CRS.from_epsg(4326), target_crs, always_xy=True)

mem_datasets = []
used = 0

for i, p in enumerate(paths):
    lat, lon = read_latlon(p)
    if lat is None:
        continue

    img = tiff.imread(p)
    img8 = to_uint8_3ch(img)
    hpx, wpx = img8.shape[:2]

    fw_m, fh_m = footprint_meters(AGL_ALT_M, HFOV_DEG, VFOV_DEG)
    px_x = fw_m / wpx
    px_y = fh_m / hpx

    x, y = to_utm.transform(lon, lat)
    x0 = x - fw_m / 2.0
    y0 = y + fh_m / 2.0
    transform = from_origin(x0, y0, px_x, px_y)

    profile = {
        "driver": "GTiff",
        "height": hpx,
        "width": wpx,
        "count": 3,
        "dtype": np.uint8,
        "crs": target_crs,
        "transform": transform,
        "compress": "LZW",
        "tiled": True
    }

    mem = MemoryFile()
    ds = mem.open(**profile)
    ds.write(np.transpose(img8, (2, 0, 1)))
    mem_datasets.append((mem, ds))
    used += 1

    if (i + 1) % 50 == 0:
        print(f"Prepared {i+1}/{len(paths)}")

print("Used images:", used)
srcs = [ds for _, ds in mem_datasets]
mosaic, out_transform = merge(srcs, method="first")

out_profile = srcs[0].profile.copy()
out_profile.update({
    "height": mosaic.shape[1],
    "width": mosaic.shape[2],
    "transform": out_transform
})

with rasterio.open(OUT_TIF, "w", **out_profile) as dst:
    dst.write(mosaic)

for mem, ds in mem_datasets:
    ds.close()
    mem.close()

print("Saved:", OUT_TIF)
