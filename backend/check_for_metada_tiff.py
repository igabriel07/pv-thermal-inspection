#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
check_for_metadata_tiff.py

Δώσε ένα .tif/.tiff και θα σου δείξει:
- Metadata από ExifTool (αν υπάρχει εγκατεστημένο) -> EXIF/XMP/MakerNotes
- Metadata από Pillow (fallback)
- DJI XMP που είναι μέσα στο XMLPacket (ακόμα και χωρίς ExifTool)
- Pixel stats (min/mean/max) από tifffile (raw, εδώ μοιάζουν με °C αν είναι float)
- Αποθήκευση πλήρους JSON report χωρίς σφάλματα serialization
- Κατηγοριοποιημένο output (image_info, device, timestamps, geolocation, flight, measurement_params, temps, pixel_data)
- Προαιρετικά: QR code PNG με Google Maps link από τις συντεταγμένες

ΣΗΜΑΝΤΙΚΟ FIX:
- Δεν χρησιμοποιούμε ΠΟΤΕ RtkStdLat/Lon ως συντεταγμένες (είναι std-dev/accuracy).
- Προτιμάμε ExifTool Composite:GPSLatitude/Longitude ή GPS:GPSLatitude/Longitude
- Robust parsing (DMS / N,S,E,W) + sanity checks lat/lon bounds

Χρήση:
  python check_for_metadata_tiff.py "C:\\path\\file.tiff"
  python check_for_metadata_tiff.py "C:\\path\\file.tiff" --out report.json
  python check_for_metadata_tiff.py "C:\\path\\file.tiff" --qr
  python check_for_metadata_tiff.py "C:\\path\\file.tiff" --qr --qr-out "C:\\path\\maps_qr.png"

Απαιτήσεις:
  pip install pillow tifffile numpy
Προαιρετικά (συνιστάται):
  εγκατάσταση ExifTool
  pip install qrcode[pil]   (για QR)
"""

import argparse
import json
import os
import re
import shutil
import subprocess
from dataclasses import dataclass
from typing import Any, Dict, Optional, Tuple, List

import numpy as np

try:
    from PIL import Image, ExifTags
except Exception:
    Image = None
    ExifTags = None

try:
    import tifffile
except Exception:
    tifffile = None

try:
    import qrcode
except Exception:
    qrcode = None

import xml.etree.ElementTree as ET


# -----------------------------
# JSON-safe conversion
# -----------------------------

def to_jsonable(x: Any) -> Any:
    """Convert common non-JSON types (Pillow IFDRational, bytes, numpy) to JSON-safe."""
    if x is None:
        return None

    if isinstance(x, (np.integer, np.floating, np.bool_)):
        return x.item()

    if isinstance(x, np.ndarray):
        return x.tolist()

    if isinstance(x, (bytes, bytearray)):
        b = bytes(x)
        for enc in ("utf-16le", "utf-8", "latin-1"):
            try:
                s = b.decode(enc, errors="strict")
                s = s.replace("\x00", "").strip()
                if s:
                    return s
            except Exception:
                pass
        return b.hex()

    if hasattr(x, "numerator") and hasattr(x, "denominator"):
        try:
            n = int(x.numerator)
            d = int(x.denominator)
            if d != 0:
                return n / d
            return None
        except Exception:
            pass

    if isinstance(x, (list, tuple)):
        return [to_jsonable(v) for v in x]

    if isinstance(x, dict):
        return {str(k): to_jsonable(v) for k, v in x.items()}

    if isinstance(x, (str, int, float, bool)):
        return x

    return str(x)


# -----------------------------
# ExifTool
# -----------------------------

def which_exiftool() -> Optional[str]:
    return shutil.which("exiftool")


def run_exiftool_json(path: str) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Returns (metadata_dict, error_message)
    """
    exiftool = which_exiftool()
    if not exiftool:
        return None, "exiftool not found in PATH"

    cmd = [exiftool, "-j", "-G1", "-a", "-s", "-u", "-ee", path]
    try:
        out = subprocess.check_output(cmd, text=True, stderr=subprocess.STDOUT)
        data = json.loads(out)
        if not data:
            return {}, None
        return data[0], None
    except subprocess.CalledProcessError as e:
        return None, f"exiftool failed: {e.output}"
    except json.JSONDecodeError as e:
        return None, f"exiftool output not JSON: {e}"


# -----------------------------
# Pillow EXIF + GPS decode
# -----------------------------

def _dms_to_deg(dms) -> Optional[float]:
    try:
        d, m, s = dms
        d = to_jsonable(d)
        m = to_jsonable(m)
        s = to_jsonable(s)
        if None in (d, m, s):
            return None
        return float(d) + float(m) / 60.0 + float(s) / 3600.0
    except Exception:
        return None


def decode_gpsinfo_if_possible(gps_ifd: Any) -> Dict[str, Any]:
    if not isinstance(gps_ifd, dict) or ExifTags is None:
        return {}

    gps_tags = {ExifTags.GPSTAGS.get(k, str(k)): v for k, v in gps_ifd.items()}

    lat = None
    lon = None

    if "GPSLatitude" in gps_tags and "GPSLatitudeRef" in gps_tags:
        lat = _dms_to_deg(gps_tags["GPSLatitude"])
        ref = str(to_jsonable(gps_tags["GPSLatitudeRef"])).upper()
        if lat is not None and ref == "S":
            lat = -lat

    if "GPSLongitude" in gps_tags and "GPSLongitudeRef" in gps_tags:
        lon = _dms_to_deg(gps_tags["GPSLongitude"])
        ref = str(to_jsonable(gps_tags["GPSLongitudeRef"])).upper()
        if lon is not None and ref == "W":
            lon = -lon

    out = {"gps_raw": {k: to_jsonable(v) for k, v in gps_tags.items()}}
    if lat is not None:
        out["latitude"] = lat
    if lon is not None:
        out["longitude"] = lon
    return out


def read_basic_exif_pillow(path: str) -> Tuple[Dict[str, Any], Optional[str]]:
    if Image is None:
        return {}, "Pillow not available (pip install pillow)"

    try:
        img = Image.open(path)
        exif = img.getexif()
        if not exif:
            return {}, None

        tag_map: Dict[str, Any] = {}
        for k, v in exif.items():
            name = ExifTags.TAGS.get(k, str(k)) if ExifTags else str(k)
            tag_map[name] = v

        if "GPSInfo" in tag_map and isinstance(tag_map["GPSInfo"], dict):
            tag_map["GPSDecoded"] = decode_gpsinfo_if_possible(tag_map["GPSInfo"])

        return tag_map, None
    except Exception as e:
        return {}, f"Pillow EXIF read error: {e}"


# -----------------------------
# DJI XMP parsing from XMLPacket
# -----------------------------

def _try_extract_xml_fragment(text: str) -> Optional[str]:
    start = text.find("<x:xmpmeta")
    if start == -1:
        start = text.find("<xmpmeta")
    if start == -1:
        return None

    end = text.find("</x:xmpmeta>")
    end_tag = "</x:xmpmeta>"
    if end == -1:
        end = text.find("</xmpmeta>")
        end_tag = "</xmpmeta>"
    if end == -1:
        return None

    return text[start:end + len(end_tag)]


def _decode_xmlpacket_to_text(xmlpacket: Any) -> Optional[str]:
    if xmlpacket is None:
        return None

    if isinstance(xmlpacket, (bytes, bytearray)):
        b = bytes(xmlpacket)
        for enc in ("utf-8", "utf-16le", "utf-16be", "latin-1"):
            try:
                t = b.decode(enc, errors="ignore")
                frag = _try_extract_xml_fragment(t)
                if frag:
                    return frag
            except Exception:
                continue
        return None

    if isinstance(xmlpacket, str):
        try:
            raw = xmlpacket.encode("latin-1", errors="ignore")
            for enc in ("utf-16le", "utf-16be", "utf-8"):
                t = raw.decode(enc, errors="ignore")
                frag = _try_extract_xml_fragment(t)
                if frag:
                    return frag
        except Exception:
            pass

        frag = _try_extract_xml_fragment(xmlpacket)
        if frag:
            return frag

    return None


def parse_dji_xmp_from_xmlpacket(xmlpacket: Any) -> Dict[str, Any]:
    xml_text = _decode_xmlpacket_to_text(xmlpacket)
    if not xml_text:
        return {}

    cleaned = "".join(ch for ch in xml_text if ch in ("\n", "\t") or (ord(ch) >= 32))
    try:
        root = ET.fromstring(cleaned)
    except Exception:
        cleaned2 = re.sub(r"[^\x09\x0A\x0D\x20-\x7E\u0080-\uFFFF]", "", xml_text)
        try:
            root = ET.fromstring(cleaned2)
        except Exception:
            return {}

    ns = {}
    for m in re.finditer(r'xmlns:([A-Za-z0-9_\-]+)="([^"]+)"', cleaned):
        ns[m.group(1)] = m.group(2)

    tags: Dict[str, Any] = {}
    dji_simple: Dict[str, Any] = {}

    for elem in root.iter():
        tag = elem.tag
        text = (elem.text or "").strip() if elem.text else ""

        if tag.startswith("{"):
            uri, local = tag[1:].split("}", 1)
            if "dji" in uri.lower() and ("drone" in uri.lower() or "dji.com" in uri.lower()):
                key = f"{{{uri}}}{local}"
                tags[key] = text
                dji_simple[local] = text

        if ":" in tag and not tag.startswith("{"):
            prefix, local = tag.split(":", 1)
            if prefix.lower() in ("drone-dji", "dji", "djidrone", "dronedji"):
                key = f"{prefix}:{local}"
                tags[key] = text
                dji_simple[local] = text

    # IMPORTANT: DO NOT infer GPS from RtkStdLat/Lon (they are std-dev)
    out = {"namespaces": ns, "tags": tags, "dji_simple": dji_simple}
    return out


# -----------------------------
# Pixels
# -----------------------------

def read_tiff_pixels(path: str) -> Tuple[Dict[str, Any], Optional[str]]:
    if tifffile is None:
        return {}, "tifffile not available (pip install tifffile)"
    try:
        arr = tifffile.imread(path)

        info: Dict[str, Any] = {"shape": list(arr.shape), "dtype": str(arr.dtype), "ndim": int(arr.ndim)}
        a = arr.astype(np.float64)

        if arr.ndim == 2:
            info["raw_min"] = float(np.min(a))
            info["raw_mean"] = float(np.mean(a))
            info["raw_max"] = float(np.max(a))
        else:
            info["raw_min_all"] = float(np.min(a))
            info["raw_mean_all"] = float(np.mean(a))
            info["raw_max_all"] = float(np.max(a))

        info["radiometric_guess"] = bool(arr.ndim == 2 and arr.dtype in (np.uint16, np.int16, np.uint32, np.int32))
        info["looks_like_celsius_guess"] = bool(
            arr.ndim == 2
            and arr.dtype in (np.float32, np.float64)
            and -50.0 <= float(np.min(a)) <= 200.0
            and -50.0 <= float(np.max(a)) <= 200.0
        )

        return info, None
    except Exception as e:
        return {}, f"tifffile read error: {e}"


# -----------------------------
# Helpers for picking keys + parsing lat/lon robustly
# -----------------------------

def normalize_key(k: str) -> str:
    return k.strip().lower().replace(" ", "").replace("_", "")


def pick_first(meta: Dict[str, Any], candidates: List[str]) -> Optional[Any]:
    for c in candidates:
        if c in meta:
            return meta[c]
    norm = {normalize_key(k): k for k in meta.keys()}
    for c in candidates:
        nc = normalize_key(c)
        if nc in norm:
            return meta[norm[nc]]
    return None


def find_matching_keys(meta: Dict[str, Any], patterns: List[str]) -> Dict[str, Any]:
    out = {}
    for k, v in meta.items():
        for pat in patterns:
            if re.search(pat, k, flags=re.IGNORECASE):
                out[k] = v
                break
    return out


def parse_latlon_value(v: Any) -> Optional[float]:
    """
    Parse lat/lon from:
    - float/int
    - "37.12345"
    - "37,12345"
    - "40 deg 22' 51.17\" N"
    - "21 deg 47' 45.78\" E"
    - "40 22 51.17 N"
    """
    if v is None:
        return None

    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)

    s = str(v).strip()
    if not s:
        return None

    s_up = s.upper()
    s_norm = s.replace(",", ".").strip()

    # try simple float first
    try:
        return float(s_norm)
    except Exception:
        pass

    hemi = None
    for h in ("N", "S", "E", "W"):
        if h in s_up:
            hemi = h
            break

    nums = re.findall(r"[-+]?\d+(?:\.\d+)?", s_norm)
    if not nums:
        return None

    try:
        if len(nums) >= 3:
            deg = float(nums[0])
            minutes = float(nums[1])
            sec = float(nums[2])
            val = abs(deg) + minutes / 60.0 + sec / 3600.0
            if str(nums[0]).startswith("-"):
                val = -val
        else:
            val = float(nums[0])
    except Exception:
        return None

    if hemi in ("S", "W"):
        val = -abs(val)
    elif hemi in ("N", "E"):
        val = abs(val)

    return val


def valid_latlon(lat: Optional[float], lon: Optional[float]) -> bool:
    if lat is None or lon is None:
        return False
    return (-90.0 <= lat <= 90.0) and (-180.0 <= lon <= 180.0)


# -----------------------------
# Summary
# -----------------------------

def build_summary(meta: Dict[str, Any], dji_simple: Dict[str, Any]) -> Dict[str, Any]:
    s = {}
    s["Make"] = pick_first(meta, ["EXIF:Make", "IFD0:Make", "Make"])
    s["Model"] = pick_first(meta, ["EXIF:Model", "IFD0:Model", "Model"])
    s["CameraSerialNumber"] = pick_first(meta, ["EXIF:CameraSerialNumber", "IFD0:CameraSerialNumber", "CameraSerialNumber"])
    s["FocalLength"] = pick_first(meta, ["EXIF:FocalLength", "ExifIFD:FocalLength", "FocalLength"])
    s["FNumber"] = pick_first(meta, ["EXIF:FNumber", "ExifIFD:FNumber", "FNumber"])
    s["ImageWidth"] = pick_first(meta, ["File:ImageWidth", "IFD0:ImageWidth", "ImageWidth", "ExifIFD:ExifImageWidth"])
    s["ImageHeight"] = pick_first(meta, ["File:ImageHeight", "IFD0:ImageHeight", "ImageLength", "ImageHeight", "ExifIFD:ExifImageHeight"])
    s["DateTimeOriginal"] = pick_first(meta, ["EXIF:DateTimeOriginal", "ExifIFD:DateTimeOriginal", "DateTimeOriginal"])
    s["DateTime"] = pick_first(meta, ["EXIF:DateTime", "IFD0:ModifyDate", "DateTime"])
    s["Software"] = pick_first(meta, ["Software", "IFD0:Software", "EXIF:Software"])
    s["ImageDescription"] = pick_first(meta, ["ImageDescription", "IFD0:ImageDescription", "EXIF:ImageDescription"])

    # GPS (prefer exiftool composite)
    lat = parse_latlon_value(pick_first(meta, ["Composite:GPSLatitude", "GPS:GPSLatitude", "EXIF:GPSLatitude", "GPSLatitude"]))
    lon = parse_latlon_value(pick_first(meta, ["Composite:GPSLongitude", "GPS:GPSLongitude", "EXIF:GPSLongitude", "GPSLongitude"]))
    if valid_latlon(lat, lon):
        s["Latitude"] = lat
        s["Longitude"] = lon

    for key in (
        "AbsoluteAltitude", "RelativeAltitude",
        "GimbalYawDegree", "GimbalPitchDegree", "GimbalRollDegree",
        "FlightYawDegree", "FlightPitchDegree", "FlightRollDegree"
    ):
        if key in dji_simple and dji_simple[key] not in (None, ""):
            s[f"DJI_{key}"] = dji_simple[key]

    s["Emissivity"] = pick_first(meta, ["Emissivity", "XMP:Emissivity", "EXIF:Emissivity"])
    s["Distance"] = pick_first(meta, ["Distance", "SubjectDistance", "EXIF:SubjectDistance", "ExifIFD:SubjectDistance"])
    s["AmbientTemperature"] = pick_first(meta, ["AmbientTemperature", "Ambient Temperature"])
    s["RelativeHumidity"] = pick_first(meta, ["RelativeHumidity", "Relative Humidity", "Humidity"])
    s["WindSpeed"] = pick_first(meta, ["WindSpeed", "Wind Speed"])
    s["Irradiance"] = pick_first(meta, ["Irradiance", "SolarIrradiance", "Solar Irradiance"])

    return s


# -----------------------------
# Categorization
# -----------------------------

def build_categories(combined: Dict[str, Any], dji_simple: Dict[str, Any], pixels_info: Dict[str, Any]) -> Dict[str, Any]:
    def first(keys: List[str]) -> Optional[Any]:
        return pick_first(combined, keys)

    cats: Dict[str, Any] = {}

    cats["image_info"] = {
        "ImageWidth": first(["File:ImageWidth", "IFD0:ImageWidth", "ImageWidth", "ExifIFD:ExifImageWidth"]),
        "ImageHeight": first(["File:ImageHeight", "IFD0:ImageHeight", "ImageLength", "ImageHeight", "ExifIFD:ExifImageHeight"]),
        "BitsPerSample": first(["IFD0:BitsPerSample", "BitsPerSample", "XMP-tiff:BitsPerSample"]),
        "Compression": first(["IFD0:Compression", "Compression", "XMP-tiff:Compression"]),
        "PhotometricInterpretation": first(["IFD0:PhotometricInterpretation", "PhotometricInterpretation"]),
        "PlanarConfiguration": first(["IFD0:PlanarConfiguration", "PlanarConfiguration"]),
        "Software": first(["IFD0:Software", "Software"]),
        "ImageDescription": first(["IFD0:ImageDescription", "ImageDescription"]),
        "SampleFormat": first(["IFD0:SampleFormat", "SampleFormat"]),
    }

    cats["device"] = {
        "Make": first(["IFD0:Make", "EXIF:Make", "Make"]),
        "Model": first(["IFD0:Model", "EXIF:Model", "Model"]),
        "SerialNumber": first(["ExifIFD:SerialNumber", "IFD0:CameraSerialNumber", "EXIF:CameraSerialNumber", "CameraSerialNumber"]),
        "LensModel": first(["EXIF:LensModel", "LensModel"]),
        "FocalLength": first(["ExifIFD:FocalLength", "EXIF:FocalLength", "FocalLength"]),
        "FNumber": first(["ExifIFD:FNumber", "EXIF:FNumber", "FNumber"]),
    }

    cats["timestamps"] = {
        "DateTimeOriginal": first(["ExifIFD:DateTimeOriginal", "EXIF:DateTimeOriginal", "DateTimeOriginal"]),
        "CreateDate": first(["ExifIFD:CreateDate", "EXIF:CreateDate", "CreateDate"]),
        "ModifyDate": first(["IFD0:ModifyDate", "EXIF:ModifyDate", "ModifyDate"]),
        "DateTime": first(["EXIF:DateTime", "DateTime"]),
        "GPSDateTime": first(["Composite:GPSDateTime", "XMP-exif:GPSDateTime", "GPS:GPSDateStamp"]),
    }

    # Geolocation: prefer exiftool composite / gps group
    lat = parse_latlon_value(first(["Composite:GPSLatitude", "GPS:GPSLatitude", "EXIF:GPSLatitude", "GPSLatitude"]))
    lon = parse_latlon_value(first(["Composite:GPSLongitude", "GPS:GPSLongitude", "EXIF:GPSLongitude", "GPSLongitude"]))

    # fallback: Pillow decoded GPS (rare on DJI TIFF)
    gps_decoded = combined.get("GPSDecoded")
    if (not valid_latlon(lat, lon)) and isinstance(gps_decoded, dict):
        lat2 = gps_decoded.get("latitude")
        lon2 = gps_decoded.get("longitude")
        if valid_latlon(lat2, lon2):
            lat, lon = float(lat2), float(lon2)

    # DO NOT use RtkStdLat/Lon as coords; only keep them as accuracy fields under flight
    if not valid_latlon(lat, lon):
        lat, lon = None, None

    cats["geolocation"] = {
        "latitude": lat,
        "longitude": lon,
        "altitude": first(["Composite:GPSAltitude", "GPS:GPSAltitude", "EXIF:GPSAltitude"]),
        "gps_position": first(["Composite:GPSPosition"]),
        "map_datum": first(["GPS:GPSMapDatum"]),
    }

    cats["flight"] = {
        "AbsoluteAltitude": dji_simple.get("AbsoluteAltitude"),
        "RelativeAltitude": dji_simple.get("RelativeAltitude"),
        "FlightYawDegree": dji_simple.get("FlightYawDegree"),
        "FlightPitchDegree": dji_simple.get("FlightPitchDegree"),
        "FlightRollDegree": dji_simple.get("FlightRollDegree"),
        "GimbalYawDegree": dji_simple.get("GimbalYawDegree"),
        "GimbalPitchDegree": dji_simple.get("GimbalPitchDegree"),
        "GimbalRollDegree": dji_simple.get("GimbalRollDegree"),
        "FlightXSpeed": dji_simple.get("FlightXSpeed"),
        "FlightYSpeed": dji_simple.get("FlightYSpeed"),
        "FlightZSpeed": dji_simple.get("FlightZSpeed"),
        "RtkFlag": dji_simple.get("RtkFlag"),
        # keep RTK std-dev as accuracy (NOT coords)
        "RtkStdLat": dji_simple.get("RtkStdLat"),
        "RtkStdLon": dji_simple.get("RtkStdLon"),
        "RtkStdHgt": dji_simple.get("RtkStdHgt"),
    }

    cats["measurement_params"] = {
        "Emissivity": first(["Emissivity", "XMP:Emissivity", "EXIF:Emissivity"]),
        "Distance": first(["Distance", "SubjectDistance", "EXIF:SubjectDistance", "ExifIFD:SubjectDistance"]),
        "AmbientTemperature": first(["AmbientTemperature", "Ambient Temperature"]),
        "RelativeHumidity": first(["RelativeHumidity", "Relative Humidity", "Humidity"]),
        "WindSpeed": first(["WindSpeed", "Wind Speed"]),
        "Irradiance": first(["Irradiance", "SolarIrradiance", "Solar Irradiance"]),
    }

    pixel_stats = None
    if "raw_min" in pixels_info:
        pixel_stats = {
            "min": pixels_info.get("raw_min"),
            "mean": pixels_info.get("raw_mean"),
            "max": pixels_info.get("raw_max"),
            "looks_like_celsius_guess": pixels_info.get("looks_like_celsius_guess"),
        }

    cats["measurement_temperatures"] = {"pixel_stats": pixel_stats}

    cats["pixel_data"] = {
        "shape": pixels_info.get("shape"),
        "dtype": pixels_info.get("dtype"),
        "ndim": pixels_info.get("ndim"),
        "raw_min_all": pixels_info.get("raw_min_all"),
        "raw_mean_all": pixels_info.get("raw_mean_all"),
        "raw_max_all": pixels_info.get("raw_max_all"),
        "radiometric_guess": pixels_info.get("radiometric_guess"),
        "looks_like_celsius_guess": pixels_info.get("looks_like_celsius_guess"),
    }

    cats["dji_xmp_simple_all"] = dji_simple

    def prune(obj: Any) -> Any:
        if isinstance(obj, dict):
            return {k: prune(v) for k, v in obj.items() if v not in (None, "", {}, [])}
        if isinstance(obj, list):
            return [prune(v) for v in obj if v not in (None, "", {}, [])]
        return obj

    return prune(cats)


# -----------------------------
# Google Maps + QR
# -----------------------------

def google_maps_link(lat: float, lon: float) -> str:
    return f"https://www.google.com/maps?q={lat:.7f},{lon:.7f}"


def make_qr_png(link: str, out_png: str) -> Tuple[Optional[str], Optional[str]]:
    if qrcode is None:
        return None, "qrcode library not available (pip install qrcode[pil])"
    try:
        img = qrcode.make(link)
        os.makedirs(os.path.dirname(out_png) or ".", exist_ok=True)
        img.save(out_png)
        return out_png, None
    except Exception as e:
        return None, f"QR generation failed: {e}"


# -----------------------------
# Main probe
# -----------------------------

@dataclass
class ProbeResult:
    file: str
    exiftool_available: bool
    exiftool_error: Optional[str]
    exiftool_meta: Dict[str, Any]
    pillow_error: Optional[str]
    pillow_exif: Dict[str, Any]
    pixels_error: Optional[str]
    pixels_info: Dict[str, Any]
    dji_xmp: Dict[str, Any]
    summary: Dict[str, Any]
    related_keys: Dict[str, Any]
    categories: Dict[str, Any]
    maps_link: Optional[str]
    qr_png: Optional[str]
    qr_error: Optional[str]


def probe_tiff(path: str) -> ProbeResult:
    exiftool_meta, exiftool_err = run_exiftool_json(path)
    exiftool_available = which_exiftool() is not None
    if exiftool_meta is None:
        exiftool_meta = {}

    pillow_exif, pillow_err = read_basic_exif_pillow(path)

    xml_packet = pillow_exif.get("XMLPacket")
    dji_xmp = parse_dji_xmp_from_xmlpacket(xml_packet) if xml_packet else {}
    dji_simple = dji_xmp.get("dji_simple", {}) if isinstance(dji_xmp, dict) else {}

    pixels_info, pixels_err = read_tiff_pixels(path)

    combined = dict(pillow_exif)
    combined.update(exiftool_meta)

    summary = build_summary(combined, dji_simple)

    patterns = [
        r"gps", r"latitude", r"longitude",
        r"serial", r"model", r"focal", r"fnumber", r"datetime|createdate|modifydate",
        r"emiss|emissivity",
        r"distance|subjectdistance",
        r"humidity|relativehumidity",
        r"ambient|atmos",
        r"wind",
        r"irradiance|solar",
        r"thermal|temperature|temp",
        r"radiometric|planck|r1|r2|b|f|o|tau",
        r"dji|mavic|drone",
        r"xmlpacket|xmp",
    ]
    related = find_matching_keys(combined, patterns)

    categories = build_categories(combined, dji_simple, pixels_info)

    return ProbeResult(
        file=os.path.abspath(path),
        exiftool_available=exiftool_available,
        exiftool_error=exiftool_err,
        exiftool_meta=exiftool_meta,
        pillow_error=pillow_err,
        pillow_exif=pillow_exif,
        pixels_error=pixels_err,
        pixels_info=pixels_info,
        dji_xmp=dji_xmp,
        summary=summary,
        related_keys=related,
        categories=categories,
        maps_link=None,
        qr_png=None,
        qr_error=None,
    )


def print_report(r: ProbeResult) -> None:
    print("\n==================== TIFF PROBE REPORT ====================")
    print("File:", r.file)

    print("\n--- ExifTool ---")
    print("ExifTool available:", r.exiftool_available)
    if r.exiftool_error:
        print("ExifTool error:", r.exiftool_error)
    else:
        print("ExifTool tags:", len(r.exiftool_meta))

    print("\n--- Pillow EXIF ---")
    if r.pillow_error:
        print("Pillow warning:", r.pillow_error)
    print("Pillow EXIF tags:", len(r.pillow_exif))

    print("\n--- Pixels (tifffile) ---")
    if r.pixels_error:
        print("Pixels warning:", r.pixels_error)
    else:
        for k, v in r.pixels_info.items():
            print(f"{k}: {v}")
        if "raw_min" in r.pixels_info:
            mn = float(r.pixels_info["raw_min"])
            av = float(r.pixels_info["raw_mean"])
            mx = float(r.pixels_info["raw_max"])
            print(f"raw MIN/AVG/MAX (pretty): {mn:.2f} / {av:.2f} / {mx:.2f}")
            if r.pixels_info.get("looks_like_celsius_guess"):
                print("NOTE: values look like °C (heuristic).")

    print("\n--- Summary (fields like report, if present) ---")
    for k, v in r.summary.items():
        if v is not None and v != "":
            print(f"{k}: {to_jsonable(v)}")

    print("\n--- Maps / QR ---")
    print("Google Maps link:", r.maps_link)
    print("QR PNG:", r.qr_png)
    if r.qr_error:
        print("QR note:", r.qr_error)

    print("\n--- DJI XMP (parsed from XMLPacket) ---")
    if not r.dji_xmp:
        print("(no DJI XMP parsed)")
    else:
        dji_simple = r.dji_xmp.get("dji_simple", {})
        print(f"Namespaces found: {len(r.dji_xmp.get('namespaces', {}))}")
        print(f"DJI simple tags: {len(dji_simple)}")
        preferred = [
            "AbsoluteAltitude", "RelativeAltitude",
            "GimbalYawDegree", "GimbalPitchDegree", "GimbalRollDegree",
            "FlightYawDegree", "FlightPitchDegree", "FlightRollDegree",
        ]
        shown = set()
        for k in preferred:
            if k in dji_simple:
                print(f"{k}: {dji_simple[k]}")
                shown.add(k)

        remaining = [(k, v) for k, v in dji_simple.items() if k not in shown]
        remaining.sort(key=lambda kv: kv[0].lower())
        max_more = 120
        if remaining:
            print("\nMore DJI tags:")
            for k, v in remaining[:max_more]:
                print(f"{k}: {v}")
            if len(remaining) > max_more:
                print(f"... ({len(remaining) - max_more} more DJI tags not shown)")

    print("\n--- Related keys found (grep-style) ---")
    if not r.related_keys:
        print("(none)")
    else:
        items = list(r.related_keys.items())
        max_show = 200
        for k, v in items[:max_show]:
            sv = str(to_jsonable(v))
            if len(sv) > 200:
                sv = sv[:200] + "…"
            print(f"{k}: {sv}")
        if len(items) > max_show:
            print(f"... ({len(items) - max_show} more keys not shown)")

    print("\n--- Categories (grouped output) ---")
    for cat, obj in r.categories.items():
        print(f"\n[{cat}]")
        print(json.dumps(to_jsonable(obj), ensure_ascii=False, indent=2))

    print("\n==========================================================\n")


def to_json_dict(r: ProbeResult) -> Dict[str, Any]:
    return {
        "file": r.file,
        "exiftool_available": r.exiftool_available,
        "exiftool_error": r.exiftool_error,
        "exiftool_meta": to_jsonable(r.exiftool_meta),
        "pillow_error": r.pillow_error,
        "pillow_exif": to_jsonable(r.pillow_exif),
        "pixels_error": r.pixels_error,
        "pixels_info": to_jsonable(r.pixels_info),
        "dji_xmp": to_jsonable(r.dji_xmp),
        "summary": to_jsonable(r.summary),
        "related_keys": to_jsonable(r.related_keys),
        "categories": to_jsonable(r.categories),
        "maps_link": r.maps_link,
        "qr_png": r.qr_png,
        "qr_error": r.qr_error,
    }


def main():
    ap = argparse.ArgumentParser(
        description="Probe a TIFF for EXIF/XMP/MakerNotes + DJI XMP + pixel stats + categorized output + optional QR"
    )
    ap.add_argument("tiff", help="Path to .tif/.tiff file")
    ap.add_argument("--out", help="Write full JSON report to this file (default: <tiff>.probe.json)", default=None)

    ap.add_argument("--qr", action="store_true", help="Generate QR code PNG for Google Maps link (if GPS exists)")
    ap.add_argument("--qr-out", default=None, help="Output path for QR PNG (default: <tiff>.maps.qr.png)")

    args = ap.parse_args()

    path = args.tiff
    if not os.path.isfile(path):
        raise SystemExit(f"File not found: {path}")

    r = probe_tiff(path)

    # QR generation (optional)
    if args.qr:
        geo = r.categories.get("geolocation", {}) if isinstance(r.categories, dict) else {}
        lat = geo.get("latitude")
        lon = geo.get("longitude")

        maps_link = None
        qr_path = None
        qr_err = None

        try:
            if lat is not None and lon is not None:
                maps_link = google_maps_link(float(lat), float(lon))
                qr_out = args.qr_out or (path + ".maps.qr.png")
                qr_path, qr_err = make_qr_png(maps_link, qr_out)
            else:
                qr_err = "No valid latitude/longitude found; QR not generated."
        except Exception as e:
            qr_err = f"QR generation exception: {e}"

        r.maps_link = maps_link
        r.qr_png = qr_path
        r.qr_error = qr_err

        r.categories.setdefault("maps", {})
        r.categories["maps"]["google_maps_link"] = maps_link
        r.categories["maps"]["qr_png"] = qr_path
        r.categories["maps"]["qr_error"] = qr_err

    print_report(r)

    out_path = args.out or (path + ".probe.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(to_json_dict(r), f, ensure_ascii=False, indent=2)
    print(f"Saved JSON report: {out_path}")


if __name__ == "__main__":
    main()
