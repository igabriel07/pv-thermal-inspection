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

Χρήση:
  python check_for_metadata_tiff.py "C:\\path\\file.tiff"
  python check_for_metadata_tiff.py "C:\\path\\file.tiff" --out report.json

Απαιτήσεις:
  pip install pillow tifffile numpy
Προαιρετικά (συνιστάται):
  εγκατάσταση ExifTool (ώστε να πάρεις GPS/Distance/Emissivity κλπ πιο εύκολα)
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

import xml.etree.ElementTree as ET


# -----------------------------
# JSON-safe conversion
# -----------------------------

def to_jsonable(x: Any) -> Any:
    """Convert common non-JSON types (Pillow IFDRational, bytes, numpy) to JSON-safe."""
    if x is None:
        return None

    # numpy scalars
    if isinstance(x, (np.integer, np.floating, np.bool_)):
        return x.item()

    # numpy arrays
    if isinstance(x, np.ndarray):
        return x.tolist()

    # bytes -> try decode (utf-16le often used in XP*), else hex
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

    # PIL IFDRational or Fraction-like
    if hasattr(x, "numerator") and hasattr(x, "denominator"):
        try:
            n = int(x.numerator)
            d = int(x.denominator)
            if d != 0:
                return n / d
            return None
        except Exception:
            pass

    # tuples/lists
    if isinstance(x, (list, tuple)):
        return [to_jsonable(v) for v in x]

    # dict
    if isinstance(x, dict):
        return {str(k): to_jsonable(v) for k, v in x.items()}

    # primitives
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
    Επιστρέφει (metadata_dict, error_message)
    metadata_dict: dict tags όπως τα δίνει το exiftool -j -G1 -a -s -u -ee
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
    """
    dms: (deg, min, sec) where each may be rational-like
    """
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
    """
    Decode GPSInfo IFD (numeric keys) into decimal lat/lon if possible.
    """
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

    out = {
        "gps_raw": {k: to_jsonable(v) for k, v in gps_tags.items()}
    }
    if lat is not None:
        out["latitude"] = lat
    if lon is not None:
        out["longitude"] = lon
    return out


def read_basic_exif_pillow(path: str) -> Tuple[Dict[str, Any], Optional[str]]:
    """
    Απλό EXIF μέσω Pillow (περιορισμένο). Σε DJI thermal TIFF, πολλά vendor πράγματα
    εμφανίζονται στο XMLPacket (XMP) και όχι ως κλασικά EXIF tags.
    """
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

        # GPS decoding only if GPSInfo is a dict; in your file it is often an offset number.
        if "GPSInfo" in tag_map and isinstance(tag_map["GPSInfo"], dict):
            tag_map["GPSDecoded"] = decode_gpsinfo_if_possible(tag_map["GPSInfo"])

        return tag_map, None
    except Exception as e:
        return {}, f"Pillow EXIF read error: {e}"


# -----------------------------
# DJI XMP parsing from XMLPacket
# -----------------------------

def _try_extract_xml_fragment(text: str) -> Optional[str]:
    """
    Find <x:xmpmeta ...> ... </x:xmpmeta> fragment.
    """
    start = text.find("<x:xmpmeta")
    if start == -1:
        start = text.find("<xmpmeta")  # sometimes without prefix
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
    """
    Pillow sometimes returns XMLPacket as:
    - bytes (ideal)
    - str with mojibake (as you saw)
    We try multiple decode/repair approaches.
    """
    if xmlpacket is None:
        return None

    # If bytes: try decode as UTF-8 / UTF-16LE / UTF-16BE
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

    # If str: attempt to "recover" to bytes and decode as utf-16le (common for XMP in TIFF)
    if isinstance(xmlpacket, str):
        # Approach 1: latin-1 encode to preserve byte values 0-255 then decode as utf-16le
        try:
            raw = xmlpacket.encode("latin-1", errors="ignore")
            for enc in ("utf-16le", "utf-16be", "utf-8"):
                t = raw.decode(enc, errors="ignore")
                frag = _try_extract_xml_fragment(t)
                if frag:
                    return frag
        except Exception:
            pass

        # Approach 2: direct search in the string (maybe already contains readable XML)
        frag = _try_extract_xml_fragment(xmlpacket)
        if frag:
            return frag

    return None


def parse_dji_xmp_from_xmlpacket(xmlpacket: Any) -> Dict[str, Any]:
    """
    Parse DJI-related tags from XMP packet.

    Returns:
      {
        "namespaces": {...},
        "tags": {"drone-dji:AbsoluteAltitude": "...", ...},
        "dji_simple": {"AbsoluteAltitude": "...", ...},
        "gps_from_xmp": {"latitude": ..., "longitude": ...} if found
      }
    """
    xml_text = _decode_xmlpacket_to_text(xmlpacket)
    if not xml_text:
        return {}

    # xml.etree requires valid XML; sometimes there are control chars; strip them lightly.
    cleaned = "".join(ch for ch in xml_text if ch == "\n" or ch == "\t" or (ord(ch) >= 32))
    try:
        root = ET.fromstring(cleaned)
    except Exception:
        # Try one more time with a looser cleanup
        cleaned2 = re.sub(r"[^\x09\x0A\x0D\x20-\x7E\u0080-\uFFFF]", "", xml_text)
        try:
            root = ET.fromstring(cleaned2)
        except Exception:
            return {}

    # Collect namespaces from the root string (quick heuristic)
    ns = {}
    for m in re.finditer(r'xmlns:([A-Za-z0-9_\-]+)="([^"]+)"', cleaned):
        ns[m.group(1)] = m.group(2)

    tags: Dict[str, Any] = {}
    dji_simple: Dict[str, Any] = {}

    # Iterate all elements and record those that contain 'drone-dji' in tag namespace URI or prefix
    for elem in root.iter():
        tag = elem.tag  # can be '{uri}local' or 'prefix:local' depending on parser
        text = (elem.text or "").strip() if elem.text else ""

        # Case A: {uri}local
        if tag.startswith("{"):
            uri, local = tag[1:].split("}", 1)
            # match uri that looks like DJI
            if "dji" in uri.lower() and ("drone" in uri.lower() or "dji.com" in uri.lower()):
                key = f"{{{uri}}}{local}"
                tags[key] = text
                dji_simple[local] = text

        # Case B: prefix:local
        if ":" in tag and not tag.startswith("{"):
            prefix, local = tag.split(":", 1)
            if prefix.lower() in ("drone-dji", "dji", "djidrone", "dronedji"):
                key = f"{prefix}:{local}"
                tags[key] = text
                dji_simple[local] = text

    # Try to infer GPS from XMP (common names: RtkStdLat/RtkStdLon, Latitude/Longitude, etc.)
    gps_from_xmp = {}
    # check a bunch of likely keys
    lat_candidates = ["Latitude", "lat", "RtkStdLat", "GpsLatitude", "GPSLatitude", "RTKStdLat"]
    lon_candidates = ["Longitude", "lon", "RtkStdLon", "GpsLongitude", "GPSLongitude", "RTKStdLon"]

    def _find_first_float(keys: List[str]) -> Optional[float]:
        for k in keys:
            if k in dji_simple:
                try:
                    return float(str(dji_simple[k]).strip())
                except Exception:
                    continue
        return None

    lat = _find_first_float(lat_candidates)
    lon = _find_first_float(lon_candidates)
    if lat is not None:
        gps_from_xmp["latitude"] = lat
    if lon is not None:
        gps_from_xmp["longitude"] = lon

    out = {
        "namespaces": ns,
        "tags": tags,
        "dji_simple": dji_simple,
    }
    if gps_from_xmp:
        out["gps_from_xmp"] = gps_from_xmp
    return out


# -----------------------------
# Pixels
# -----------------------------

def read_tiff_pixels(path: str) -> Tuple[Dict[str, Any], Optional[str]]:
    if tifffile is None:
        return {}, "tifffile not available (pip install tifffile)"
    try:
        arr = tifffile.imread(path)

        info: Dict[str, Any] = {
            "shape": list(arr.shape),
            "dtype": str(arr.dtype),
            "ndim": int(arr.ndim),
        }

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
# Report-like extraction
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


def build_summary(meta: Dict[str, Any], dji_simple: Dict[str, Any]) -> Dict[str, Any]:
    s = {}
    s["Make"] = pick_first(meta, ["EXIF:Make", "Make"])
    s["Model"] = pick_first(meta, ["EXIF:Model", "Model"])
    s["CameraSerialNumber"] = pick_first(meta, ["EXIF:CameraSerialNumber", "CameraSerialNumber", "SerialNumber", "BodySerialNumber"])
    s["FocalLength"] = pick_first(meta, ["EXIF:FocalLength", "FocalLength", "Focal Length"])
    s["FNumber"] = pick_first(meta, ["EXIF:FNumber", "FNumber", "F-Number", "F Number"])
    s["ImageWidth"] = pick_first(meta, ["File:ImageWidth", "ImageWidth", "EXIF:ExifImageWidth"])
    s["ImageHeight"] = pick_first(meta, ["File:ImageHeight", "ImageLength", "ImageHeight", "EXIF:ExifImageHeight"])
    s["DateTimeOriginal"] = pick_first(meta, ["EXIF:DateTimeOriginal", "DateTimeOriginal", "Date/Time Original"])
    s["DateTime"] = pick_first(meta, ["EXIF:DateTime", "DateTime", "Date Time"])
    s["Software"] = pick_first(meta, ["Software", "EXIF:Software"])
    s["ImageDescription"] = pick_first(meta, ["ImageDescription", "EXIF:ImageDescription"])

    # GPS: from Pillow-decoded GPS if available
    gps_decoded = meta.get("GPSDecoded")
    if isinstance(gps_decoded, dict):
        s["Latitude"] = gps_decoded.get("latitude")
        s["Longitude"] = gps_decoded.get("longitude")

    # GPS: from DJI XMP if found
    if "Latitude" not in s or s["Latitude"] is None:
        if "gps_from_xmp" in dji_simple:
            # (not used here)
            pass

    # Likely DJI flight fields (nice to have)
    for key in ("AbsoluteAltitude", "RelativeAltitude", "GimbalYawDegree", "GimbalPitchDegree", "GimbalRollDegree",
                "FlightYawDegree", "FlightPitchDegree", "FlightRollDegree"):
        if key in dji_simple and dji_simple[key] not in (None, ""):
            s[f"DJI_{key}"] = dji_simple[key]

    # thermal/measurement params (may or may not exist)
    s["Emissivity"] = pick_first(meta, ["Emissivity", "XMP:Emissivity", "EXIF:Emissivity"])
    s["Distance"] = pick_first(meta, ["Distance", "SubjectDistance", "EXIF:SubjectDistance"])
    s["AmbientTemperature"] = pick_first(meta, ["AmbientTemperature", "Ambient Temperature"])
    s["RelativeHumidity"] = pick_first(meta, ["RelativeHumidity", "Relative Humidity", "Humidity"])
    s["WindSpeed"] = pick_first(meta, ["WindSpeed", "Wind Speed"])
    s["Irradiance"] = pick_first(meta, ["Irradiance", "SolarIrradiance", "Solar Irradiance"])

    return s


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
    dji_xmp: Dict[str, Any]          # parsed output (namespaces/tags/dji_simple)
    summary: Dict[str, Any]
    related_keys: Dict[str, Any]


def probe_tiff(path: str) -> ProbeResult:
    # 1) ExifTool
    exiftool_meta, exiftool_err = run_exiftool_json(path)
    exiftool_available = which_exiftool() is not None
    if exiftool_meta is None:
        exiftool_meta = {}

    # 2) Pillow
    pillow_exif, pillow_err = read_basic_exif_pillow(path)

    # 3) DJI XMP from XMLPacket
    xml_packet = pillow_exif.get("XMLPacket")
    dji_xmp = parse_dji_xmp_from_xmlpacket(xml_packet) if xml_packet else {}
    dji_simple = dji_xmp.get("dji_simple", {}) if isinstance(dji_xmp, dict) else {}

    # 4) Pixels
    pixels_info, pixels_err = read_tiff_pixels(path)

    # Combine meta (priority to exiftool)
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

    print("\n--- DJI XMP (parsed from XMLPacket) ---")
    if not r.dji_xmp:
        print("(no DJI XMP parsed)")
    else:
        dji_simple = r.dji_xmp.get("dji_simple", {})
        print(f"Namespaces found: {len(r.dji_xmp.get('namespaces', {}))}")
        print(f"DJI simple tags: {len(dji_simple)}")
        # show a curated subset first
        preferred = [
            "AbsoluteAltitude", "RelativeAltitude",
            "GimbalYawDegree", "GimbalPitchDegree", "GimbalRollDegree",
            "FlightYawDegree", "FlightPitchDegree", "FlightRollDegree",
            "RtkStdLat", "RtkStdLon", "Latitude", "Longitude",
        ]
        shown = set()
        for k in preferred:
            if k in dji_simple:
                print(f"{k}: {dji_simple[k]}")
                shown.add(k)

        # then show the rest (limited)
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
    }


def main():
    ap = argparse.ArgumentParser(description="Probe a TIFF for EXIF/XMP/MakerNotes + DJI XMP + pixel stats")
    ap.add_argument("tiff", help="Path to .tif/.tiff file")
    ap.add_argument("--out", help="Write full JSON report to this file (default: <tiff>.probe.json)", default=None)
    args = ap.parse_args()

    path = args.tiff
    if not os.path.isfile(path):
        raise SystemExit(f"File not found: {path}")

    r = probe_tiff(path)
    print_report(r)

    out_path = args.out or (path + ".probe.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(to_json_dict(r), f, ensure_ascii=False, indent=2)
    print(f"Saved JSON report: {out_path}")


if __name__ == "__main__":
    main()
