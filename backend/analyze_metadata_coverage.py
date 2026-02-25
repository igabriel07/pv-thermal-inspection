import os
import json
from collections import defaultdict

# Set this to your metadata directory
METADATA_DIR = r'F:\project\thermal1_images\thermal\faults\metadata'
CALIBRATION_CUTOFF = 'DJI_20250310135130_0129_V.JPG'

# Bucketing granularity (degrees/meters)
YAW_BUCKET = 1
PITCH_BUCKET = 5
ALT_BUCKET = 5

def bucket(val, step):
    try:
        return int(round(float(val) / step) * step)
    except Exception:
        return None

def extract_fields(meta):
    def get(*keys):
        for k in keys:
            if k in meta:
                return meta[k]
        return None
    yaw = get('XMP-drone-dji:FlightYawDegree')
    gimbal = get('XMP-drone-dji:GimbalYawDegree')
    pitch = get('XMP-drone-dji:GimbalPitchDegree')
    alt = get('XMP-drone-dji:RelativeAltitude')
    return yaw, gimbal, pitch, alt

def extract_jpg_name(json_path):
    base = os.path.basename(json_path)
    if base.endswith('_T.JPG.json'):
        return base.replace('_T.JPG.json', '_V.JPG')
    return base

def main():
    files = [f for f in os.listdir(METADATA_DIR) if f.endswith('.json')]
    files.sort()
    calibrated = []
    uncalibrated = []
    cutoff_found = False
    for f in files:
        jpg = extract_jpg_name(f)
        if not cutoff_found:
            calibrated.append(f)
            if CALIBRATION_CUTOFF in jpg:
                cutoff_found = True
        else:
            uncalibrated.append(f)

    def get_bucket(f):
        try:
            with open(os.path.join(METADATA_DIR, f), encoding='utf-8') as jf:
                meta = json.load(jf)
            yaw, gimbal, pitch, alt = extract_fields(meta.get('exiftool_meta', {}))
            byaw = bucket(yaw, YAW_BUCKET)
            bgimbal = bucket(gimbal, YAW_BUCKET)
            bpitch = bucket(pitch, PITCH_BUCKET)
            balt = bucket(alt, ALT_BUCKET)
            return (byaw, bgimbal, bpitch, balt)
        except Exception as e:
            return None

    buckets_cal = set(get_bucket(f) for f in calibrated if get_bucket(f) is not None)
    buckets_uncal = set(get_bucket(f) for f in uncalibrated if get_bucket(f) is not None)

    print(f"Total calibrated images: {len(calibrated)}")
    print(f"Total uncalibrated images: {len(uncalibrated)}")
    print(f"Calibrated (yaw,gimbal,pitch,alt) buckets: {len(buckets_cal)}")
    print(f"Uncalibrated buckets: {len(buckets_uncal)}")
    print()
    print("Buckets present in uncalibrated images but not in calibrated set:")
    missing = buckets_uncal - buckets_cal
    for b in sorted(missing):
        print(f"  flight_yaw={b[0]}, gimbal_yaw={b[1]}, gimbal_pitch={b[2]}, rel_alt={b[3]}")
    if not missing:
        print("  (None! All buckets are covered by your calibrations.)")

if __name__ == '__main__':
    main()
