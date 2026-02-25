import os
import json

# Set this to your metadata directory
METADATA_DIR = r'F:\project\thermal1_images\thermal\faults\metadata'  # <-- update as needed
CALIBRATION_CUTOFF = 'DJI_20250310135130_0129_V.JPG'

def extract_jpg_name(json_path):
    base = os.path.basename(json_path)
    if base.endswith('_T.JPG.json'):
        return base.replace('_T.JPG.json', '_V.JPG')
    return base

def round_deg(val):
    try:
        return int(round(float(val)))
    except Exception:
        return None

def extract_yaws(meta):
    def get(*keys):
        for k in keys:
            if k in meta:
                return meta[k]
        return None
    fyaw = get('XMP-drone-dji:FlightYawDegree')
    gyaw = get('XMP-drone-dji:GimbalYawDegree')
    return round_deg(fyaw), round_deg(gyaw)

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

    def bucket(f):
        try:
            with open(os.path.join(METADATA_DIR, f), encoding='utf-8') as jf:
                meta = json.load(jf)
            fy, gy = extract_yaws(meta.get('exiftool_meta', {}))
            return (fy, gy)
        except Exception:
            return (None, None)

    buckets_cal = set(bucket(f) for f in calibrated if bucket(f)[0] is not None and bucket(f)[1] is not None)
    buckets_uncal = set(bucket(f) for f in uncalibrated if bucket(f)[0] is not None and bucket(f)[1] is not None)

    print(f"Total calibrated images: {len(calibrated)}")
    print(f"Total uncalibrated images: {len(uncalibrated)}")
    print(f"Calibrated yaw buckets: {len(buckets_cal)}")
    print(f"Uncalibrated yaw buckets: {len(buckets_uncal)}")
    print()
    print("Buckets present in uncalibrated images but not in calibrated set:")
    missing = buckets_uncal - buckets_cal
    # Map from bucket to list of image names
    bucket_to_images = {}
    for f in uncalibrated:
        b = bucket(f)
        if b in missing:
            bucket_to_images.setdefault(b, []).append(extract_jpg_name(f))
    for fy, gy in sorted(missing):
        imgs = bucket_to_images.get((fy, gy), [])
        img_list = ', '.join(imgs) if imgs else '(no image found)'
        print(f"  flight_yaw_360={fy}, gimbal_yaw_360={gy} | images: {img_list}")
    if not missing:
        print("  (None! All yaw buckets are covered by your calibrations.)")

if __name__ == '__main__':
    main()
