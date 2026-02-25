import cv2
import os



RGB_IMAGE_PATH = "F:\\project\\thermal1_images\\rgb\\DJI_20250310134904_0011_V.jpg"
RGB_LABEL_DIR  = "F:\\project\\thermal1_images\\rgb\\labels"

# label name = ίδιο basename με την εικόνα, .txt
base = os.path.splitext(os.path.basename(RGB_IMAGE_PATH))[0]
LABEL_PATH = os.path.join(RGB_LABEL_DIR, base + ".txt")

img = cv2.imread(RGB_IMAGE_PATH)
if img is None:
    raise FileNotFoundError(f"Δεν φορτώθηκε εικόνα: {RGB_IMAGE_PATH}")

h, w = img.shape[:2]
print("RGB size:", w, h)
print("Label path:", LABEL_PATH)

if not os.path.exists(LABEL_PATH):
    raise FileNotFoundError(f"Δεν βρέθηκε label: {LABEL_PATH}")

with open(LABEL_PATH, "r") as f:
    lines = [ln.strip() for ln in f if ln.strip()]

print("Lines in label:", len(lines))

draw = img.copy()

def clamp(v, lo, hi): 
    return max(lo, min(hi, v))

boxes_drawn = 0

for line in lines:
    parts = line.split()
    if len(parts) not in (5, 6):
        print("Skip (unexpected cols):", line)
        continue

    cls = parts[0]
    xc, yc, bw, bh = map(float, parts[1:5])
    conf = float(parts[5]) if len(parts) == 6 else None

    x_center = xc * w
    y_center = yc * h
    box_w = bw * w
    box_h = bh * h

    x_min = int(round(x_center - box_w/2))
    y_min = int(round(y_center - box_h/2))
    x_max = int(round(x_center + box_w/2))
    y_max = int(round(y_center + box_h/2))

    # clamp στα όρια εικόνας
    x_min = clamp(x_min, 0, w-1)
    x_max = clamp(x_max, 0, w-1)
    y_min = clamp(y_min, 0, h-1)
    y_max = clamp(y_max, 0, h-1)

    # αν κατέρρευσε
    if x_max <= x_min or y_max <= y_min:
        print("Invalid box after clamp:", line, "->", (x_min,y_min,x_max,y_max))
        continue

    # Draw bbox (παχύτερο)
    cv2.rectangle(draw, (x_min, y_min), (x_max, y_max), (0, 255, 0), 6)

    # Draw center cross
    cx = int(round(x_center))
    cy = int(round(y_center))
    cx = clamp(cx, 0, w-1)
    cy = clamp(cy, 0, h-1)
    cv2.drawMarker(draw, (cx, cy), (0, 255, 0), markerType=cv2.MARKER_CROSS, markerSize=30, thickness=4)

    label_txt = f"cls {cls}" + (f" conf {conf:.2f}" if conf is not None else "")
    cv2.putText(draw, label_txt, (x_min, max(30, y_min-10)),
                cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 255, 0), 3, cv2.LINE_AA)

    # Zoom crop γύρω από το κουτί (για να το δεις σίγουρα)
    pad = 200
    zx1 = clamp(x_min - pad, 0, w-1)
    zy1 = clamp(y_min - pad, 0, h-1)
    zx2 = clamp(x_max + pad, 0, w-1)
    zy2 = clamp(y_max + pad, 0, h-1)
    zoom = draw[zy1:zy2, zx1:zx2].copy()

    if zoom.size > 0:
        zoom = cv2.resize(zoom, (800, 800), interpolation=cv2.INTER_NEAREST)
        cv2.imshow("ZOOM (800x800)", zoom)

    print("BOX px:", (x_min, y_min, x_max, y_max), "from:", line)
    boxes_drawn += 1

# Save αποτέλεσμα
out_path = os.path.join(os.path.dirname(RGB_IMAGE_PATH), base + "_with_boxes.jpg")
cv2.imwrite(out_path, draw)
print("Saved:", out_path, "| boxes drawn:", boxes_drawn)

# Show full image resized για οθόνη
max_w = 1400
scale = min(max_w / w, 1.0)
show = cv2.resize(draw, (int(w*scale), int(h*scale)), interpolation=cv2.INTER_AREA) if scale < 1.0 else draw
cv2.namedWindow("RGB with boxes", cv2.WINDOW_NORMAL)
cv2.imshow("RGB with boxes", show)
cv2.waitKey(0)
cv2.destroyAllWindows()
