import cv2
import numpy as np
import time
from flirimageextractor import FlirImageExtractor

# ====== ΡΥΘΜΙΣΗ: βάλε εδώ το αρχείο σου ======
PATH = r"F:\toposol_1.0\original-images-yes-no\park_2\yesD\DJI_20250804103534_0039_T_1.jpg"   # π.χ. FLIR radiometric JPG
# =============================================

# Clipboard via tkinter (built-in)
import tkinter as tk
root = tk.Tk()
root.withdraw()  # hide window

def copy_to_clipboard(text: str):
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()  # keeps clipboard after program ends (most OS)

flir = FlirImageExtractor()
flir.process_image(PATH)

# Thermal data σε °C (float array: height x width)
temp_c = flir.get_thermal_np()

# Display image (pseudo-color)
disp = np.nan_to_num(temp_c, nan=np.nanmin(temp_c))
disp_norm = cv2.normalize(disp, None, 0, 255, cv2.NORM_MINMAX).astype(np.uint8)
disp_color = cv2.applyColorMap(disp_norm, cv2.COLORMAP_INFERNO)

h, w = temp_c.shape
win = "Thermal Hover + Click Copy (Q/ESC to quit)"
cv2.namedWindow(win, cv2.WINDOW_NORMAL)

last_xy = (-1, -1)
copied_text = ""
copied_until = 0.0

def on_mouse(event, x, y, flags, param):
    global last_xy, copied_text, copied_until
    if event == cv2.EVENT_MOUSEMOVE:
        if 0 <= x < w and 0 <= y < h:
            last_xy = (x, y)

    if event == cv2.EVENT_LBUTTONDOWN:
        if 0 <= x < w and 0 <= y < h:
            t = float(temp_c[y, x])
            copied_text = f"{t:.2f} °C"
            copy_to_clipboard(copied_text)
            copied_until = time.time() + 1.2  # show "Copied!" for 1.2s

cv2.setMouseCallback(win, on_mouse)

while True:
    frame = disp_color.copy()

    x, y = last_xy
    if 0 <= x < w and 0 <= y < h:
        t = float(temp_c[y, x])

        # Crosshair
        cv2.drawMarker(frame, (x, y), (255, 255, 255),
                       markerType=cv2.MARKER_CROSS, markerSize=14, thickness=1)

        # Tooltip
        text = f"({x},{y})  {t:.2f} °C"
        (tw, th), _ = cv2.getTextSize(text, cv2.FONT_HERSHEY_SIMPLEX, 0.6, 2)
        bx, by = x + 10, y - 10
        if bx + tw + 10 > w: bx = x - tw - 20
        if by - th - 10 < 0: by = y + th + 20

        cv2.rectangle(frame, (bx, by - th - 8), (bx + tw + 8, by + 8), (0, 0, 0), -1)
        cv2.putText(frame, text, (bx + 4, by),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2, cv2.LINE_AA)

    # "Copied!" overlay
    if time.time() < copied_until:
        msg = f"Copied: {copied_text}"
        (mw, mh), _ = cv2.getTextSize(msg, cv2.FONT_HERSHEY_SIMPLEX, 0.7, 2)
        cv2.rectangle(frame, (10, 10), (10 + mw + 16, 10 + mh + 16), (0, 0, 0), -1)
        cv2.putText(frame, msg, (18, 10 + mh + 6),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2, cv2.LINE_AA)

    cv2.imshow(win, frame)
    key = cv2.waitKey(20) & 0xFF
    if key in (ord('q'), 27):  # q or ESC
        break

cv2.destroyAllWindows()
root.destroy()
