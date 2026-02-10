from pathlib import Path
import cv2

# ================= ΡΥΘΜΙΣΕΙΣ =================
INPUT_DIR = Path(r"F:\toposol_1.0\original-images-yes-no\park_edafos_2\yesF")
OUTPUT_DIR = Path(r"F:\toposol_1.0\original-images-yes-no\park_edafos_2\yesF_gray")


IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}
# =============================================


def is_image(p: Path) -> bool:
    return p.suffix.lower() in IMG_EXTS


def convert_image(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)

    img = cv2.imread(str(src), cv2.IMREAD_UNCHANGED)
    if img is None:
        print(f"[WARN] Δεν διαβάστηκε: {src}")
        return

    # Αν είναι ήδη grayscale
    if len(img.shape) == 2:
        gray = img
    else:
        # BGR / BGRA → GRAY
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Gray → 3-channel (H,W,3)
    gray3 = cv2.merge([gray, gray, gray])

    ok = cv2.imwrite(str(dst), gray3)
    if not ok:
        print(f"[WARN] Δεν γράφτηκε: {dst}")


def main():
    if not INPUT_DIR.exists():
        raise RuntimeError(f"Δεν υπάρχει ο φάκελος: {INPUT_DIR}")

    count = 0
    for img_path in INPUT_DIR.rglob("*"):
        if not img_path.is_file() or not is_image(img_path):
            continue

        rel = img_path.relative_to(INPUT_DIR)
        out_path = OUTPUT_DIR / rel

        convert_image(img_path, out_path)
        count += 1

    print(f"✔ Ολοκληρώθηκε μετατροπή {count} εικόνων σε grayscale (3-channel).")


if __name__ == "__main__":
    main()
