import cv2
import numpy as np

DEFAULT_RUST_RANGES_HSV = [
    ((5,  60,  40), (30, 255, 255)),
    ((0,  40,  20), (20, 255, 200)),
]

def analyze_rust_bgr(
    img_bgr: np.ndarray,
    rust_ranges_hsv=DEFAULT_RUST_RANGES_HSV,
    exclude_dark_pixels: bool = True,
    min_v_for_valid: int = 35,
    kernel_size: int = 5,
    open_iters: int = 1,
    close_iters: int = 2,
    # --- New stronger criteria knobs ---
    use_clahe: bool = True,
    clahe_clip: float = 2.0,
    clahe_grid: int = 8,
    use_texture_gate: bool = True,
    texture_thr: int = 35,
    min_blob_area: int = 200,
):
    """
    Returns:
      rust_pct, rust_pixels, valid_pixels, mask_final (uint8 0/255), overlay_bgr
    """

    if img_bgr is None or img_bgr.size == 0:
        raise ValueError("Empty image passed to analyze_rust_bgr")

    h, w = img_bgr.shape[:2]

    # -----------------------------
    # 1) Illumination normalization (optional)
    # -----------------------------
    if use_clahe:
        lab = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)
        clahe = cv2.createCLAHE(clipLimit=float(clahe_clip), tileGridSize=(int(clahe_grid), int(clahe_grid)))
        l2 = clahe.apply(l)
        lab2 = cv2.merge([l2, a, b])
        img_norm = cv2.cvtColor(lab2, cv2.COLOR_LAB2BGR)
    else:
        img_norm = img_bgr

    img_hsv = cv2.cvtColor(img_norm, cv2.COLOR_BGR2HSV)

    # -----------------------------
    # 2) Rust color mask (HSV ranges)
    # -----------------------------
    mask_rust = np.zeros((h, w), dtype=np.uint8)
    for lo, hi in rust_ranges_hsv:
        lo = np.array(lo, dtype=np.uint8)
        hi = np.array(hi, dtype=np.uint8)
        mask_rust = cv2.bitwise_or(mask_rust, cv2.inRange(img_hsv, lo, hi))

    # -----------------------------
    # 3) Valid area mask (exclude dark pixels/shadows)
    # -----------------------------
    if exclude_dark_pixels:
        v = img_hsv[:, :, 2]
        valid = (v >= int(min_v_for_valid)).astype(np.uint8) * 255
    else:
        valid = np.ones((h, w), dtype=np.uint8) * 255

    # -----------------------------
    # 4) Morphology cleanup (remove noise, fill small gaps)
    # -----------------------------
    ksz = max(1, int(kernel_size))
    if ksz % 2 == 0:
        ksz += 1  # keep odd kernel size
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ksz, ksz))

    mask_clean = cv2.morphologyEx(mask_rust, cv2.MORPH_OPEN, kernel, iterations=int(open_iters))
    mask_clean = cv2.morphologyEx(mask_clean, cv2.MORPH_CLOSE, kernel, iterations=int(close_iters))

    # Start with color+valid
    mask_final = cv2.bitwise_and(mask_clean, valid)

    # -----------------------------
    # 5) Texture gate (optional): rust tends to be rough
    # -----------------------------
    if use_texture_gate:
        gray = cv2.cvtColor(img_norm, cv2.COLOR_BGR2GRAY)
        gx = cv2.Sobel(gray, cv2.CV_32F, 1, 0, ksize=3)
        gy = cv2.Sobel(gray, cv2.CV_32F, 0, 1, ksize=3)
        grad = cv2.magnitude(gx, gy)
        grad = cv2.normalize(grad, None, 0, 255, cv2.NORM_MINMAX).astype(np.uint8)

        thr = int(texture_thr)
        mask_texture = (grad >= thr).astype(np.uint8) * 255

        mask_final = cv2.bitwise_and(mask_final, mask_texture)

    # -----------------------------
    # 6) Remove small specks (connected components)
    # -----------------------------
    mba = int(min_blob_area)
    if mba > 1:
        num, labels, stats, _ = cv2.connectedComponentsWithStats(mask_final, connectivity=8)
        mask_filtered = np.zeros_like(mask_final)
        for i in range(1, num):
            if stats[i, cv2.CC_STAT_AREA] >= mba:
                mask_filtered[labels == i] = 255
        mask_final = mask_filtered

    # -----------------------------
    # 7) Compute rust %
    # -----------------------------
    rust_pixels = int(np.count_nonzero(mask_final))
    valid_pixels = int(np.count_nonzero(valid))
    rust_pct = 100.0 * rust_pixels / max(valid_pixels, 1)

    # -----------------------------
    # 8) Overlay for visualization
    # -----------------------------
    overlay = img_norm.copy()
    overlay[mask_final > 0] = (0, 0, 255)  # red
    alpha = 0.45
    overlay_bgr = cv2.addWeighted(img_norm, 1 - alpha, overlay, alpha, 0)

    return rust_pct, rust_pixels, valid_pixels, mask_final, overlay_bgr
