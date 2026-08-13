"""Pure image/video processing functions for the Personal Sticker Maker.

No streamlit imports here — everything is unit-testable in isolation.
All images flow through as PIL Images (RGB for sources, RGBA for cutouts).
"""

from __future__ import annotations

import io
from dataclasses import dataclass

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageOps

CANVAS_SIZE = 512
WEBP_BYTE_LIMIT = 100_000
MAX_SOURCE_SIDE = 1600  # RAM guard: downscale huge phone photos before processing


# ---------------------------------------------------------------------------
# Loading / basic adjustments
# ---------------------------------------------------------------------------

def load_image_bytes(data: bytes) -> Image.Image:
    """Decode uploaded image bytes -> RGB PIL image, EXIF-rotated, size-capped."""
    img = Image.open(io.BytesIO(data))
    img = ImageOps.exif_transpose(img)
    img = img.convert("RGB")
    if max(img.size) > MAX_SOURCE_SIDE:
        img.thumbnail((MAX_SOURCE_SIDE, MAX_SOURCE_SIDE), Image.LANCZOS)
    return img


def auto_enhance(img: Image.Image) -> Image.Image:
    """CLAHE on the luminance channel + gentle auto-contrast.

    Fixes the common phone-photo problems (dim indoor light, flat contrast)
    deterministically, without any model.
    """
    arr = cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2LAB)
    l_chan, a_chan, b_chan = cv2.split(arr)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    l_chan = clahe.apply(l_chan)
    arr = cv2.cvtColor(cv2.merge((l_chan, a_chan, b_chan)), cv2.COLOR_LAB2RGB)
    out = Image.fromarray(arr)
    return ImageOps.autocontrast(out, cutoff=1)


def rotate(img: Image.Image, degrees: float) -> Image.Image:
    """Rotate around center, expanding the canvas; fills corners with edge color."""
    if abs(degrees) < 0.05:
        return img
    return img.rotate(-degrees, resample=Image.BICUBIC, expand=True,
                      fillcolor=(255, 255, 255))


# ---------------------------------------------------------------------------
# Video handling
# ---------------------------------------------------------------------------

@dataclass
class VideoMeta:
    duration_ms: float
    fps: float

    @property
    def frame_ms(self) -> float:
        return 1000.0 / self.fps if self.fps > 0 else 33.3


def get_video_meta(path: str) -> VideoMeta | None:
    """Probe duration/fps. Handles containers (webm) that report no frame count."""
    cap = cv2.VideoCapture(path)
    if not cap.isOpened():
        return None
    try:
        fps = cap.get(cv2.CAP_PROP_FPS) or 0.0
        frame_count = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0.0
        if fps > 0 and frame_count > 0:
            return VideoMeta(duration_ms=frame_count / fps * 1000.0, fps=fps)
        # Fallback (typical for webm): binary-search-free capped scan by msec.
        # Seek far ahead and walk back until a frame decodes.
        for probe_ms in (600_000, 300_000, 120_000, 60_000, 30_000, 10_000,
                         5_000, 2_000, 1_000, 500, 100):
            cap.set(cv2.CAP_PROP_POS_MSEC, probe_ms)
            ok, _ = cap.read()
            if ok:
                end_ms = cap.get(cv2.CAP_PROP_POS_MSEC)
                return VideoMeta(duration_ms=max(end_ms, probe_ms),
                                 fps=fps if fps > 0 else 30.0)
        # Last resort: count frames (capped) sequentially.
        cap.set(cv2.CAP_PROP_POS_MSEC, 0)
        count = 0
        while count < 108_000:  # cap: 1h at 30fps
            ok, _ = cap.read()
            if not ok:
                break
            count += 1
        if count == 0:
            return None
        use_fps = fps if fps > 0 else 30.0
        return VideoMeta(duration_ms=count / use_fps * 1000.0, fps=use_fps)
    finally:
        cap.release()


def read_frame(path: str, time_ms: float) -> Image.Image | None:
    """Decode the single frame at (or nearest before) time_ms."""
    cap = cv2.VideoCapture(path)
    if not cap.isOpened():
        return None
    try:
        cap.set(cv2.CAP_PROP_POS_MSEC, max(0.0, time_ms))
        ok, frame = cap.read()
        if not ok:
            # Past the end — grab the final decodable frame instead.
            cap.set(cv2.CAP_PROP_POS_MSEC, 0)
            last = None
            while True:
                ok, f = cap.read()
                if not ok:
                    break
                last = f
            if last is None:
                return None
            frame = last
        img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        if max(img.size) > MAX_SOURCE_SIDE:
            img.thumbnail((MAX_SOURCE_SIDE, MAX_SOURCE_SIDE), Image.LANCZOS)
        return img
    finally:
        cap.release()


def sharpest_frame_time(path: str, meta: VideoMeta, samples: int = 30) -> float:
    """Pick the sharpest of `samples` evenly spaced frames (Laplacian variance).

    Returns the timestamp in ms. Downscales each probe frame so the whole scan
    stays well under a second on CPU.
    """
    cap = cv2.VideoCapture(path)
    if not cap.isOpened():
        return 0.0
    best_ms, best_score = 0.0, -1.0
    try:
        for i in range(samples):
            t = meta.duration_ms * (i + 0.5) / samples
            cap.set(cv2.CAP_PROP_POS_MSEC, t)
            ok, frame = cap.read()
            if not ok:
                continue
            h, w = frame.shape[:2]
            if max(h, w) > 480:
                scale = 480 / max(h, w)
                frame = cv2.resize(frame, (int(w * scale), int(h * scale)))
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            score = cv2.Laplacian(gray, cv2.CV_64F).var()
            if score > best_score:
                best_ms, best_score = t, score
    finally:
        cap.release()
    return best_ms


# ---------------------------------------------------------------------------
# Cutout
# ---------------------------------------------------------------------------

def remove_background(img: Image.Image, session, refine_edges: bool = False) -> Image.Image:
    """rembg cutout -> RGBA. `session` is a rembg session (cached by the app)."""
    from rembg import remove

    kwargs = {"session": session, "post_process_mask": True}
    if refine_edges:
        kwargs.update(alpha_matting=True,
                      alpha_matting_foreground_threshold=240,
                      alpha_matting_background_threshold=15,
                      alpha_matting_erode_size=10)
    try:
        out = remove(img, **kwargs)
    except Exception:
        if not refine_edges:
            raise
        # pymatting can fail on degenerate trimaps — fall back to the plain cutout.
        kwargs = {"session": session, "post_process_mask": True}
        out = remove(img, **kwargs)
    return out.convert("RGBA")


def autocrop(rgba: Image.Image, pad_frac: float = 0.06) -> Image.Image:
    """Crop to the subject's alpha bounding box with a small padding margin."""
    alpha = np.asarray(rgba.split()[-1])
    ys, xs = np.where(alpha > 8)
    if len(xs) == 0:
        return rgba
    x0, x1, y0, y1 = xs.min(), xs.max(), ys.min(), ys.max()
    pad = int(max(x1 - x0, y1 - y0) * pad_frac)
    x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
    x1, y1 = min(rgba.width - 1, x1 + pad), min(rgba.height - 1, y1 + pad)
    return rgba.crop((x0, y0, x1 + 1, y1 + 1))


def add_border(rgba: Image.Image, border_px: int) -> Image.Image:
    """Classic white sticker outline: dilate the alpha mask, feather its edge,
    and composite the subject over the resulting white shape."""
    if border_px <= 0:
        return rgba
    # Grow the canvas so the outline never clips.
    grow = border_px + 4
    padded = Image.new("RGBA", (rgba.width + 2 * grow, rgba.height + 2 * grow),
                       (0, 0, 0, 0))
    padded.paste(rgba, (grow, grow), rgba)

    alpha = np.asarray(padded.split()[-1])
    kernel = cv2.getStructuringElement(
        cv2.MORPH_ELLIPSE, (2 * border_px + 1, 2 * border_px + 1))
    dilated = cv2.dilate(alpha, kernel)
    # Feather just the outline edge so it isn't jagged.
    dilated = cv2.GaussianBlur(dilated, (5, 5), 0)

    outline = Image.new("RGBA", padded.size, (255, 255, 255, 0))
    outline.putalpha(Image.fromarray(dilated))
    outline = Image.composite(
        Image.new("RGBA", padded.size, (255, 255, 255, 255)),
        Image.new("RGBA", padded.size, (0, 0, 0, 0)),
        Image.fromarray(dilated))
    outline.alpha_composite(padded)
    return outline


def face_center_hint(img: Image.Image) -> tuple[float, float, float] | None:
    """Detect the largest face -> suggested (zoom, dx, dy) for compose_canvas.

    Returns None when no face is found (caller keeps centered defaults).
    Uses the Haar cascade bundled inside opencv — no download, no model file.
    """
    cascade_path = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
    cascade = cv2.CascadeClassifier(cascade_path)
    if cascade.empty():
        return None
    gray = cv2.cvtColor(np.asarray(img.convert("RGB")), cv2.COLOR_RGB2GRAY)
    faces = cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5,
                                     minSize=(40, 40))
    if len(faces) == 0:
        return None
    x, y, w, h = max(faces, key=lambda f: f[2] * f[3])
    face_cx = (x + w / 2) / img.width
    face_cy = (y + h / 2) / img.height
    # Nudge the composition so the face sits in the upper-middle of the canvas.
    dx = 0.5 - face_cx
    dy = 0.38 - face_cy
    return (1.0, float(np.clip(dx, -0.3, 0.3)), float(np.clip(dy, -0.3, 0.3)))


def reference_composition(ref_rgba: Image.Image) -> tuple[float, float, float, float] | None:
    """Read the reference's framing off its own cutout mask.

    Returns (rotation_deg, zoom, dx, dy) shaped for the app's sliders, so the
    user's sticker can be posed the same way as the reference. Returns None
    when the reference has no usable subject (empty or full-frame mask).
    """
    alpha = np.asarray(ref_rgba.split()[-1])
    mask = (alpha > 8).astype(np.uint8)
    total = mask.size
    covered = int(mask.sum())
    # No subject, or the cutout kept everything (nothing to learn from).
    if covered < total * 0.01 or covered > total * 0.995:
        return None

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    biggest = max(contours, key=cv2.contourArea)

    # Tilt: minAreaRect's angle is ambiguous mod 90 — fold it into [-45, 45]
    # and keep only plausible sticker tilts.
    (_cx, _cy), (rw, rh), angle = cv2.minAreaRect(biggest)
    if rw < rh:
        angle += 90.0
    angle = (angle + 45.0) % 90.0 - 45.0
    rotation = float(np.clip(angle, -25.0, 25.0))

    # Scale: how much of its frame the reference subject fills, expressed as a
    # zoom relative to compose_canvas' default 92% fit.
    # Floor keeps a small-subject reference from suggesting a uselessly tiny
    # sticker; ceiling keeps the subject inside 95% of the canvas so a
    # frame-filling reference never crops the user's head off.
    x, y, w, h = cv2.boundingRect(biggest)
    fill = max(w / ref_rgba.width, h / ref_rgba.height)
    zoom = float(np.clip(fill / 0.92, 0.80, 0.95 / 0.92))

    # Offset: where the subject sits inside its frame.
    dx = float(np.clip(0.5 - (x + w / 2) / ref_rgba.width, -0.4, 0.4))
    dy = float(np.clip(0.5 - (y + h / 2) / ref_rgba.height, -0.4, 0.4))
    return rotation, zoom, dx, dy


def stylize(img: Image.Image, strength: int,
            warmth: tuple[int, int, int] | None = None) -> Image.Image:
    """Push a photo toward illustration so it reads as a sticker, not a snapshot.

    strength: 0 = untouched, 100 = full effect. Blended with the original so
    the result never goes fully waxy. `warmth` (an RGB color sampled from the
    reference) pulls the palette toward the reference's mood.

    Preserves size, mode and the alpha channel exactly — styling must never
    eat the cutout mask.
    """
    if strength <= 0:
        return img
    amount = min(max(strength, 0), 100) / 100.0
    has_alpha = img.mode == "RGBA"
    alpha_channel = img.split()[-1] if has_alpha else None
    rgb = np.asarray(img.convert("RGB"))

    # 1. Edge-preserving smoothing: cleans skin without destroying features.
    styled = cv2.edgePreservingFilter(rgb, flags=cv2.RECURS_FILTER,
                                      sigma_s=60, sigma_r=0.35)
    # 2. Mild posterize toward flat illustration tones (capped so faces
    #    don't band into blotches).
    levels = 10
    styled = (styled // (256 // levels) * (256 // levels) + (256 // levels) // 2)
    styled = np.clip(styled, 0, 255).astype(np.uint8)
    # 3. Saturation + contrast lift — the "sticker pop".
    hsv = cv2.cvtColor(styled, cv2.COLOR_RGB2HSV).astype(np.float32)
    hsv[..., 1] *= 1.35
    hsv = np.clip(hsv, 0, 255).astype(np.uint8)
    styled = cv2.cvtColor(hsv, cv2.COLOR_HSV2RGB)
    styled = np.clip((styled.astype(np.float32) - 128) * 1.12 + 128, 0, 255)
    # 4. Nudge the palette toward the reference's mood.
    if warmth is not None:
        target = np.array(warmth, dtype=np.float32)
        styled = styled * 0.88 + (styled * (target / max(target.mean(), 1.0))) * 0.12
        styled = np.clip(styled, 0, 255)

    blended = (rgb.astype(np.float32) * (1 - amount)
               + styled.astype(np.float32) * amount)
    out = Image.fromarray(np.clip(blended, 0, 255).astype(np.uint8))
    if has_alpha:
        out = out.convert("RGBA")
        out.putalpha(alpha_channel)
    return out


# ---------------------------------------------------------------------------
# Styling / composition
# ---------------------------------------------------------------------------

def dominant_color(img: Image.Image) -> tuple[int, int, int]:
    """Dominant border-region color of the reference image (its 'background')."""
    rgb = img.convert("RGB")
    small = rgb.resize((64, 64))
    arr = np.asarray(small).reshape(-1, 3)
    # Sample only the outer ring — that's where a sticker's background lives.
    mask = np.zeros((64, 64), dtype=bool)
    mask[:8, :] = mask[-8:, :] = mask[:, :8] = mask[:, -8:] = True
    ring = np.asarray(small)[mask]
    if len(ring) == 0:
        ring = arr
    # Median is robust to captions/logos in the ring.
    r, g, b = np.median(ring, axis=0).astype(int)
    return int(r), int(g), int(b)


def compose_canvas(rgba: Image.Image, size: int = CANVAS_SIZE, zoom: float = 1.0,
                   dx: float = 0.0, dy: float = 0.0,
                   bg: tuple[int, int, int] | None = None) -> Image.Image:
    """Fit the subject onto a size×size canvas.

    zoom: 1.0 fits the subject to ~92% of the canvas; >1 zooms in.
    dx/dy: offsets as a fraction of the canvas (-0.5..0.5).
    bg: None = transparent, else an RGB fill color.
    """
    canvas = Image.new(
        "RGBA", (size, size),
        (0, 0, 0, 0) if bg is None else (bg[0], bg[1], bg[2], 255))
    if rgba.width == 0 or rgba.height == 0:
        return canvas
    fit = 0.92 * zoom
    scale = min(size * fit / rgba.width, size * fit / rgba.height)
    w = max(1, int(rgba.width * scale))
    h = max(1, int(rgba.height * scale))
    subject = rgba.resize((w, h), Image.LANCZOS)
    px = int((size - w) / 2 + dx * size)
    py = int((size - h) / 2 + dy * size)
    canvas.alpha_composite(subject, (max(-w + 1, min(size - 1, px)),
                                     max(-h + 1, min(size - 1, py))))
    return canvas


def draw_caption(canvas: Image.Image, text: str, position: str = "bottom",
                 font_path: str | None = None) -> Image.Image:
    """Meme-style caption: bold white text with a black stroke, top or bottom."""
    text = (text or "").strip()
    if not text:
        return canvas
    out = canvas.copy()
    draw = ImageDraw.Draw(out)
    size = out.width // 8
    font = None
    while size >= 12:
        try:
            font = (ImageFont.truetype(font_path, size)
                    if font_path else ImageFont.load_default(size=size))
        except OSError:
            font = ImageFont.load_default(size=size)
        bbox = draw.textbbox((0, 0), text, font=font,
                             stroke_width=max(2, size // 12))
        if bbox[2] - bbox[0] <= out.width * 0.94:
            break
        size = int(size * 0.85)
    stroke = max(2, size // 12)
    bbox = draw.textbbox((0, 0), text, font=font, stroke_width=stroke)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (out.width - tw) // 2 - bbox[0]
    y = (int(out.height * 0.02) - bbox[1] if position == "top"
         else out.height - int(out.height * 0.02) - th - bbox[1])
    draw.text((x, y), text, font=font, fill=(255, 255, 255, 255),
              stroke_width=stroke, stroke_fill=(0, 0, 0, 255))
    return out


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def export_webp_max(img: Image.Image, limit: int = WEBP_BYTE_LIMIT) -> tuple[bytes, int]:
    """WebP with alpha under `limit` bytes: step quality down, then downscale."""
    for canvas_size in (img.width, 448, 384):
        frame = (img if canvas_size == img.width
                 else img.resize((canvas_size, canvas_size), Image.LANCZOS))
        for quality in range(90, 25, -10):
            buf = io.BytesIO()
            frame.save(buf, "WEBP", quality=quality, method=6)
            data = buf.getvalue()
            if len(data) <= limit:
                return data, quality
    return data, quality  # smallest attempt, even if over limit


def export_png(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, "PNG", optimize=True)
    return buf.getvalue()
