"""Pipeline tests using synthetic media — no browser, no real photos needed."""

import io
import os
import sys

import cv2
import numpy as np
import pytest
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import sticker_pipeline as sp


# ---------------------------------------------------------------------------
# Fixtures: a synthetic "person photo" and a synthetic video
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def person_photo() -> Image.Image:
    """A bright figure (circle head + rectangle body) on a contrasting background."""
    img = Image.new("RGB", (640, 800), (20, 90, 40))
    d = ImageDraw.Draw(img)
    d.ellipse((240, 120, 400, 280), fill=(250, 220, 180))          # head
    d.rectangle((220, 280, 420, 640), fill=(200, 40, 40))          # body
    return img


@pytest.fixture(scope="session")
def rembg_session():
    from rembg import new_session
    return new_session("u2netp")


@pytest.fixture(scope="session")
def cutout(person_photo, rembg_session) -> Image.Image:
    return sp.remove_background(person_photo, rembg_session)


@pytest.fixture(scope="session")
def video_path(tmp_path_factory) -> str:
    """30-frame mp4 where frame i is a flat gray level of (i*8) — decodable
    frames identify themselves through their mean pixel value."""
    path = str(tmp_path_factory.mktemp("vid") / "test.mp4")
    writer = cv2.VideoWriter(path, cv2.VideoWriter_fourcc(*"mp4v"), 30.0,
                             (320, 240))
    assert writer.isOpened()
    for i in range(30):
        frame = np.full((240, 320, 3), i * 8, dtype=np.uint8)
        writer.write(frame)
    writer.release()
    return path


# ---------------------------------------------------------------------------
# Image basics
# ---------------------------------------------------------------------------

def test_load_image_bytes_downscales_and_rgb():
    big = Image.new("RGB", (4000, 3000), (1, 2, 3))
    buf = io.BytesIO()
    big.save(buf, "JPEG")
    out = sp.load_image_bytes(buf.getvalue())
    assert out.mode == "RGB"
    assert max(out.size) <= sp.MAX_SOURCE_SIDE


def test_auto_enhance_keeps_shape(person_photo):
    out = sp.auto_enhance(person_photo)
    assert out.size == person_photo.size
    assert out.mode == "RGB"


def test_rotate_noop_and_real(person_photo):
    assert sp.rotate(person_photo, 0.0) is person_photo
    out = sp.rotate(person_photo, 10.0)
    assert out.size != person_photo.size  # expand=True grows the canvas


# ---------------------------------------------------------------------------
# Cutout pipeline
# ---------------------------------------------------------------------------

def test_cutout_has_alpha_variation(cutout):
    assert cutout.mode == "RGBA"
    alpha = np.asarray(cutout.split()[-1])
    assert (alpha > 200).any(), "no opaque subject pixels"
    assert (alpha < 50).any(), "no transparent background pixels"


def test_autocrop_tightens(cutout):
    cropped = sp.autocrop(cutout)
    assert cropped.width <= cutout.width
    assert cropped.height <= cutout.height
    alpha = np.asarray(cropped.split()[-1])
    assert (alpha > 8).any()


def test_add_border_grows_opaque_area(cutout):
    cropped = sp.autocrop(cutout)
    bordered = sp.add_border(cropped, 12)
    a0 = (np.asarray(cropped.split()[-1]) > 128).sum()
    a1 = (np.asarray(bordered.split()[-1]) > 128).sum()
    assert a1 > a0


def test_face_center_hint_no_face_returns_none(person_photo):
    # The synthetic figure has no facial features; Haar should find nothing
    # and the caller then keeps centered defaults.
    assert sp.face_center_hint(person_photo) is None


# ---------------------------------------------------------------------------
# Reference composition matching
# ---------------------------------------------------------------------------

def _tilted_subject(angle_deg: float, fill: float = 0.5) -> Image.Image:
    """RGBA with an opaque rectangle rotated by a known angle."""
    side = 400
    w, h = int(side * fill), int(side * fill * 1.6)
    block = Image.new("RGBA", (w, h), (255, 0, 0, 255))
    block = block.rotate(-angle_deg, resample=Image.BICUBIC, expand=True)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.alpha_composite(block, ((side - block.width) // 2,
                                   (side - block.height) // 2))
    return canvas


@pytest.mark.parametrize("angle", [0.0, 12.0, -18.0])
def test_reference_composition_recovers_tilt(angle):
    result = sp.reference_composition(_tilted_subject(angle))
    assert result is not None
    rotation, zoom, dx, dy = result
    assert rotation == pytest.approx(angle, abs=4.0)
    assert 0.80 <= zoom <= 0.95 / 0.92


def test_reference_composition_none_for_empty():
    empty = Image.new("RGBA", (200, 200), (0, 0, 0, 0))
    assert sp.reference_composition(empty) is None


def test_reference_composition_none_for_full_frame():
    full = Image.new("RGBA", (200, 200), (10, 10, 10, 255))
    assert sp.reference_composition(full) is None


def test_reference_composition_clamps_are_respected():
    # A 40-degree tilt folds into minAreaRect's ambiguity; whatever comes out
    # must still respect the guard rails the UI depends on.
    result = sp.reference_composition(_tilted_subject(40.0, fill=0.95))
    assert result is not None
    rotation, zoom, dx, dy = result
    assert abs(rotation) <= 25.0
    # Never suggests a crop: 0.95/0.92 zoom == subject at 95% of the canvas.
    assert 0.80 <= zoom <= 0.95 / 0.92
    assert abs(dx) <= 0.4 and abs(dy) <= 0.4


def test_reference_composition_reads_scale():
    """A subject filling more of its frame should suggest a larger zoom."""
    small = sp.reference_composition(_tilted_subject(0.0, fill=0.3))
    large = sp.reference_composition(_tilted_subject(0.0, fill=0.9))
    assert small is not None and large is not None
    assert large[1] > small[1]


# ---------------------------------------------------------------------------
# Stylization
# ---------------------------------------------------------------------------

def test_stylize_zero_is_noop(cutout):
    assert sp.stylize(cutout, 0) is cutout


def test_stylize_changes_pixels_and_keeps_alpha(cutout):
    out = sp.stylize(cutout, 100, warmth=(240, 200, 160))
    assert out.size == cutout.size
    assert out.mode == "RGBA"
    # Alpha must be untouched — styling may never eat the cutout mask.
    assert np.array_equal(np.asarray(out.split()[-1]),
                          np.asarray(cutout.split()[-1]))
    rgb_before = np.asarray(cutout.convert("RGB")).astype(int)
    rgb_after = np.asarray(out.convert("RGB")).astype(int)
    assert np.abs(rgb_after - rgb_before).sum() > 0


def test_stylize_strength_is_monotonic(cutout):
    base = np.asarray(cutout.convert("RGB")).astype(int)
    mild = np.abs(np.asarray(sp.stylize(cutout, 20).convert("RGB")).astype(int)
                  - base).sum()
    full = np.abs(np.asarray(sp.stylize(cutout, 100).convert("RGB")).astype(int)
                  - base).sum()
    assert full > mild


def test_stylize_handles_rgb_input(person_photo):
    out = sp.stylize(person_photo, 50)
    assert out.mode == "RGB"
    assert out.size == person_photo.size


# ---------------------------------------------------------------------------
# Composition / caption / export
# ---------------------------------------------------------------------------

def test_compose_canvas_is_512(cutout):
    out = sp.compose_canvas(sp.autocrop(cutout))
    assert out.size == (sp.CANVAS_SIZE, sp.CANVAS_SIZE)
    assert out.mode == "RGBA"


def test_compose_canvas_bg_fill(cutout):
    out = sp.compose_canvas(sp.autocrop(cutout), bg=(255, 0, 0))
    corner = out.getpixel((2, 2))
    assert corner == (255, 0, 0, 255)


def test_dominant_color_reads_ring():
    ref = Image.new("RGB", (200, 200), (10, 20, 200))
    d = ImageDraw.Draw(ref)
    d.ellipse((60, 60, 140, 140), fill=(255, 255, 0))  # center content ignored
    r, g, b = sp.dominant_color(ref)
    assert (r, g, b) == (10, 20, 200)


def test_draw_caption_renders_pixels(cutout):
    canvas = sp.compose_canvas(sp.autocrop(cutout))
    font = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    captioned = sp.draw_caption(canvas, "HELLO", position="bottom",
                                font_path=font if os.path.exists(font) else None)
    diff = np.asarray(captioned).astype(int) - np.asarray(canvas).astype(int)
    assert np.abs(diff).sum() > 0, "caption drew nothing"
    assert sp.draw_caption(canvas, "  ", position="top").tobytes() == canvas.tobytes()


def test_export_webp_under_limit(cutout):
    out = sp.compose_canvas(sp.autocrop(sp.add_border(cutout, 8)))
    data, quality = sp.export_webp_max(out)
    assert len(data) <= sp.WEBP_BYTE_LIMIT
    reopened = Image.open(io.BytesIO(data))
    assert reopened.format == "WEBP"
    assert reopened.mode in ("RGBA", "RGB")
    assert "A" in reopened.mode  # alpha preserved


def test_export_png_roundtrip(cutout):
    out = sp.compose_canvas(sp.autocrop(cutout))
    data = sp.export_png(out)
    reopened = Image.open(io.BytesIO(data))
    assert reopened.format == "PNG"
    assert reopened.size == (sp.CANVAS_SIZE, sp.CANVAS_SIZE)


# ---------------------------------------------------------------------------
# Video
# ---------------------------------------------------------------------------

def test_video_meta(video_path):
    meta = sp.get_video_meta(video_path)
    assert meta is not None
    assert meta.fps == pytest.approx(30.0, abs=1.0)
    assert meta.duration_ms == pytest.approx(1000.0, rel=0.2)  # 30 frames @ 30fps


def test_read_frame_by_time(video_path):
    # Frame at ~500ms should be frame ~15 -> gray level ~120.
    img = sp.read_frame(video_path, 500.0)
    assert img is not None
    mean = np.asarray(img).mean()
    assert 96 <= mean <= 144, f"unexpected frame content (mean={mean})"


def test_read_frame_past_end_returns_last(video_path):
    img = sp.read_frame(video_path, 99_000.0)
    assert img is not None
    mean = np.asarray(img).mean()
    assert mean > 180  # late frames are bright (i*8 -> up to 232)


def test_sharpest_frame_time_runs(video_path):
    meta = sp.get_video_meta(video_path)
    t = sp.sharpest_frame_time(video_path, meta, samples=10)
    assert 0.0 <= t <= meta.duration_ms
