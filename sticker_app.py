"""Personal Sticker Maker — every sticker starts from a fresh photo of you.

Flow: pick a reference sticker → the app asks for a photo/video of yourself
matching it → local cutout (no cloud AI, no API keys) → style → download.

Run locally:  streamlit run sticker_app.py
Deploy:       Streamlit Community Cloud, entrypoint sticker_app.py
"""

from __future__ import annotations

import io
import os
import tempfile

import streamlit as st
from PIL import Image

import sticker_pipeline as sp

FONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "assets", "DejaVuSans-Bold.ttf")

st.set_page_config(page_title="Personal Sticker Maker", page_icon="🎟️",
                   layout="centered")


# ---------------------------------------------------------------------------
# Cached heavy lifting
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner="Warming up the cutout engine (first launch only)…")
def get_rembg_session():
    from rembg import new_session
    return new_session("u2netp")


@st.cache_data(show_spinner="Cutting you out of the background…", max_entries=8)
def cutout_cached(source_png: bytes, refine_edges: bool) -> bytes:
    img = Image.open(io.BytesIO(source_png)).convert("RGB")
    rgba = sp.remove_background(img, get_rembg_session(),
                                refine_edges=refine_edges)
    rgba = sp.autocrop(rgba)
    buf = io.BytesIO()
    rgba.save(buf, "PNG")
    return buf.getvalue()


@st.cache_data(show_spinner="Reading the reference's framing…", max_entries=4)
def reference_composition_cached(ref_png: bytes):
    """Cut out the *reference's* subject to learn its tilt/scale/position."""
    img = Image.open(io.BytesIO(ref_png)).convert("RGB")
    rgba = sp.remove_background(img, get_rembg_session())
    return sp.reference_composition(rgba)


def to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

ss = st.session_state
ss.setdefault("nonce", 0)
ss.setdefault("gallery", [])          # [{"name", "webp", "png"}]
ss.setdefault("source_bytes", None)   # PNG bytes of the chosen photo/frame
ss.setdefault("video_tmp_path", None)
ss.setdefault("video_meta", None)
ss.setdefault("sticker_count", 0)

STYLE_DEFAULTS = {
    "auto_enhance": False,
    "rotation_deg": 0.0,
    "refine_edges": False,
    "stickerness": 35,
    "border_px": 12,
    "bg_mode": "Transparent",
    "bg_color": "#FFFFFF",
    "zoom": 1.0,
    "dx": 0.0,
    "dy": 0.0,
    "caption_text": "",
    "caption_pos": "Bottom",
}
for key, value in STYLE_DEFAULTS.items():
    ss.setdefault(key, value)


def set_source(png_bytes: bytes) -> None:
    """A new photo/frame was chosen: store it and reset per-photo style state."""
    if png_bytes == ss.source_bytes:
        return
    ss.source_bytes = png_bytes
    for key, value in STYLE_DEFAULTS.items():
        ss[key] = value
    # Face-aware default composition (manual sliders stay as the override).
    hint = sp.face_center_hint(Image.open(io.BytesIO(png_bytes)))
    if hint is not None:
        ss.zoom, ss.dx, ss.dy = hint


def reset_wizard() -> None:
    ss.nonce += 1
    ss.source_bytes = None
    ss.video_meta = None
    if ss.video_tmp_path and os.path.exists(ss.video_tmp_path):
        os.remove(ss.video_tmp_path)
    ss.video_tmp_path = None
    for key, value in STYLE_DEFAULTS.items():
        ss[key] = value


nonce = ss.nonce

# ---------------------------------------------------------------------------
# Header + eager model warm-up
# ---------------------------------------------------------------------------

st.title("🎟️ Personal Sticker Maker")
st.caption("Every sticker starts from a **fresh photo of you** — pick a "
           "reference, strike its pose, get your sticker. Nothing is "
           "auto-generated, nothing leaves this app.")

get_rembg_session()  # download + load the 4.7 MB model up front, not mid-flow

# ---------------------------------------------------------------------------
# Step 1 — Reference sticker
# ---------------------------------------------------------------------------

st.header("1 · Pick your reference sticker")
ref_file = st.file_uploader(
    "Upload the sticker/meme/template whose vibe you want to copy",
    type=["png", "jpg", "jpeg", "webp"], key=f"ref_{nonce}")

ref_img = None
if ref_file is not None:
    ref_img = sp.load_image_bytes(ref_file.getvalue())

if ref_img is None:
    st.info("Start by uploading a reference sticker — a template, meme, or "
            "any sticker whose style or pose you want to recreate as *you*.")
    st.stop()

# ---------------------------------------------------------------------------
# Step 2 — Your photo or video, matching the reference
# ---------------------------------------------------------------------------

st.header("2 · Now, a photo or video of you")
left, right = st.columns([1, 2])
with left:
    st.image(ref_img, caption="Your reference — match this pose/vibe",
             width="stretch")
with right:
    st.markdown("Give me a photo or video of yourself **matching this "
                "sticker** — same pose, same energy. That's what keeps every "
                "sticker personal.")
    st.caption("📸 Tip: plain background + good lighting = clean cutout.")

tab_upload, tab_camera, tab_video = st.tabs(
    ["⬆️ Upload photo", "🤳 Camera", "🎬 From video"])

with tab_upload:
    st.caption("Recommended on phones — your camera app gives better shots "
               "than the browser camera.")
    photo_file = st.file_uploader("Photo of you", type=["png", "jpg", "jpeg", "webp"],
                                  key=f"photo_{nonce}")
    if photo_file is not None:
        set_source(to_png_bytes(sp.load_image_bytes(photo_file.getvalue())))

with tab_camera:
    shot = st.camera_input("Strike the reference pose", key=f"cam_{nonce}")
    if shot is not None:
        set_source(to_png_bytes(sp.load_image_bytes(shot.getvalue())))

with tab_video:
    video_file = st.file_uploader("Video of you (I'll help you pick the frame)",
                                  type=["mp4", "mov", "webm", "m4v"],
                                  key=f"video_{nonce}")
    if video_file is not None:
        data = video_file.getvalue()
        if ss.video_tmp_path is None or ss.get("video_sig") != (video_file.name, len(data)):
            suffix = os.path.splitext(video_file.name)[1] or ".mp4"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(data)
            if ss.video_tmp_path and os.path.exists(ss.video_tmp_path):
                os.remove(ss.video_tmp_path)
            ss.video_tmp_path = tmp.name
            ss.video_sig = (video_file.name, len(data))
            ss.video_meta = sp.get_video_meta(tmp.name)
            ss.pop("frame_s", None)

        meta = ss.video_meta
        if meta is None:
            st.error("Couldn't read that video — try an mp4 or a different clip.")
        else:
            duration_s = max(round(meta.duration_ms / 1000.0, 1), 0.1)
            frame_step_s = meta.frame_ms / 1000.0
            ss.setdefault("frame_s", min(0.5, duration_s / 2))

            def nudge(delta_s: float) -> None:
                ss.frame_s = float(min(max(ss.frame_s + delta_s, 0.0), duration_s))

            def suggest_sharpest() -> None:
                t_ms = sp.sharpest_frame_time(ss.video_tmp_path, meta)
                ss.frame_s = float(min(round(t_ms / 1000.0, 2), duration_s))

            st.slider("Scrub to your moment (seconds)", 0.0, duration_s,
                      step=0.1, key="frame_s")
            b1, b2, b3 = st.columns(3)
            b1.button("◀ 1 frame", on_click=nudge, args=(-frame_step_s,),
                      width="stretch")
            b2.button("1 frame ▶", on_click=nudge, args=(frame_step_s,),
                      width="stretch")
            b3.button("✨ Sharpest frame", on_click=suggest_sharpest,
                      width="stretch",
                      help="Auto-picks the least blurry frame — you can still scrub.")

            frame = sp.read_frame(ss.video_tmp_path, ss.frame_s * 1000.0)
            if frame is None:
                st.error("Couldn't decode a frame at that position.")
            else:
                st.image(frame, caption=f"Frame at {ss.frame_s:.1f}s",
                         width="stretch")
                if st.button("✅ Use this frame", type="primary",
                             width="stretch"):
                    set_source(to_png_bytes(frame))
                    st.rerun()

if ss.source_bytes is None:
    st.stop()

# ---------------------------------------------------------------------------
# Step 3 — Cutout & style
# ---------------------------------------------------------------------------

st.header("3 · Style your sticker")

def apply_reference_composition() -> None:
    """Pose the sticker like the reference — pre-sets sliders, never forces."""
    suggestion = reference_composition_cached(to_png_bytes(ref_img))
    if suggestion is None:
        ss.compose_msg = ("Couldn't read a clear subject in that reference — "
                          "sliders left as they were.")
        return
    ss.rotation_deg, ss.zoom, ss.dx, ss.dy = suggestion
    ss.compose_msg = "ok"


st.button("🎯 Match reference composition", on_click=apply_reference_composition,
          width="stretch",
          help="Tilts, scales and positions you like the reference. "
               "The sliders below move — nudge anything that looks off.")
if ss.get("compose_msg") == "ok":
    st.caption("Matched the reference's framing — tweak the sliders below to taste.")
elif ss.get("compose_msg"):
    st.warning(ss.compose_msg)

with st.expander("🛠️ Photo fixes", expanded=False):
    st.checkbox("Auto-enhance (fix dim / flat lighting)", key="auto_enhance")
    st.slider("Straighten / tilt (°)", -25.0, 25.0, step=0.5, key="rotation_deg")
    st.checkbox("Refine edges (slower, cleaner hair)", key="refine_edges")

source_img = Image.open(io.BytesIO(ss.source_bytes)).convert("RGB")
if ss.auto_enhance:
    source_img = sp.auto_enhance(source_img)
if abs(ss.rotation_deg) >= 0.05:
    source_img = sp.rotate(source_img, ss.rotation_deg)

cutout_png = cutout_cached(to_png_bytes(source_img), ss.refine_edges)
cutout = Image.open(io.BytesIO(cutout_png)).convert("RGBA")

st.slider("✨ Sticker-ness", 0, 100, key="stickerness",
          help="Smooths, warms and flattens your photo toward illustration so "
               "it reads as a sticker instead of a snapshot. 0 = untouched.")

c1, c2 = st.columns(2)
with c1:
    st.slider("Sticker border", 0, 24, key="border_px")
    st.radio("Background", ["Transparent", "Color", "Match reference"],
             key="bg_mode", horizontal=True)
    if ss.bg_mode == "Color":
        st.color_picker("Background color", key="bg_color")
with c2:
    st.slider("Zoom", 0.5, 2.0, step=0.05, key="zoom")
    st.slider("Move ↔", -0.4, 0.4, step=0.02, key="dx")
    st.slider("Move ↕", -0.4, 0.4, step=0.02, key="dy")

st.text_input("Caption (optional)", key="caption_text",
              placeholder="LOL / same / on my way…")
if ss.caption_text.strip():
    st.radio("Caption position", ["Top", "Bottom"], key="caption_pos",
             horizontal=True)

if ss.bg_mode == "Transparent":
    bg = None
elif ss.bg_mode == "Match reference":
    bg = sp.dominant_color(ref_img)
else:
    h = ss.bg_color.lstrip("#")
    bg = tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

subject = sp.stylize(cutout, ss.stickerness, warmth=sp.dominant_color(ref_img))
styled = sp.add_border(subject, ss.border_px)
sticker = sp.compose_canvas(styled, zoom=ss.zoom, dx=ss.dx, dy=ss.dy, bg=bg)
sticker = sp.draw_caption(
    sticker, ss.caption_text, position=ss.caption_pos.lower(),
    font_path=FONT_PATH if os.path.exists(FONT_PATH) else None)

p1, p2 = st.columns(2)
with p1:
    st.image(ref_img, caption="Reference", width="stretch")
with p2:
    st.image(sticker, caption="Your sticker", width="stretch")

# ---------------------------------------------------------------------------
# Step 4 — Export & use
# ---------------------------------------------------------------------------

st.header("4 · Save & use it")

webp_bytes, webp_quality = sp.export_webp_max(sticker)
png_bytes = sp.export_png(sticker)
name = f"sticker_{ss.sticker_count + 1}"

d1, d2 = st.columns(2)
d1.download_button(
    f"⬇️ WebP · {len(webp_bytes) / 1024:.0f} KB",
    webp_bytes, f"{name}.webp", "image/webp", width="stretch",
    help="512×512 WebP — the format sticker apps and Telegram expect.")
d2.download_button(
    f"⬇️ PNG · {len(png_bytes) / 1024:.0f} KB",
    png_bytes, f"{name}.png", "image/png", width="stretch",
    help="Transparent PNG — send it in any chat like an image.")

with st.expander("💬 How do I get this into WhatsApp / Telegram?"):
    st.markdown(
        "- **Telegram** — send the **WebP** file to the official "
        "[@Stickers](https://t.me/stickers) bot to build your own pack, or "
        "just send it in a chat.\n"
        "- **WhatsApp** — WhatsApp has no direct sticker import; add the WebP "
        "via any free *sticker maker* app (e.g. \"Sticker Maker\" on the "
        "Play Store / App Store), or simply send the **PNG** as an image.\n"
        "- **Anywhere else** — the PNG with transparency works in most chats.")

g1, g2 = st.columns(2)
if g1.button("📌 Add to this session's gallery", width="stretch"):
    ss.sticker_count += 1
    ss.gallery.append({"name": f"sticker_{ss.sticker_count}",
                       "webp": webp_bytes, "png": png_bytes})
    st.toast("Added to gallery ✔")
g2.button("🔄 Make another sticker", on_click=reset_wizard, width="stretch",
          type="primary")

# ---------------------------------------------------------------------------
# Gallery (session-only — stickers are meant to be sent right away)
# ---------------------------------------------------------------------------

if ss.gallery:
    st.divider()
    st.subheader(f"🖼️ This session's stickers ({len(ss.gallery)})")
    st.caption("Gallery lives for this session only — download what you love.")
    cols = st.columns(4)
    for i, item in enumerate(ss.gallery):
        with cols[i % 4]:
            st.image(item["png"], width="stretch")
            st.download_button("⬇️", item["webp"], f"{item['name']}.webp",
                               "image/webp", key=f"dl_{i}", width="stretch")
