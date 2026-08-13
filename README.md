# aitp_ufc

Two independent Streamlit apps live in this repo:

| App | Entrypoint | What it does |
|---|---|---|
| 🎟️ Personal Sticker Maker | `sticker_app.py` | Turn a fresh photo/video of you into a chat sticker, guided by a reference sticker |
| 📝 Doc-to-MD Converter | `app.py` | Convert PDF/Word/Excel/PowerPoint/HTML to Markdown |

## 🎟️ Personal Sticker Maker

Every sticker is personal, made from a **fresh photo of you each time** — never
auto-generated from a stored avatar.

1. **Pick a reference sticker** — any template/meme/sticker whose style, pose, or vibe you want.
2. **The app asks for a photo or video of you** matching it — upload, browser camera, or scrub a video to the exact frame (with a "sharpest frame" auto-suggest).
3. **Local cutout + styling** — background removal runs on a tiny on-device model (`rembg`/u2netp, no cloud AI, no API keys), then: white sticker border, auto-enhance, straighten, face-aware centering, background fill (including one-click "match reference background"), and meme-style captions.
   - **🎯 Match reference composition** — reads the reference's own tilt, scale, and position off its cutout mask and poses you the same way, so the two actually look like the same sticker. It pre-sets the sliders rather than forcing a transform, so you can nudge anything that looks off.
   - **✨ Sticker-ness slider** — smooths, warms, and flattens your photo toward illustration so it reads as a sticker instead of a snapshot. The color target is sampled from your reference, so its palette pulls your sticker toward matching it. 0 = untouched.
4. **Download** — 512×512 WebP (under 100 KB, sticker-app ready) and transparent PNG.

**Getting stickers into chats:** Telegram — send the WebP to the official
[@Stickers](https://t.me/stickers) bot. WhatsApp — import the WebP with any
free "sticker maker" app, or just send the PNG as an image.

### Run locally

```bash
pip install -r requirements.txt
streamlit run sticker_app.py
```

First launch downloads the ~5 MB cutout model; after that everything is offline.

### Deploy free on Streamlit Community Cloud

1. Push this repo to GitHub (already done if you're reading this there).
2. On [share.streamlit.io](https://share.streamlit.io), create an app from this repo.
3. Set **Main file path** to `sticker_app.py` and, under *Advanced settings*, Python **3.12**.

The app sleeps when idle on the free tier — the first visit after a while
takes ~30–60 s to wake up and warm the model. That's normal.

## 📝 Doc-to-MD Converter

```bash
streamlit run app.py
```

## Tests

```bash
pip install pytest
pytest -q
```

The tests exercise the whole image/video pipeline with synthetic media — no
real photos needed.
