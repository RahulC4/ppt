# ============================================================
# generate_ppt_auto.py – TITLE + AGENDA + AUTO SIZE + IMAGES + COLOR
# ============================================================

import os
import tempfile
import uuid
import json
import re
import base64

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

from utils import (
    get_env, logger, now_ts,
    ensure_dir, text_client, image_client
)
from search_utils import semantic_search
from azure_blob_utils import upload_ppt_to_blob, upload_json_to_blob

ensure_dir("generated")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)

# ✅ CORPORATE COLORS
CORP_BLUE = RGBColor(0, 84, 165)
BULLET_BLACK = RGBColor(0, 0, 0)
BG_COLOR = RGBColor(242, 246, 250)  # Clean light background


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def parse_user_intent(prompt: str):
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    return int(match.group(1)) if match else None


def fallback_plan(n):
    slides = []
    for i in range(n):
        slides.append({
            "title": f"Slide {i+1}",
            "bullets": [
                f"Key takeaway {i+1}",
                f"Supporting detail {i+1}.1",
                f"Supporting detail {i+1}.2",
            ]
        })
    return slides


# ------------------------------------------------------------
# ✅ LLM (TEXT MODE)
# ------------------------------------------------------------
def call_llm_plan_auto(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return EXACTLY in this format:\n\n"
        "Slide 1: Title\n"
        "- Bullet\n"
        "- Bullet\n"
        "- Bullet\n\n"
        f"Create EXACTLY {num_slides} slides."
    )

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": prompt},
            ],
            max_completion_tokens=1400,
            temperature=1,
        )

        raw = resp.choices[0].message.content.strip()
        return parse_text_plan(raw, num_slides)

    except Exception as e:
        logger.warning(f"LLM failed → fallback used: {e}")
        return fallback_plan(num_slides)


def parse_text_plan(text, num_slides):
    slides = []
    blocks = re.split(r"\n(?=Slide \d+:)", text)

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue

        title = lines[0].split(":", 1)[-1].strip()
        bullets = [l.replace("-", "").strip() for l in lines[1:] if l.startswith("-")]

        slides.append({
            "title": title,
            "bullets": bullets[:6],
        })

    return slides[:num_slides]


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=prompt + " professional minimal illustration, no text",
            size="1024x1024",
        )

        b64 = getattr(resp.data[0], "b64_json", None)
        if not b64:
            return None

        img_bytes = base64.b64decode(b64)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.write(img_bytes)
        tmp.close()
        return tmp.name

    except Exception:
        logger.exception("Image generation failed")
        return None


# ------------------------------------------------------------
# ✅✅✅ PPT BUILDER WITH TITLE + AGENDA ✅✅✅
# ------------------------------------------------------------
def build_ppt(slides, ppt_title, image_required):
    prs = Presentation()

    # ---------- ✅ TITLE SLIDE (NO IMAGE) ----------
    title_slide = prs.slides.add_slide(prs.slide_layouts[1])
    title_slide.shapes.title.text = ppt_title

    body = title_slide.placeholders[1]
    body.text_frame.text = "Auto-generated presentation"
    body.text_frame.paragraphs[0].font.size = Pt(18)

    fill = title_slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # ---------- ✅ AGENDA SLIDE (IMAGE ALLOWED) ----------
    agenda_slide = prs.slides.add_slide(prs.slide_layouts[1])
    agenda_slide.shapes.title.text = "Agenda"

    fill = agenda_slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    agenda_body = agenda_slide.placeholders[1]
    tf = agenda_body.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = True
    tf.clear()

    for sp in slides:
        p = tf.add_paragraph()
        p.text = sp["title"]
        p.font.size = Pt(20)
        p.font.color.rgb = BULLET_BLACK

    # Optional agenda image
    if image_required:
        img = generate_visual_image("Presentation agenda overview")
        if img:
            agenda_body.width = prs.slide_width - Inches(4)
            agenda_body.left = Inches(0.5)
            agenda_slide.shapes.add_picture(
                img,
                prs.slide_width - Inches(3.5),
                agenda_body.top,
                width=Inches(3),
            )

    # ---------- ✅ CONTENT SLIDES ----------
    for sp in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = BG_COLOR

        title_shape = slide.shapes.title
        title_shape.text = sp["title"]
        title_shape.text_frame.paragraphs[0].font.color.rgb = CORP_BLUE

        body = slide.placeholders[1]
        tf = body.text_frame

        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.word_wrap = True
        tf.clear()

        for b in sp["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(20)
            p.font.color.rgb = BULLET_BLACK
            p.level = 0

        body.top = title_shape.top + title_shape.height + Inches(0.3)

        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                body.width = prs.slide_width - Inches(4)
                body.left = Inches(0.5)

                slide.shapes.add_picture(
                    img,
                    prs.slide_width - Inches(3.5),
                    body.top,
                    width=Inches(3),
                )
        else:
            body.left = Inches(0.5)
            body.width = prs.slide_width - Inches(1)

    out_path = os.path.join(
        tempfile.gettempdir(),
        f"generated_{uuid.uuid4().hex[:8]}.pptx"
    )
    prs.save(out_path)
    return out_path


# ------------------------------------------------------------
# ✅ FINAL PIPELINE
# ------------------------------------------------------------
def generate_presentation_auto(
    prompt,
    requested_num_slides=5,
    tag_filters=None,
    template_style=None,
    image_required=False,
):
    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []

    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan_auto(prompt, reference_text, num_slides)

    if not plan:
        return None, {"error": True, "message": "Failed to generate slides."}

    ppt_title = f"{prompt.strip().title()} Overview"

    out_path = build_ppt(plan, ppt_title, image_required)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(plan) + 2,
        "ppt_file": fname,
        "error": False,
        "image_required": image_required,
        "template_style": template_style,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
