# ============================================================
# generate_ppt_auto.py – AUTO | TITLE + AGENDA + CONTENT | AUTO SIZE
# ============================================================

import os
import tempfile
import uuid
import json
import re
import base64
from datetime import datetime

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


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def parse_user_intent(prompt: str):
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    return int(match.group(1)) if match else None


def fallback_plan(n):
    return [{
        "title": f"Slide {i+1}",
        "bullets": [
            f"Key takeaway {i+1}",
            f"Supporting detail {i+1}.1",
            f"Supporting detail {i+1}.2",
        ]
    } for i in range(n)]


# ------------------------------------------------------------
# ✅ TEXT-BASED LLM
# ------------------------------------------------------------
def call_llm_plan_auto(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return ONLY this format:\n"
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


# ------------------------------------------------------------
# TEXT → SLIDES
# ------------------------------------------------------------
def parse_text_plan(text, num_slides):
    slides = []
    blocks = re.split(r"\n(?=Slide \d+:)", text)

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue

        title = lines[0].split(":", 1)[-1].strip()
        bullets = [l.replace("-", "").strip() for l in lines[1:] if l.startswith("-")]
        bullets += [f"Additional point {i+1}" for i in range(3 - len(bullets))]

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
            prompt=prompt + " Professional minimal illustration",
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
# ✅✅✅ PPT BUILDER WITH CORRECT TITLE SLIDE ✅✅✅
# ------------------------------------------------------------
def build_ppt(slides, image_required):
    prs = Presentation()

    # ✅✅✅ TITLE SLIDE (FIXED TO MATCH YOUR 2ND IMAGE)
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])

    title_tf = title_slide.shapes.title.text_frame
    title_tf.clear()
    p = title_tf.paragraphs[0]
    p.text = slides[0]["title"]
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)

    subtitle = title_slide.placeholders[1]
    subtitle.text = datetime.now().strftime("%B %Y")
    subtitle.text_frame.paragraphs[0].font.size = Pt(20)
    subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(120, 120, 120)

    # ✅ NO IMAGE ON TITLE SLIDE

    # --------------------------------------------------------
    # ✅ AGENDA SLIDE (IMAGE ALLOWED)
    # --------------------------------------------------------
    agenda_slide = prs.slides.add_slide(prs.slide_layouts[1])
    agenda_slide.shapes.title.text = "Agenda"

    agenda_tf = agenda_slide.placeholders[1].text_frame
    agenda_tf.auto_size = MSO_AUTO_SIZE.NONE
    agenda_tf.word_wrap = True
    agenda_tf.clear()

    agenda_titles = [s["title"] for s in slides]

    for t in agenda_titles:
        p = agenda_tf.add_paragraph()
        p.text = t
        p.font.size = Pt(20)

    if image_required:
        img = generate_visual_image("Presentation Agenda")
        if img:
            agenda_slide.shapes.add_picture(
                img,
                prs.slide_width - Inches(3.5),
                agenda_slide.shapes.title.top + Inches(1),
                width=Inches(3),
            )

    # --------------------------------------------------------
    # ✅ CONTENT SLIDES
    # --------------------------------------------------------
    for sp in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = sp["title"]
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 102, 204)

        body = slide.placeholders[1]
        tf = body.text_frame
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.word_wrap = True
        tf.clear()

        for b in sp["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(20)

        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                body.width = prs.slide_width - Inches(4)
                slide.shapes.add_picture(
                    img,
                    prs.slide_width - Inches(3.5),
                    body.top,
                    width=Inches(3),
                )

    out_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
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
    detected_slides = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected_slides or 5

    plan = call_llm_plan_auto(prompt, reference_text, num_slides)

    slides = []
    for sp in plan:
        slides.append({
            "title": sp["title"],
            "bullets": sp["bullets"],
        })

    out_path = build_ppt(slides, image_required)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides) + 2,  # ✅ title + agenda
        "ppt_file": fname,
        "error": False,
        "image_required": image_required,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")
    return out_path, log
