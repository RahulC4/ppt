# ============================================================
# generate_ppt.py – CORPORATE TEMPLATE + AUTO SIZE FIX
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
from PIL import Image

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
    if match:
        return int(match.group(1))
    return None


def fallback_plan(n):
    slides = []
    for i in range(n):
        slides.append({
            "title": f"Slide {i+1}",
            "bullets": [
                f"Key takeaway for slide {i+1}",
                f"Supporting detail {i+1}.1",
                f"Supporting detail {i+1}.2",
            ]
        })
    return slides


# ------------------------------------------------------------
# ✅ TEXT-BASED LLM (NO JSON)
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return this EXACT format ONLY:\n\n"
        "Slide 1: Title\n"
        "- Bullet\n"
        "- Bullet\n"
        "- Bullet\n\n"
        f"You MUST create exactly {num_slides} slides.\n"
        "Do NOT return JSON.\n"
    )

    user_prompt = f"Create a professional presentation for: {prompt}"

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt},
            ],
            max_completion_tokens=1400,
            temperature=1,
        )

        raw_text = resp.choices[0].message.content.strip()
        return parse_text_plan(raw_text, num_slides)

    except Exception as e:
        logger.exception(e)
        return fallback_plan(num_slides)


# ------------------------------------------------------------
# ✅ TEXT → STRUCTURED SLIDES
# ------------------------------------------------------------
def parse_text_plan(text, num_slides):
    slides = []
    blocks = re.split(r"\n(?=Slide \d+:)", text)

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue

        title = lines[0].split(":", 1)[-1].strip()

        bullets = []
        for l in lines[1:]:
            if l.startswith("-"):
                bullets.append(l.replace("-", "").strip())

        while len(bullets) < 3:
            bullets.append("Additional point")

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
            prompt=prompt + " Professional minimal illustration. No text.",
            size="1024x1024",
        )

        b64 = resp.data[0].b64_json
        img_bytes = base64.b64decode(b64)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.write(img_bytes)
        tmp.close()
        return tmp.name

    except Exception as e:
        logger.exception(e)
        return None


# ------------------------------------------------------------
# ✅ CORPORATE PPT BUILDER
# ------------------------------------------------------------
def build_ppt(slides, presentation_title):
    template_path = os.path.join("templates", "corporate.pptx")

    if not os.path.exists(template_path):
        raise FileNotFoundError("corporate.pptx NOT FOUND in templates folder")

    prs = Presentation(template_path)

    # ✅ SLIDE 1 → TITLE REPLACEMENT
    slide1 = prs.slides[0]
    slide1.shapes.title.text = presentation_title

    # ✅ DATE REPLACEMENT
    for shape in slide1.shapes:
        if shape.has_text_frame:
            shape.text = datetime.now().strftime("%B %Y")

    # ✅ SLIDE 2 → AGENDA
    agenda_slide = prs.slides[1]
    body = agenda_slide.placeholders[1]
    tf = body.text_frame
    tf.clear()
    tf.auto_size = MSO_AUTO_SIZE.NONE

    for s in slides:
        p = tf.add_paragraph()
        p.text = s["title"]
        p.font.size = Pt(18)

    # ✅ GENERATE 5 CONTENT SLIDES
    for sp in slides:
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)

        slide.shapes.title.text = sp["title"]

        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()
        tf.auto_size = MSO_AUTO_SIZE.NONE

        for b in sp["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(20)

        # ✅ FOOTER (RIGHT SIDE BLUE)
        footer = slide.shapes.add_textbox(
            prs.slide_width - Inches(3),
            prs.slide_height - Inches(0.6),
            Inches(2.5),
            Inches(0.4)
        )
        tf = footer.text_frame
        tf.text = "Cognizant"
        tf.paragraphs[0].font.size = Pt(12)

        # ✅ IMAGE SUPPORT
        if sp.get("image_path"):
            body.width = prs.slide_width - Inches(4)
            body.left = Inches(0.5)

            slide.shapes.add_picture(
                sp["image_path"],
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
def generate_presentation(
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

    plan = call_llm_plan(
        prompt=prompt,
        references_text=reference_text,
        num_slides=num_slides,
    )

    slides = []
    for sp in plan:
        img_path = generate_visual_image(sp["title"]) if image_required else None
        slides.append({
            "title": sp["title"],
            "bullets": sp["bullets"],
            "image_path": img_path,
        })

    out_path = build_ppt(slides, prompt)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "ppt_file": fname,
        "error": False,
        "image_required": image_required,
        "template_style": template_style,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
