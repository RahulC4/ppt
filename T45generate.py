# ============================================================
# generate_ppt.py – CORPORATE TEMPLATE + AUTO SIZE + SAFE ALIGNMENT
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
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
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

CORPORATE_TEMPLATE_PATH = "templates/corporate.pptx"


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
# ✅ TEXT LLM (NO JSON)
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = f"""
You are a professional presentation creator.

Return EXACT format only:

Slide 1: Title
- Bullet
- Bullet
- Bullet

Create exactly {num_slides} slides.
No JSON. No explanations.

Reference:
{' '.join(references_text)[:2500]}
"""

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": f"Create a presentation on: {prompt}"}
            ],
            max_completion_tokens=1400,
            temperature=1,
        )

        return parse_text_plan(resp.choices[0].message.content.strip(), num_slides)

    except Exception as e:
        logger.warning(f"LLM failed → fallback used: {e}")
        return fallback_plan(num_slides)


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

        while len(bullets) < 3:
            bullets.append("Additional point")

        slides.append({"title": title, "bullets": bullets[:6]})

    return slides[:num_slides] if len(slides) >= num_slides else None


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
# ✅ APPLY CORPORATE TEMPLATE
# ------------------------------------------------------------
def apply_corporate_template(prs_main, ai_slides, image_required):
    corp = Presentation(CORPORATE_TEMPLATE_PATH)

    # --- 1️⃣ COVER SLIDE ---
    cover = corp.slides[0]
    cover_title = cover.shapes.title
    cover_title.text = ai_slides[0]["title"]

    for shape in cover.shapes:
        if shape.has_text_frame:
            if any(month in shape.text for month in ["202", "Jan", "Feb", "Mar"]):
                shape.text = datetime.now().strftime("%B %Y")

    prs_main.slides._sldIdLst.append(corp.slides._sldIdLst[0])

    # --- 2️⃣ AGENDA SLIDE ---
    agenda = corp.slides[1]
    agenda_titles = [s["title"] for s in ai_slides]
    for shape in agenda.shapes:
        if shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            for t in agenda_titles:
                p = tf.add_paragraph()
                p.text = t
                p.level = 0

    prs_main.slides._sldIdLst.append(corp.slides._sldIdLst[1])

    # --- 3️⃣ AI SLIDES ---
    for i, sp in enumerate(ai_slides):
        slide = prs_main.slides.add_slide(prs_main.slide_layouts[1])
        slide.shapes.title.text = sp["title"]

        body = slide.placeholders[1]
        tf = body.text_frame
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.word_wrap = True
        tf.clear()

        for b in sp["bullets"]:
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(20)

        # ✅ BLUE FOOTER
        footer = slide.shapes.add_textbox(
            prs_main.slide_width - Inches(2),
            prs_main.slide_height - Inches(0.6),
            Inches(1.5),
            Inches(0.3),
        )
        footer_tf = footer.text_frame
        footer_tf.text = "Cognizant"
        footer_tf.paragraphs[0].font.size = Pt(10)
        footer_tf.paragraphs[0].font.bold = True
        footer_tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

        # ✅ IMAGE (NOT first & last)
        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                body.width = prs_main.slide_width - Inches(4)
                body.left = Inches(0.5)

                slide.shapes.add_picture(
                    img,
                    prs_main.slide_width - Inches(3.5),
                    body.top,
                    width=Inches(3),
                )

    # --- 4️⃣ THANK YOU ---
    prs_main.slides._sldIdLst.append(corp.slides._sldIdLst[2])


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

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan(prompt, reference_text, num_slides)
    if not plan:
        return None, {"error": True, "message": "Slide generation failed"}

    prs = Presentation()

    if template_style == "corporate":
        apply_corporate_template(prs, plan, image_required)

    out_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out_path)

    fname = os.path.basename(out_path)
    upload_ppt_to_blob(out_path, fname)

    upload_json_to_blob(
        json.dumps({
            "timestamp": now_ts(),
            "prompt": prompt,
            "slides_generated": len(plan),
            "ppt_file": fname,
            "template_style": template_style,
            "image_required": image_required,
            "error": False
        }, indent=2).encode(),
        f"logs/{fname}.json"
    )

    return out_path, {"ppt_file": fname, "error": False}
