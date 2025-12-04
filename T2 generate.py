# ============================================================
# generate_ppt_auto.py – CLEAN AUTO MODE (TITLE + AGENDA + IMAGES)
# ============================================================

import os
import uuid
import json
import re
import tempfile
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from utils import (
    get_env,
    logger,
    now_ts,
    ensure_dir,
    text_client,
    image_client,
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
    return [
        {
            "title": f"Slide {i+1}",
            "bullets": [
                f"Main point {i+1}",
                f"Supporting detail {i+1}.1",
                f"Supporting detail {i+1}.2",
            ],
        }
        for i in range(n)
    ]


# ------------------------------------------------------------
# ✅ AUTO LLM PLAN
# ------------------------------------------------------------

def call_llm_auto(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return ONLY this format:\n"
        "Slide 1: Title\n"
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
        return parse_text_plan_auto(raw, num_slides)

    except Exception as e:
        logger.warning(f"AUTO LLM failed → fallback used: {e}")
        return fallback_plan(num_slides)


def parse_text_plan_auto(text, num_slides):
    slides = []
    blocks = re.split(r"\n(?=Slide\s+\d+\s*:)", text)

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue

        title = lines[0].split(":", 1)[-1].strip()
        bullets = [l[1:].strip() for l in lines[1:] if l.startswith("-")]
        bullets = bullets[:8]

        slides.append({"title": title, "bullets": bullets})

    return slides[:num_slides]


# ------------------------------------------------------------
# ✅ SLIDE WRITERS
# ------------------------------------------------------------

def write_title(slide, title_text):
    for shp in slide.shapes:
        if shp.has_text_frame:
            tf = shp.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(36)
            return


def write_bullets(slide, bullets):
    for shp in slide.shapes:
        if shp.has_text_frame:
            tf = shp.text_frame
            tf.clear()

            for b in bullets:
                p = tf.add_paragraph()
                p.text = b
                p.level = 0
                p.font.size = Pt(20)
                return


def add_image(slide, prompt):
    try:
        img = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=prompt,
            size="1024x1024",
        )

        img_b64 = img.data[0].b64_json
        img_bytes = base64.b64decode(img_b64)

        tmp_img = os.path.join(tempfile.gettempdir(), f"{uuid.uuid4().hex}.png")
        with open(tmp_img, "wb") as f:
            f.write(img_bytes)

        slide.shapes.add_picture(
            tmp_img,
            left=Inches(6.8),
            top=Inches(1.5),
            width=Inches(3),
        )

    except Exception as e:
        logger.warning(f"AUTO image failed: {e}")


# ------------------------------------------------------------
# ✅✅✅ AUTO PPT BUILDER (FIXED)
# ------------------------------------------------------------

def build_auto_ppt(plan):
    prs = Presentation()

    TITLE_LAYOUT = 0
    CONTENT_LAYOUT = 1

    # ---- 1) TITLE SLIDE (NO IMAGE) ----
    title_slide = prs.slides.add_slide(prs.slide_layouts[TITLE_LAYOUT])
    write_title(title_slide, plan[0]["title"])

    # ---- 2) AGENDA SLIDE (IMAGE ALLOWED) ----
    agenda_slide = prs.slides.add_slide(prs.slide_layouts[CONTENT_LAYOUT])
    write_title(agenda_slide, "Agenda")

    agenda_items = [p["title"] for p in plan]
    write_bullets(agenda_slide, agenda_items)
    add_image(agenda_slide, "business agenda presentation")

    # ---- 3..N CONTENT SLIDES (IMAGES ALLOWED) ----
    for sp in plan:
        slide = prs.slides.add_slide(prs.slide_layouts[CONTENT_LAYOUT])
        write_title(slide, sp["title"])
        write_bullets(slide, sp["bullets"])
        add_image(slide, sp["title"])

    out = os.path.join(
        tempfile.gettempdir(),
        f"generated_auto_{uuid.uuid4().hex[:8]}.pptx",
    )

    prs.save(out)
    return out


# ------------------------------------------------------------
# ✅✅✅ MAIN AUTO PIPELINE
# ------------------------------------------------------------

def generate_presentation_auto(
    prompt,
    requested_num_slides=5,
    tag_filters=None,
    image_required=True,
):
    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []
    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_auto(prompt, reference_text, num_slides)

    ppt_path = build_auto_ppt(plan)
    fname = f"generated_auto_{uuid.uuid4().hex[:8]}.pptx"

    upload_ppt_to_blob(ppt_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(plan) + 2,  # title + agenda
        "ppt_file": fname,
        "image_required": True,
        "template_style": "auto",
        "error": False,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return ppt_path, log
