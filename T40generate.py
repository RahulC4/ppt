# ============================================================
# generate_ppt.py – FINAL STABLE VERSION (Corporate + Default)
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


# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
ensure_dir("generated")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
CORPORATE_TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "corporate.pptx")


# ------------------------------------------------------------
# PROMPT HELPERS
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
                f"Main point {i+1}",
                f"Supporting detail {i+1}.1",
                f"Supporting detail {i+1}.2",
            ],
        })
    return slides


# ------------------------------------------------------------
# LLM PLAN (TEXT MODE)
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return slide plan in this format only:\n\n"
        "Slide 1: Title\n"
        "- Bullet\n"
        "- Bullet\n\n"
        f"You MUST create exactly {num_slides} slides.\n"
        "Do NOT return JSON.\n"
        "Do NOT add explanations.\n\n"
        f"Reference content:\n{' '.join(references_text)[:2500]}"
    )

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": prompt},
            ],
            max_completion_tokens=1400,
            temperature=0.7,
        )

        return parse_text_plan(resp.choices[0].message.content.strip(), num_slides)

    except Exception as e:
        logger.warning(f"LLM failed → fallback used: {e}")
        return fallback_plan(num_slides)


def parse_text_plan(text, num_slides):
    slides = []
    blocks = re.split(r"\n(?=Slide\s+\d+\s*:)", text)

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue

        title = lines[0].split(":", 1)[-1].strip()
        bullets = [l[1:].strip() for l in lines[1:] if l.startswith("-")]

        if len(bullets) < 3:
            bullets += [f"Additional point {i+1}" for i in range(3 - len(bullets))]

        slides.append({"title": title, "bullets": bullets[:8]})

    return slides[:num_slides] if len(slides) >= num_slides else None


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=prompt + " professional minimal illustration",
            size="1024x1024",
        )

        b64 = getattr(resp.data[0], "b64_json", None)
        img_bytes = base64.b64decode(b64)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.write(img_bytes)
        tmp.close()
        return tmp.name

    except Exception:
        logger.exception("Image generation failed")
        return None


# ------------------------------------------------------------
# SAFE SLIDE WRITERS (TEMPLATE COMPATIBLE)
# ------------------------------------------------------------
def set_title(slide, text):
    if slide.shapes.title:
        slide.shapes.title.text = text


def clear_and_set_bullets(slide, bullets):
    body = None

    for shape in slide.shapes:
        if shape.has_text_frame and shape != slide.shapes.title:
            body = shape
            break

    if not body:
        logger.warning("No bullet textbox found")
        return None

    tf = body.text_frame
    tf.clear()

    for b in bullets:
        p = tf.add_paragraph()
        p.text = b
        p.level = 0
        p.font.size = Pt(20)

    return body


def add_image(slide, img_path, body_shape=None):
    try:
        if body_shape:
            body_shape.width = int(body_shape.width * 0.60)
            left = body_shape.left + body_shape.width + Inches(0.2)
            top = body_shape.top
        else:
            left = Inches(7)
            top = Inches(1.5)

        slide.shapes.add_picture(img_path, left, top, width=Inches(2.8))
    except:
        logger.exception("Image placement failed")


def update_date(slide):
    current = datetime.now().strftime("%B %Y")
    for shape in slide.shapes:
        if shape.has_text_frame and re.search(r"\b20\d{2}\b", shape.text):
            shape.text = current


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slide_id = xml_slides[index].rId
    prs.part.drop_rel(slide_id)
    del xml_slides[index]


# ------------------------------------------------------------
# DEFAULT PPT BUILDER
# ------------------------------------------------------------
def build_default_ppt(plan, image_required):
    prs = Presentation()

    for sp in plan:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        set_title(slide, sp["title"])
        body = clear_and_set_bullets(slide, sp["bullets"])

        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                add_image(slide, img, body)

    out = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out)
    return out


# ------------------------------------------------------------
# CORPORATE PPT BUILDER (FINAL)
# ------------------------------------------------------------
def build_corporate_ppt(plan, image_required):

    prs = Presentation(CORPORATE_TEMPLATE_PATH)

    title_layout = prs.slides[0].slide_layout
    agenda_layout = prs.slides[1].slide_layout
    content_layout = prs.slides[2].slide_layout
    thanks_layout = prs.slides[-1].slide_layout

    for i in reversed(range(len(prs.slides))):
        delete_slide(prs, i)

    # ---- Title ----
    slide1 = prs.slides.add_slide(title_layout)
    set_title(slide1, plan[0]["title"])
    update_date(slide1)

    # ---- Agenda ----
    slide2 = prs.slides.add_slide(agenda_layout)
    set_title(slide2, "Agenda")
    agenda_items = [s["title"] for s in plan]
    clear_and_set_bullets(slide2, agenda_items)

    # ---- Main Content Slides (3 → N) ----
    for sp in plan:
        slide = prs.slides.add_slide(content_layout)

        set_title(slide, sp["title"])
        body = clear_and_set_bullets(slide, sp["bullets"])

        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                add_image(slide, img, body)

    # ---- Thank You ----
    prs.slides.add_slide(thanks_layout)

    out = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out)
    return out


# ------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------
def generate_presentation(prompt, requested_num_slides=5, tag_filters=None,
                          template_style=None, image_required=False):

    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []
    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan(prompt, references_text=reference_text, num_slides=num_slides)

    if not plan:
        return None, {"error": True, "message": "Not enough content."}

    if str(template_style).lower() == "corporate":
        ppt_path = build_corporate_ppt(plan, image_required)
        total_slides = len(plan) + 3
    else:
        ppt_path = build_default_ppt(plan, image_required)
        total_slides = len(plan)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(ppt_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": total_slides,
        "ppt_file": fname,
        "image_required": image_required,
        "template_style": template_style,
        "error": False,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")
    return ppt_path, log
