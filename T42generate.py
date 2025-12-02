# ============================================================
# generate_ppt.py – NO JSON | Images Checkbox | Corporate Template (v2)
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
        slides.append(
            {
                "title": f"Slide {i+1}",
                "bullets": [
                    f"Main point {i+1}",
                    f"Supporting detail {i+1}.1",
                    f"Supporting detail {i+1}.2",
                ],
            }
        )
    return slides


# ------------------------------------------------------------
# LLM PLAN (TEXT MODE, NO JSON)
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "You are a professional presentation writer.\n\n"
        "Return slide plan in this format ONLY:\n\n"
        "Slide 1: Title of slide 1\n"
        "- Bullet point one\n"
        "- Bullet point two\n"
        "- Bullet point three\n\n"
        "Slide 2: Title of slide 2\n"
        "- Bullet point one\n"
        "- Bullet point two\n\n"
        f"You MUST create exactly {num_slides} slides.\n"
        "Do NOT return JSON.\n"
        "Do NOT add explanations.\n\n"
        f"Use this reference content when relevant:\n{' '.join(references_text)[:2500]}"
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
        raw = resp.choices[0].message.content.strip()
        return parse_text_plan(raw, num_slides)
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

        # "Slide X: Title"
        title = lines[0].split(":", 1)[-1].strip()
        bullets = [l[1:].strip() for l in lines[1:] if l.startswith("-")]

        if len(bullets) < 3:
            bullets += [f"Additional point {i+1}" for i in range(3 - len(bullets))]

        slides.append({"title": title, "bullets": bullets[:8]})

    if len(slides) < num_slides:
        logger.warning(
            f"Parsed only {len(slides)} slides but {num_slides} were requested."
        )
        return None

    return slides[:num_slides]


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    if not prompt:
        return None

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
# GENERIC SLIDE HELPERS
# ------------------------------------------------------------
def set_title(slide, text: str):
    """Set slide title if there is a title shape."""
    try:
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            slide.shapes.title.text = text
    except Exception:
        pass


def clear_and_set_bullets_in_shape(shape, bullets):
    """Helper to fill an existing text frame with bullets."""
    tf = shape.text_frame
    tf.clear()
    for b in bullets:
        p = tf.add_paragraph()
        p.text = b
        p.level = 0
        p.font.size = Pt(20)


def add_bullet_textbox(slide, bullets, left, top, width, height):
    """
    Create a new textbox at a fixed position and fill bullets.
    Used for corporate Agenda + content slides so we don't depend
    on weird template placeholders.
    """
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.clear()
    for b in bullets:
        p = tf.add_paragraph()
        p.text = b
        p.level = 0
        p.font.size = Pt(20)
    return tb


def add_image(slide, img_path, body_shape=None):
    """Place image to the right of the given body_shape, or top-right."""
    try:
        if body_shape is not None:
            body_shape.width = int(body_shape.width * 0.60)
            left = body_shape.left + body_shape.width + Inches(0.25)
            top = body_shape.top
        else:
            left = Inches(6.5)
            top = Inches(1.5)

        slide.shapes.add_picture(img_path, left, top, width=Inches(2.8))
    except Exception:
        logger.exception("Image placement failed")


def update_date(slide):
    """Replace any year-like text with current 'Month YYYY'."""
    current = datetime.now().strftime("%B %Y")
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False) and shp.text:
            if re.search(r"\b20\d{2}\b", shp.text):
                shp.text = current


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst  # type: ignore[attr-defined]
    slide_id = xml_slides[index].rId
    prs.part.drop_rel(slide_id)
    del xml_slides[index]


# ------------------------------------------------------------
# DEFAULT PPT BUILDER (no corporate theme)
# ------------------------------------------------------------
def build_default_ppt(plan, image_required):
    prs = Presentation()

    for sp in plan:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title & Content
        set_title(slide, sp["title"])

        # use the default body placeholder if present
        body = None
        for shp in slide.shapes:
            if getattr(shp, "has_text_frame", False) and shp is not slide.shapes.title:
                body = shp
                break

        if body is None:
            # fallback: our own textbox
            body = add_bullet_textbox(
                slide, sp["bullets"], left=Inches(1), top=Inches(2),
                width=Inches(6), height=Inches(4)
            )
        else:
            clear_and_set_bullets_in_shape(body, sp["bullets"])

        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                add_image(slide, img, body)

    out = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out)
    return out


# ------------------------------------------------------------
# CORPORATE PPT BUILDER
# ------------------------------------------------------------
def build_corporate_ppt(plan, image_required):
    """
    Build deck using templates/corporate.pptx.

    Structure:
      1. Title slide (from template, title + date, speaker removed)
      2. Agenda slide (template background, our own bullets textbox of slide titles)
      3..N+2. Content slides (template background, our own bullets textbox; optional image)
      Last. Thank-you slide (template as-is, no image)
    """
    if not os.path.exists(CORPORATE_TEMPLATE_PATH):
        logger.warning("Corporate template not found; falling back to default.")
        return build_default_ppt(plan, image_required)

    prs = Presentation(CORPORATE_TEMPLATE_PATH)

    if len(prs.slides) < 3:
        logger.warning("Corporate template has too few slides; falling back to default.")
        return build_default_ppt(plan, image_required)

    title_layout = prs.slides[0].slide_layout
    agenda_layout = prs.slides[1].slide_layout
    content_layout = prs.slides[2].slide_layout
    thanks_layout = prs.slides[-1].slide_layout

    # delete all original slides; keep layouts only
    for i in reversed(range(len(prs.slides))):
        delete_slide(prs, i)

    # ---- 1) TITLE SLIDE ----
    deck_title = plan[0]["title"] if plan and plan[0].get("title") else "Executive Summary"
    title_slide = prs.slides.add_slide(title_layout)
    set_title(title_slide, deck_title)
    update_date(title_slide)

    # try to clear speaker/subtitle placeholders so only our title remains
    for shp in title_slide.shapes:
        if shp is title_slide.shapes.title:
            continue
        if getattr(shp, "has_text_frame", False):
            shp.text = ""

    # ---- 2) AGENDA SLIDE ----
    agenda_slide = prs.slides.add_slide(agenda_layout)
    set_title(agenda_slide, "Agenda")

    agenda_items = [s["title"] for s in plan] or [
        f"Section {i+1}" for i in range(len(plan))
    ]

    # Draw our own bullet textbox so we don't depend on the template body
    agenda_body = add_bullet_textbox(
        agenda_slide,
        agenda_items,
        left=Inches(1.0),
        top=Inches(2.0),
        width=Inches(8.0),
        height=Inches(4.0),
    )

    # ---- 3..N+2) CONTENT SLIDES ----
    for sp in plan:
        slide = prs.slides.add_slide(content_layout)
        set_title(slide, sp["title"])

        # our own textbox for bullets, always on the left
        body = add_bullet_textbox(
            slide,
            sp["bullets"],
            left=Inches(1.0),
            top=Inches(2.0),
            width=Inches(6.0),   # leave space on the right for image
            height=Inches(4.0),
        )

        if image_required:
            img = generate_visual_image(sp["title"])
            if img:
                add_image(slide, img, body)

    # ---- LAST) THANK-YOU SLIDE ----
    prs.slides.add_slide(thanks_layout)

    out = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out)
    return out


# ------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------
def generate_presentation(
    prompt,
    requested_num_slides=5,
    tag_filters=None,
    template_style=None,
    image_required=False,
):
    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []
    if not refs:
        return None, {
            "error": True,
            "message": "I couldn’t find relevant content in your sample PPTs. "
                       "Try a prompt closer to your uploaded decks.",
        }

    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan(prompt, references_text=reference_text, num_slides=num_slides)
    if not plan:
        return None, {
            "error": True,
            "message": "Not enough relevant content to generate this many slides. "
                       "Try fewer slides or rephrase your prompt.",
        }

    template_key = (template_style or "").lower()
    if template_key == "corporate":
        ppt_path = build_corporate_ppt(plan, image_required)
        total_slides = len(plan) + 3  # title + agenda + thank-you
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

    upload_json_to_blob(json.dumps(log, indent=2).encode("utf-8"), f"logs/{fname}.json")
    return ppt_path, log
