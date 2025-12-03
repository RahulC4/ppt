# ============================================================
# generate_ppt.py – NO JSON | Images Checkbox | Corporate Template (GENUPDATE3 – BULLET + DATE FIX)
# ============================================================

import os
import tempfile
import uuid
import json
import re
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
)
from search_utils import semantic_search
from azure_blob_utils import upload_ppt_to_blob, upload_json_to_blob

ensure_dir("generated")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
CORPORATE_TEMPLATE_PATH = os.path.join(TEMPLATES_DIR, "corporate.pptx")


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
# LLM PLAN
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
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


def parse_text_plan(text, num_slides):
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
# ✅ PLACEHOLDER-SAFE BULLET WRITER (THE REAL FIX)
# ------------------------------------------------------------
def write_bullets_to_placeholder(slide, bullets, font_size):
    for shape in slide.shapes:
        if shape.is_placeholder and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()

            for b in bullets:
                p = tf.add_paragraph()
                p.text = b
                p.level = 0
                p.font.size = Pt(font_size)
                p.font.color.rgb = RGBColor(0, 0, 0)
            return


def set_title(slide, text):
    if slide.shapes.title:
        slide.shapes.title.text = text


# ✅ ✅ ✅ DATE FIX (RUNS BEFORE SHAPES ARE CLEARED)
def update_date(slide):
    current = datetime.now().strftime("%B %Y")
    for shp in slide.shapes:
        if shp.has_text_frame and re.search(r"\b20\d{2}\b", shp.text):
            shp.text = current


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slide_id = xml_slides[index].rId
    prs.part.drop_rel(slide_id)
    del xml_slides[index]


# ------------------------------------------------------------
# ✅✅✅ FIXED CORPORATE BUILDER (BULLETS + DATE WORKING)
# ------------------------------------------------------------
def build_corporate_ppt(plan):
    if not os.path.exists(CORPORATE_TEMPLATE_PATH):
        raise ValueError("Corporate template not found")

    prs = Presentation(CORPORATE_TEMPLATE_PATH)

    title_layout = prs.slide_layouts[0]
    agenda_layout = prs.slide_layouts[1]
    content_layout = prs.slide_layouts[2]
    thankyou_layout = prs.slide_layouts[-1]

    for i in reversed(range(len(prs.slides))):
        delete_slide(prs, i)

    # ---- 1) TITLE SLIDE ----
    title_slide = prs.slides.add_slide(title_layout)

    update_date(title_slide)
    set_title(title_slide, plan[0]["title"])

    # ---- 2) AGENDA SLIDE ----
    agenda_slide = prs.slides.add_slide(agenda_layout)
    set_title(agenda_slide, "Agenda")

    agenda_items = [s["title"] for s in plan]
    write_bullets_to_placeholder(agenda_slide, agenda_items, font_size=28)

    # ---- 3..N+2 CONTENT SLIDES ----
    for sp in plan:
        slide = prs.slides.add_slide(content_layout)
        set_title(slide, sp["title"])
        write_bullets_to_placeholder(slide, sp["bullets"], font_size=20)

    # ---- THANK YOU ----
    prs.slides.add_slide(thankyou_layout)

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
    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan(prompt, reference_text, num_slides)

    ppt_path = build_corporate_ppt(plan)
    total_slides = len(plan) + 3

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(ppt_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": total_slides,
        "ppt_file": fname,
        "image_required": False,
        "template_style": template_style,
        "error": False,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")
    return ppt_path, log
