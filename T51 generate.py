# ============================================================
# generate_ppt.py – CORPORATE TEMPLATE (FINAL FIXED VERSION)
# ============================================================

import os
import tempfile
import uuid
import json
import re
from datetime import datetime

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.dml.color import RGBColor

from utils import get_env, logger, now_ts, ensure_dir, text_client
from search_utils import semantic_search
from azure_blob_utils import upload_ppt_to_blob, upload_json_to_blob

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
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
# LLM PLAN
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return slide plan in this format ONLY:\n\n"
        "Slide 1: Title\n"
        "- Bullet\n"
        "- Bullet\n"
        "- Bullet\n\n"
        f"You MUST create exactly {num_slides} slides.\n"
        "Do NOT return JSON.\n"
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

        title = lines[0].split(":", 1)[-1].strip()
        bullets = [l[1:].strip() for l in lines[1:] if l.startswith("-")]

        if len(bullets) < 3:
            bullets += [f"Additional point {i+1}" for i in range(3 - len(bullets))]

        slides.append({"title": title, "bullets": bullets[:8]})

    if len(slides) < num_slides:
        return None

    return slides[:num_slides]


# ------------------------------------------------------------
# TEXTBOX BULLET BUILDER ✅ FIXED
# ------------------------------------------------------------
def add_bullet_textbox(slide, bullets, left, top, width, height, font_size=20):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame

    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = True
    tf.clear()

    for b in bullets:
        p = tf.add_paragraph()
        p.text = b
        p.level = 0                # ✅ REAL BULLETS
        p.font.size = Pt(font_size)
        p.font.color.rgb = RGBColor(0, 0, 0)   # ✅ BLACK TEXT
        p.font.bold = False

    return tb


# ------------------------------------------------------------
# DATE UPDATER ✅ FIXED
# ------------------------------------------------------------
def force_update_date(slide):
    current = datetime.now().strftime("%B %Y")

    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False):
            text = shp.text.strip()
            if "date" in text.lower() or re.search(r"\b20\d{2}\b", text):
                shp.text = current


# ------------------------------------------------------------
# SLIDE DELETE
# ------------------------------------------------------------
def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slide_id = xml_slides[index].rId
    prs.part.drop_rel(slide_id)
    del xml_slides[index]


# ------------------------------------------------------------
# ✅✅✅ CORPORATE PPT BUILDER – FINAL FIXED VERSION ✅✅✅
# ------------------------------------------------------------
def build_corporate_ppt(plan):
    if not os.path.exists(CORPORATE_TEMPLATE_PATH):
        raise FileNotFoundError("corporate.pptx not found")

    prs = Presentation(CORPORATE_TEMPLATE_PATH)

    title_layout = prs.slide_layouts[0]
    agenda_layout = prs.slide_layouts[1]
    content_layout = prs.slide_layouts[2]
    thanks_layout = prs.slide_layouts[-1]

    for i in reversed(range(len(prs.slides))):
        delete_slide(prs, i)

    # ---- 1) TITLE SLIDE ✅ TITLE + ✅ DATE FIX ----
    title_slide = prs.slides.add_slide(title_layout)
    title_slide.shapes.title.text = plan[0]["title"]
    force_update_date(title_slide)

    # ---- 2) AGENDA SLIDE ✅ BULLETS FIX ----
    agenda_slide = prs.slides.add_slide(agenda_layout)
    agenda_slide.shapes.title.text = "Agenda"

    agenda_titles = [s["title"] for s in plan]

    add_bullet_textbox(
        agenda_slide,
        agenda_titles,
        left=Inches(1.2),
        top=Inches(2.2),
        width=Inches(7.5),
        height=Inches(4.8),
        font_size=28,
    )

    # ---- 3–7) CONTENT SLIDES ✅ BULLETS + WIDTH FIX ----
    for sp in plan:
        slide = prs.slides.add_slide(content_layout)
        slide.shapes.title.text = sp["title"]

        add_bullet_textbox(
            slide,
            sp["bullets"],
            left=Inches(1.2),
            top=Inches(2.2),
            width=Inches(7.8),   # ✅ FULL WIDTH FIX
            height=Inches(4.8),
            font_size=20,
        )

    # ---- LAST) THANK YOU SLIDE ✅ UNCHANGED ----
    prs.slides.add_slide(thanks_layout)

    out = os.path.join(temptempfile := tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out)
    return out


# ------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------
def generate_presentation(prompt, requested_num_slides=5, tag_filters=None):

    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []
    if not refs:
        return None, {"error": True, "message": "No matching content found."}

    reference_text = [(r.get("text") or "")[:500] for r in refs]

    detected = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected or 5

    plan = call_llm_plan(prompt, reference_text, num_slides)
    if not plan:
        return None, {"error": True, "message": "LLM content insufficient"}

    ppt_path = build_corporate_ppt(plan)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(ppt_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(plan) + 3,
        "ppt_file": fname,
        "template_style": "corporate",
        "error": False,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")
    return ppt_path, log
