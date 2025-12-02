# ============================================================
# generate_ppt.py – CORPORATE TEMPLATE + AUTO SIZE + SAFE CLONE
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

CORPORATE_TEMPLATE_PATH = "templates/corporate.pptx"


# ------------------------------------------------------------
# ✅ HELPERS
# ------------------------------------------------------------
def parse_user_intent(prompt: str):
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    return int(match.group(1)) if match else None


def fallback_plan(n):
    return [{
        "title": f"Slide {i+1}",
        "bullets": [
            f"Key takeaway for slide {i+1}",
            f"Supporting detail {i+1}.1",
            f"Supporting detail {i+1}.2",
        ]
    } for i in range(n)]


# ------------------------------------------------------------
# ✅ LLM PLAN
# ------------------------------------------------------------
def call_llm_plan(prompt, references_text=None, num_slides=5):
    references_text = references_text or []

    sys_prompt = (
        "Return EXACT format:\n"
        "Slide 1: Title\n- Bullet\n- Bullet\n- Bullet\n\n"
        f"Create exactly {num_slides} slides.\n"
        "NO JSON.\n\n"
        f"{' '.join(references_text)[:2500]}"
    )

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": f"Create presentation for: {prompt}"},
            ],
            max_completion_tokens=1400,
        )

        raw_text = resp.choices[0].message.content.strip()
        return parse_text_plan(raw_text, num_slides)

    except Exception as e:
        logger.warning(f"LLM failed → fallback used: {e}")
        return fallback_plan(num_slides)


# ------------------------------------------------------------
# ✅ TEXT PARSER
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

        if len(bullets) < 3:
            bullets += [f"Additional point {i+1}" for i in range(3 - len(bullets))]

        slides.append({"title": title, "bullets": bullets[:6]})

    return slides[:num_slides] if len(slides) >= num_slides else None


# ------------------------------------------------------------
# ✅ IMAGE GENERATION
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
        return None


# ------------------------------------------------------------
# ✅ SAFE BULLET SLIDE BUILDER
# ------------------------------------------------------------
def add_ai_slide(prs, sp, image_required=True):
    slide = prs.slides.add_slide(prs.slide_layouts[1])

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
        p.level = 0

    body.left = Inches(0.5)
    body.width = prs.slide_width - Inches(1)

    # ✅ IMAGE
    if image_required and sp.get("image_path"):
        try:
            body.width = prs.slide_width - Inches(4)
            slide.shapes.add_picture(
                sp["image_path"],
                prs.slide_width - Inches(3.5),
                body.top,
                width=Inches(3),
            )
        except:
            pass

    # ✅ FOOTER (AI SLIDES ONLY)
    footer = slide.shapes.add_textbox(
        prs.slide_width - Inches(2),
        prs.slide_height - Inches(0.6),
        Inches(1.8),
        Inches(0.4),
    )
    tf = footer.text_frame
    tf.text = "Cognizant"
    tf.paragraphs[0].font.size = Pt(10)
    tf.paragraphs[0].font.color.rgb = RGBColor(0, 112, 192)


# ------------------------------------------------------------
# ✅✅✅ CORPORATE + AI COMBINED GENERATOR
# ------------------------------------------------------------
def build_ppt(slides, prompt, template_style, image_required):
    if template_style == "corporate":
        prs = Presentation(CORPORATE_TEMPLATE_PATH)
    else:
        prs = Presentation()

    # ✅ FIRST SLIDE TITLE + DATE
    first_slide = prs.slides[0]
    for shape in first_slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.lower()
            if "overview" in text:
                shape.text_frame.text = prompt
            if "202" in text:
                shape.text_frame.text = datetime.now().strftime("%B %Y")

    # ✅ AGENDA SLIDE
    agenda_slide = prs.slides[1]
    agenda_body = None
    for shape in agenda_slide.shapes:
        if shape.has_text_frame:
            agenda_body = shape.text_frame
            break

    if agenda_body:
        agenda_body.clear()
        for s in slides:
            p = agenda_body.add_paragraph()
            p.text = s["title"]
            p.level = 0

    # ✅ SKIP IMAGE FOR FIRST & LAST
    for i, sp in enumerate(slides):
        if i == 0 or i == len(slides) - 1:
            sp["image_path"] = None

        add_ai_slide(prs, sp, image_required)

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

    plan = call_llm_plan(prompt, reference_text, num_slides)

    slides = []
    for sp in plan:
        img_path = generate_visual_image(sp["title"]) if image_required else None
        slides.append({
            "title": sp["title"],
            "bullets": sp["bullets"],
            "image_path": img_path,
        })

    out_path = build_ppt(slides, prompt, template_style, image_required)

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
