# ============================================================
# generate_ppt.py – TEMPLATE VERSION WITH STRICT SLIDE COUNT
# ============================================================

import os
import tempfile
import uuid
import json
import re
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

from utils import (
    get_env, safe_json_load, logger, now_ts,
    ensure_dir, text_client, image_client
)
from search_utils import semantic_search
from azure_blob_utils import upload_ppt_to_blob, upload_json_to_blob

# ------------------------------------------------------------
# ENV + SETUP
# ------------------------------------------------------------

ensure_dir("design_jsons")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)


# ------------------------------------------------------------
# USER INTENT PARSER
# ------------------------------------------------------------

def parse_user_intent(prompt: str):
    """Detect 'N slides' from user prompt if present."""
    num_slides = None
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    if match:
        num_slides = int(match.group(1))
    return num_slides


# ------------------------------------------------------------
# TEMPLATE STYLE ENGINE (DARK / CORPORATE / MODERN)
# ------------------------------------------------------------

def apply_template_style(slide, template_style: str):

    bg = slide.background.fill
    bg.solid()

    style = (template_style or "corporate").lower()

    if style == "dark":
        bg.fore_color.rgb = RGBColor(20, 20, 20)
        title_color = RGBColor(255, 255, 255)
        body_color = RGBColor(220, 220, 220)

    elif style == "modern":
        bg.fore_color.rgb = RGBColor(15, 76, 129)
        title_color = RGBColor(255, 255, 255)
        body_color = RGBColor(230, 230, 230)

    elif style == "minimal":
        bg.fore_color.rgb = RGBColor(255, 255, 255)
        title_color = RGBColor(30, 30, 30)
        body_color = RGBColor(60, 60, 60)

    elif style == "plant":
        bg.fore_color.rgb = RGBColor(236, 248, 237)
        title_color = RGBColor(27, 94, 32)
        body_color = RGBColor(51, 51, 51)

    else:  # corporate default
        bg.fore_color.rgb = RGBColor(245, 246, 250)
        title_color = RGBColor(32, 55, 100)
        body_color = RGBColor(60, 60, 60)

    # Title styling
    if slide.shapes.title and slide.shapes.title.text_frame:
        for p in slide.shapes.title.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = title_color
                r.font.size = Pt(32)
                r.font.bold = True

    # Body styling
    try:
        body = slide.placeholders[1]
        if body.text_frame:
            for p in body.text_frame.paragraphs:
                for r in p.runs:
                    r.font.color.rgb = body_color
                    r.font.size = Pt(20)
    except Exception:
        pass


# ------------------------------------------------------------
# FALLBACK PLAN (NO / BAD LLM JSON)
# ------------------------------------------------------------

def build_fallback_plan(references_text, num_slides: int):
    """
    Build a simple multi-slide plan directly from reference text
    when LLM JSON is invalid or fails.
    Ensures EXACTLY num_slides slides.
    """
    if num_slides <= 0:
        num_slides = 1

    combined = " ".join(references_text) if references_text else ""
    combined = combined.strip()

    if not combined:
        # totally generic fallback
        plan = []
        for i in range(num_slides):
            title = "Overview" if i == 0 else f"Details {i}"
            plan.append({
                "title": title,
                "bullets": ["Content not available from knowledge base."],
                "visual_required": False,
                "visual_prompt": ""
            })
        return plan

    # split into sentence-like chunks
    parts = re.split(r'(?<=[.!?])\s+', combined)
    parts = [p.strip() for p in parts if p.strip()]

    if not parts:
        parts = [combined]

    # chunk sentences evenly across slides
    chunks = [[] for _ in range(num_slides)]
    for idx, sentence in enumerate(parts):
        chunks[idx % num_slides].append(sentence)

    plan = []
    for i in range(num_slides):
        bullets = chunks[i] or ["(continued)"]
        # limit bullets per slide so they don't explode
        bullets = bullets[:6]
        plan.append({
            "title": f"Slide {i+1}",
            "bullets": bullets,
            "visual_required": False,
            "visual_prompt": ""
        })

    return plan


# ------------------------------------------------------------
# SAFE LLM PLAN GENERATOR (STRICT SLIDE COUNT)
# ------------------------------------------------------------

def call_llm_plan(prompt, style, references_text, num_slides=None):
    """
    Ask LLM to create structured plan from reference text.
    Always returns a list of length == num_slides (via fix-up).
    """

    if not num_slides or num_slides <= 0:
        num_slides = 5

    sys_prompt = (
        "You are a presentation assistant.\n"
        "You MUST respond with STRICT JSON, no extra text, no markdown.\n"
        "JSON format:\n"
        "[\n"
        "  {\n"
        "    \"title\": \"...\",\n"
        "    \"bullets\": [\"...\",\"...\"],\n"
        "    \"visual_required\": true or false,\n"
        "    \"visual_prompt\": \"short prompt for an illustration\"\n"
        "  },\n"
        "  ... (exactly N slides)\n"
        "]\n"
        "N = number_of_slides I ask for.\n"
        "Use ONLY the reference content I provide (paraphrasing is ok). "
        "Do NOT invent unrelated topics.\n"
    )

    user_prompt = (
        f"Create a {style} presentation plan based ONLY on this content:\n\n"
        + "\n\n".join(references_text[:10]) +
        f"\n\nMake EXACTLY {num_slides} slides in the JSON list."
    )

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt}
            ],
            max_completion_tokens=1200,
            temperature=1
        )

        raw = resp.choices[0].message.content
        plan = safe_json_load(raw)

        if not isinstance(plan, list) or not plan:
            raise ValueError("Plan is not a non-empty list")

        # ensure items have needed keys
        cleaned = []
        for item in plan:
            if not isinstance(item, dict):
                continue
            cleaned.append({
                "title": item.get("title", "Slide"),
                "bullets": item.get("bullets", []),
                "visual_required": bool(item.get("visual_required", False)),
                "visual_prompt": item.get("visual_prompt", "")
            })

        if not cleaned:
            raise ValueError("Cleaned plan is empty")

        plan = cleaned

        # ✅ ENFORCE EXACT SLIDE COUNT
        if len(plan) > num_slides:
            plan = plan[:num_slides]
        elif len(plan) < num_slides:
            while len(plan) < num_slides:
                last = dict(plan[-1])
                last["title"] = last.get("title", "Slide") + " (cont.)"
                plan.append(last)

        return plan

    except Exception:
        logger.exception("LLM plan failed – using fallback multi-slide plan")
        return build_fallback_plan(references_text, num_slides)


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------

def generate_visual_image(prompt: str):

    if not prompt:
        return None

    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=prompt + " professional business illustration, no text labels",
            size="1024x1024"
        )

        img_b64 = resp.data[0].b64_json
        img_bytes = base64.b64decode(img_b64)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.write(img_bytes)
        tmp.close()

        return tmp.name

    except Exception:
        logger.exception("Image generation failed")
        return None


# ------------------------------------------------------------
# PPT BUILDER (TEXT WIDTH & IMAGE LAYOUT)
# ------------------------------------------------------------

def build_ppt(slides, template_style):

    prs = Presentation()

    for s in slides:

        slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Apply template styling
        apply_template_style(slide, template_style)

        # Title
        slide.shapes.title.text = s.get("title", "Slide")

        # Body
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()

        for b in s.get("bullets", []):
            p = tf.add_paragraph()
            p.text = b

        # Body position
        body.top = Inches(1.5)
        body.left = Inches(0.7)

        img_path = s.get("image_path")

        # No image => full width body
        if not img_path:
            body.width = prs.slide_width - Inches(1.4)
            continue

        # With image => shrink body to left
        body.width = prs.slide_width - Inches(4.5)

        try:
            img = Image.open(img_path)
            w, h = img.size
            aspect = w / h

            max_w = Inches(3.5)
            max_h = Inches(3.0)

            if aspect > 1:
                final_w = max_w
                final_h = max_w / aspect
            else:
                final_h = max_h
                final_w = max_h * aspect

            left = prs.slide_width - final_w - Inches(0.5)
            top = slide.shapes.title.top + slide.shapes.title.height + Inches(0.2)

            slide.shapes.add_picture(
                img_path,
                left=left,
                top=top,
                width=final_w,
                height=final_h
            )

        except Exception:
            logger.exception("Image placement failed")

    out_path = os.path.join(
        tempfile.gettempdir(),
        f"generated_{uuid.uuid4().hex[:8]}.pptx"
    )
    prs.save(out_path)
    return out_path


# ------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------

def generate_presentation(prompt: str,
                          template_style="corporate",
                          requested_num_slides=None):

    detected_slides = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected_slides or 5

    # Semantic search from Chroma
    refs = semantic_search(prompt, top_k=5)

    if not refs:
        return None, {
            "status": "no_match",
            "error": True,
            "message": "No matching knowledge found. Please refine your prompt."
        }

    reference_text = [r["text"] for r in refs if r.get("text")]

    # Build slide plan (LLM or fallback)
    plan = call_llm_plan(
        prompt=prompt,
        style=template_style,
        references_text=reference_text,
        num_slides=num_slides,
    )

    slides = []
    force_images = "image" in prompt.lower() or "images" in prompt.lower()

    for sp in plan:
        img = None
        if sp.get("visual_required") or force_images:
            img = generate_visual_image(sp.get("visual_prompt"))

        slides.append({
            "title": sp.get("title", "Slide"),
            "bullets": sp.get("bullets", []),
            "image_path": img
        })

    out_path = build_ppt(slides, template_style)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "template": template_style,
        "ppt_file": fname
    }

    upload_json_to_blob(
        json.dumps(log, indent=2).encode("utf-8"),
        f"logs/{fname}.json"
    )

    return out_path, log
