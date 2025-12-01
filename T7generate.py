# ============================================================
# generate_ppt.py – Semantic + Fixed Templates + Safe Pipeline
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
# CONFIG
# ------------------------------------------------------------
ensure_dir("generated")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)

# Simple similarity filter – lower = stricter. You can tune via .env if needed
SIMILARITY_THRESHOLD = float(get_env("SIMILARITY_THRESHOLD", "1.1"))

# ------------------------------------------------------------
# FIXED TEMPLATE PRESETS (Gamma-style)
# ------------------------------------------------------------
TEMPLATE_PRESETS = {
    "Plant": {
        "background_color": "#F7FAF0",
        "title_font": "Calibri",
        "title_color": "#111827",
        "body_font": "Calibri",
        "body_color": "#111827",
    },
    "Dark": {
        "background_color": "#111827",
        "title_font": "Segoe UI",
        "title_color": "#F9FAFB",
        "body_font": "Segoe UI",
        "body_color": "#E5E7EB",
    },
    "Minimal": {
        "background_color": "#FFFFFF",
        "title_font": "Arial",
        "title_color": "#111827",
        "body_font": "Arial",
        "body_color": "#111827",
    },
    "Corporate": {
        "background_color": "#F3F4F6",
        "title_font": "Calibri",
        "title_color": "#0F172A",
        "body_font": "Calibri",
        "body_color": "#111827",
    },
}


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def parse_user_intent(prompt: str):
    """Try to detect 'N slides' from user prompt."""
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    if match:
        return int(match.group(1))
    return None


def parse_hex_color(hex_str):
    try:
        if not hex_str:
            return None
        hex_str = hex_str.replace("#", "")
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
        return RGBColor(r, g, b)
    except Exception:
        return None


def apply_template_style(slide, template_style: str):
    """Apply background + font style based on preset."""
    preset = TEMPLATE_PRESETS.get(template_style)
    if not preset:
        return

    bg_color = parse_hex_color(preset.get("background_color"))
    title_color = parse_hex_color(preset.get("title_color"))
    body_color = parse_hex_color(preset.get("body_color"))
    title_font = preset.get("title_font")
    body_font = preset.get("body_font")

    # Background
    if bg_color:
        try:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = bg_color
        except Exception:
            pass

    # Title font + color
    try:
        title_shape = slide.shapes.title
        if title_shape and title_shape.has_text_frame:
            for p in title_shape.text_frame.paragraphs:
                for r in p.runs:
                    if title_font:
                        r.font.name = title_font
                    if title_color:
                        r.font.color.rgb = title_color
                    r.font.bold = True
    except Exception:
        pass

    # Body font + color
    try:
        # usually placeholder 1 is body
        body_shape = slide.placeholders[1]
        if body_shape and body_shape.has_text_frame:
            for p in body_shape.text_frame.paragraphs:
                for r in p.runs:
                    if body_font:
                        r.font.name = body_font
                    if body_color:
                        r.font.color.rgb = body_color
    except Exception:
        pass


# ------------------------------------------------------------
# LLM PLAN GENERATOR
# ------------------------------------------------------------
def call_llm_plan(
    prompt,
    style,
    design_context=None,
    references_text=None,
    num_slides=None,
    theme=None,
    text_density=None,
):
    """
    Ask GPT to create a slide plan.
    We return a list of dicts: {title, bullets, visual_required, visual_prompt}.
    If anything goes wrong, we fall back to a simple deterministic plan.
    """
    references_text = references_text or []

    density_instructions = {
        "Minimal": "Use at most 1–2 short bullet points per slide.",
        "Concise": "Use about 3 bullet points per slide.",
        "Detailed": "Use about 5 bullet points per slide.",
        "Extensive": "Use 6–8 detailed bullet points per slide.",
    }
    density_line = density_instructions.get(text_density, "")

    sys_prompt = (
        "You are a presentation planner.\n"
        "Return STRICT JSON ONLY in this exact format:\n"
        "[{\"title\": str, \"bullets\": [str], "
        "\"visual_required\": bool, \"visual_prompt\": str }]\n"
        "If the user asks for images, set visual_required=true.\n"
        "Do NOT put any words or labels inside visual_prompt images.\n"
        f"{density_line}\n\n"
        "You MUST ground content in the reference snippets when possible.\n"
        "Reference snippets (may be truncated):\n"
        f"{json.dumps(references_text)[:2500]}"
    )

    user_prompt = f"Create a professional presentation plan for: {prompt}"
    if num_slides:
        user_prompt += f" Use exactly {num_slides} slides."

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt},
            ],
            max_completion_tokens=1200,
            temperature=1,
        )

        raw = resp.choices[0].message.content
        plan = safe_json_load(raw)

        if not isinstance(plan, list) or len(plan) == 0:
            raise ValueError("Invalid or empty plan JSON")

        return plan

    except Exception as e:
        logger.warning(f"Invalid plan JSON, using fallback plan: {e}")

        # Fallback: simple generic slides so we never crash
        n = num_slides or 3
        fallback = []
        for i in range(n):
            fallback.append(
                {
                    "title": f"Slide {i + 1}",
                    "bullets": [
                        f"Key point 1 for slide {i + 1}",
                        f"Key point 2 for slide {i + 1}",
                    ],
                    "visual_required": False,
                    "visual_prompt": "",
                }
            )
        return fallback


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    """Generate an image file from GPT image model. Returns a temp file path or None."""
    if not prompt:
        return None

    img_prompt = prompt + " Minimal, professional illustration. No text labels."

    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=img_prompt,
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
# PPT BUILDER (FULL-WIDTH TEXT WHEN NO IMAGE) + TEMPLATE
# ------------------------------------------------------------
def build_ppt(slides, template_style="Auto"):
    prs = Presentation()

    effective_template = template_style if template_style != "Auto" else None

    for sp in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Title
        slide.shapes.title.text = sp.get("title", "")

        # Body placeholder
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()

        for b in sp.get("bullets", []):
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(20)

        # Space under title
        body.top = slide.shapes.title.top + slide.shapes.title.height + Inches(0.3)

        img_path = sp.get("image_path")

        if not img_path:
            # NO IMAGE → full width
            body.left = Inches(0.5)
            body.width = prs.slide_width - Inches(1.0)

            # Apply overall template after text is in place
            if effective_template:
                apply_template_style(slide, effective_template)
            continue

        # IMAGE PRESENT → split layout
        body.left = Inches(0.5)
        body.width = prs.slide_width - Inches(4.0)

        try:
            img = Image.open(img_path)
            w, h = img.size
            aspect = w / h if h else 1.0

            max_w = Inches(3.0)
            max_h = Inches(2.5)

            if aspect >= 1:
                final_w = max_w
                final_h = final_w / aspect
            else:
                final_h = max_h
                final_w = final_h * aspect

            left = prs.slide_width - final_w - Inches(0.5)
            top = body.top

            slide.shapes.add_picture(img_path, left, top, width=final_w, height=final_h)
        except Exception:
            logger.exception("Image placement failed")

        # Apply template after everything is added
        if effective_template:
            apply_template_style(slide, effective_template)

    out_path = os.path.join(
        tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx"
    )
    prs.save(out_path)
    return out_path


# ------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------
def generate_presentation(
    prompt: str,
    style="Auto",
    requested_num_slides=None,
    theme=None,
    text_density="Concise",
    tag_filters=None,
    template_style="Auto",
):
    """
    Main entry used by app.py / test scripts.

    Returns:
        (ppt_path or None, log_dict)
        - If log_dict["error"] is True, ppt_path will be None and no PPT should be offered.
    """

    # 1) Semantic search in Chroma
    raw_refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []

    # Optional similarity filter if score present
    refs = [
        r for r in raw_refs
        if r.get("score") is None or r["score"] <= SIMILARITY_THRESHOLD
    ]

    if not refs:
        msg = (
            "I couldn’t find any relevant content in your uploaded presentations "
            "for this prompt. Please try a prompt related to the topics in your sample PPTs."
        )
        logger.warning(msg)
        return None, {"error": True, "message": msg}

    # 2) Build reference text snippets
    reference_text = []
    for r in refs:
        snippet = (r.get("text") or "")[:500]
        if snippet:
            reference_text.append(snippet)

    logger.info(f"Using {len(reference_text)} reference snippets from Chroma.")

    # 3) Slide count from prompt / UI
    detected_slides = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected_slides or 5

    # 4) Ask LLM for slide plan
    plan = call_llm_plan(
        prompt=prompt,
        style=style,
        references_text=reference_text,
        num_slides=num_slides,
        theme=theme,
        text_density=text_density,
    )

    # Validate plan: if everything is empty, bail out with a friendly message
    if not plan or not isinstance(plan, list):
        msg = (
            "I couldn't generate a valid slide plan from your prompt and the "
            "available sample PPTs. Please simplify or rephrase your request."
        )
        logger.warning(msg)
        return None, {"error": True, "message": msg}

    all_empty = True
    for s in plan:
        title = (s.get("title") or "").strip()
        bullets = [str(b).strip() for b in (s.get("bullets") or []) if str(b).strip()]
        if title or bullets:
            all_empty = False
            break

    if all_empty:
        msg = (
            "I generated only empty slides for this request. "
            "Please try a clearer prompt or one that is closer to your sample PPT content."
        )
        logger.warning(msg)
        return None, {"error": True, "message": msg}

    # 5) Build slide objects (and optionally generate images)
    force_images = "image" in prompt.lower() or "images" in prompt.lower()
    slides = []

    for sp in plan:
        img_path = None
        if sp.get("visual_required") or force_images:
            img_path = generate_visual_image(sp.get("visual_prompt"))

        slides.append(
            {
                "title": sp.get("title", "Untitled"),
                "bullets": sp.get("bullets", []),
                "image_path": img_path,
            }
        )

    # 6) Build PPT with selected template
    out_path = build_ppt(slides, template_style=template_style)

    # 7) Upload to Blob + log
    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "ppt_file": fname,
        "error": False,
        "text_density": text_density,
        "template_style": template_style,
    }

    upload_json_to_blob(
        json.dumps(log, indent=2).encode("utf-8"),
        f"logs/{fname}.json",
    )

    return out_path, log
