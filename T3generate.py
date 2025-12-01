# ============================================================
# generate_ppt.py – FIXED TEMPLATE + SEMANTIC CONTENT ONLY
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

# We no longer use design_jsons, but keep folder in case of future usage
ensure_dir("design_jsons")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)

# Distance threshold for Chroma matches (smaller = closer)
SIMILARITY_THRESHOLD = float(get_env("SIMILARITY_THRESHOLD", "1.1"))

# ------------------------------------------------------------
# FIXED TEMPLATE PRESETS (Gamma-like)
# ------------------------------------------------------------
TEMPLATE_PRESETS = {
    "Plant": {
        "background_color": "#F7FAF0",
        "title_font": "Calibri",
        "body_font": "Calibri",
    },
    "Dark": {
        "background_color": "#1F2933",
        "title_font": "Segoe UI",
        "body_font": "Segoe UI",
    },
    "Minimal": {
        "background_color": "#FFFFFF",
        "title_font": "Arial",
        "body_font": "Arial",
    },
    "Corporate": {
        "background_color": "#F3F4F6",
        "title_font": "Calibri",
        "body_font": "Calibri",
    }
}


# ------------------------------------------------------------
# USER INTENT PARSER
# ------------------------------------------------------------
def parse_user_intent(prompt: str):
    num_slides = None
    theme = None

    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    if match:
        num_slides = int(match.group(1))

    for t in ["corporate", "modern", "minimal", "professional", "dark", "light"]:
        if t in prompt.lower():
            theme = t.capitalize()
            break

    return num_slides, theme


# ------------------------------------------------------------
# GPT PLAN GENERATOR (REFERENCE-ONLY)
# ------------------------------------------------------------
def call_llm_plan(
    prompt: str,
    style: str,
    references_text,
    num_slides: int | None = None,
    text_density: str = "Concise",
):
    density_rules = {
        "Minimal": "Use at most 1–2 short bullet points per slide.",
        "Concise": "Use about 3 bullet points per slide.",
        "Detailed": "Use about 5 bullet points per slide.",
        "Extensive": "Use 6–8 detailed bullet points per slide.",
    }

    density_text = density_rules.get(text_density, "")

    sys_prompt = (
        "You are a senior presentation designer.\n"
        "You MUST use ONLY the provided reference snippets from existing slides.\n"
        "Do NOT invent new topics that are not present in the references.\n\n"
        "Return STRICT JSON ONLY in this exact format:\n"
        "[{\"title\": str, \"bullets\": [str], \"visual_required\": bool, "
        "\"visual_prompt\": str }]\n\n"
        "If the user asks for images → set visual_required=true.\n"
        "Do NOT include any text inside visual_prompt (no labels / words on the image).\n"
        f"{density_text}\n\n"
        "REFERENCE SNIPPETS (may be truncated):\n"
        f"{json.dumps(references_text)[:2500]}"
    )

    user_prompt = f"Create a {style} presentation plan for this prompt: {prompt}"
    if num_slides:
        user_prompt += f". Use exactly {num_slides} slides."

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt},
            ],
            max_completion_tokens=1500,
            temperature=0.4,  # lower for more stable JSON
        )

        plan = safe_json_load(resp.choices[0].message.content)

        return plan

    except Exception:
        logger.exception("Plan generation failed")
        return None


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    img_prompt = (prompt or "") + " Minimal professional illustration. No text labels."

    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=img_prompt,
            size="1024x1024",
        )

        b64 = getattr(resp.data[0], "b64_json", None)
        if b64:
            img_bytes = base64.b64decode(b64)
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(img_bytes)
            tmp.close()
            return tmp.name

        url = getattr(resp.data[0], "url", None)
        if url:
            import requests
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            tmp.write(requests.get(url).content)
            tmp.close()
            return tmp.name

        return None

    except Exception:
        logger.exception("Image generation failed")
        return None


# ------------------------------------------------------------
# COLOR PARSER + TEMPLATE APPLY
# ------------------------------------------------------------
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
    preset = TEMPLATE_PRESETS.get(template_style)
    if not preset:
        return

    # Background color
    bg = parse_hex_color(preset.get("background_color"))
    if bg:
        try:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = bg
        except Exception:
            pass

    title_font = preset.get("title_font")
    body_font = preset.get("body_font")

    # Title font
    try:
        if title_font:
            title = slide.shapes.title
            for p in title.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = title_font
    except Exception:
        pass

    # Body font
    try:
        if body_font:
            body = slide.placeholders[1]
            for p in body.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = body_font
    except Exception:
        pass


# ------------------------------------------------------------
# PPT BUILDER (TEXT FULL-WIDTH WHEN NO IMAGE)
# ------------------------------------------------------------
def build_ppt(slides, template_style="Auto"):
    prs = Presentation()

    # If template_style is "Auto", just keep PowerPoint default theme
    effective_template = template_style if template_style != "Auto" else None

    for s in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        # Apply fixed template style if chosen
        if effective_template:
            apply_template_style(slide, effective_template)

        # Title
        slide.shapes.title.text = s.get("title", "")

        # Body
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()

        for b in s.get("bullets", []):
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(18)

        body.top = slide.shapes.title.top + slide.shapes.title.height + Inches(0.3)

        # NO IMAGE → full-width text
        if not s.get("image_path"):
            body.left = Inches(0.5)
            body.width = prs.slide_width - Inches(1.0)
            body.height = prs.slide_height - Inches(1.2)
            tf.word_wrap = True
            continue

        # IMAGE EXISTS → text left, image right
        body.left = Inches(0.5)
        body.width = prs.slide_width - Inches(4.0)
        tf.word_wrap = True

        try:
            img_path = s["image_path"]
            img = Image.open(img_path)
            w, h = img.size
            aspect = w / h

            max_w = Inches(3.0)
            max_h = Inches(2.8)

            if aspect >= 1:
                final_w = max_w
                final_h = final_w / aspect
            else:
                final_h = max_h
                final_w = final_h * aspect

            left = prs.slide_width - final_w - Inches(0.4)
            top = body.top

            slide.shapes.add_picture(img_path, left, top, width=final_w, height=final_h)
        except Exception:
            logger.exception("Image placement failed")

    out_path = os.path.join(tempfile.gettempdir(), f"generated_presentation_{uuid.uuid4().hex[:8]}.pptx")
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
    tag_filters=None,
    text_density="Concise",
    template_style="Auto",
):
    # 1) Parse user intent for slide count
    detected_slides, detected_theme = parse_user_intent(prompt)
    requested_num_slides = requested_num_slides or detected_slides
    theme = theme or detected_theme

    # 2) Semantic search from Chroma (content only)
    raw_refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []

    # Filter by similarity threshold (if score present)
    refs = [
        r for r in raw_refs
        if r.get("score") is None or r["score"] <= SIMILARITY_THRESHOLD
    ]

    if not refs:
        msg = (
            "I couldn’t find any relevant content in your uploaded sample presentations "
            "for this request. Please rephrase your prompt using topics related to your "
            "existing PPTs, or upload a new sample PPT."
        )
        logger.warning(msg)
        return None, {"error": True, "message": msg}

    # 3) Build reference text from retrieved slides
    reference_text = []
    for r in refs:
        snippet = (r.get("text") or "")[:500]
        if snippet:
            reference_text.append(snippet)

    logger.info(f"Using {len(reference_text)} reference snippets from Chroma.")

    # 4) LLM plan generation based ONLY on references
    plan = call_llm_plan(
        prompt=prompt,
        style=style,
        references_text=reference_text,
        num_slides=requested_num_slides,
        text_density=text_density,
    )

    # 5) Validate plan to avoid empty PPT
    if not plan:
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

    # 6) Optional image generation
    force_images = "image" in prompt.lower() or "images" in prompt.lower()

    slides = []
    for sp in plan:
        img = None
        if sp.get("visual_required") or force_images:
            img = generate_visual_image(sp.get("visual_prompt"))

        slides.append(
            {
                "title": sp.get("title"),
                "bullets": sp.get("bullets", []),
                "image_path": img,
            }
        )

    # 7) Build PPT with fixed template
    out_path = build_ppt(slides, template_style=template_style)

    # 8) Upload & log
    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "ppt_file": fname,
        "text_density": text_density,
        "template_style": template_style,
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
