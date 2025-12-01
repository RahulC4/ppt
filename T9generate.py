# ============================================================
# generate_ppt.py – Semantic + Fixed Layout + UI Friendly (STABLE)
# ============================================================

import os
import tempfile
import uuid
import json
import re
import base64

from pptx import Presentation
from pptx.util import Inches, Pt
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

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def parse_user_intent(prompt: str):
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    if match:
        return int(match.group(1))
    return None


# ------------------------------------------------------------
# LLM PLAN GENERATOR (FULLY SAFE)
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
        "Do NOT include markdown or commentary.\n"
        f"{density_line}\n\n"
        f"Reference content:\n{json.dumps(references_text)[:2000]}"
    )

    user_prompt = f"Create a presentation plan for: {prompt}"
    if num_slides:
        user_prompt += f" Use exactly {num_slides} slides."

    # ---------- FALLBACK ----------
    def fallback(n):
        n = n or 5
        slides = []
        for i in range(n):
            slides.append({
                "title": f"Slide {i+1}",
                "bullets": [f"Key point 1", f"Key point 2"],
                "visual_required": False,
                "visual_prompt": ""
            })
        return slides

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
        plan_raw = safe_json_load(raw)

        if not isinstance(plan_raw, list):
            raise ValueError("Invalid plan JSON")

        cleaned = []

        MAX_BULLETS = 6 if text_density in ["Detailed", "Extensive"] else 4
        SAFE_CHAR_LIMIT = 140

        for item in plan_raw:
            if not isinstance(item, dict):
                continue

            raw_bullets = item.get("bullets", [])
            clean_bullets = []

            for b in raw_bullets[:MAX_BULLETS]:
                text = str(b).strip()
                if len(text) > SAFE_CHAR_LIMIT:
                    text = text[:SAFE_CHAR_LIMIT] + "..."
                clean_bullets.append(text)

            cleaned.append({
                "title": str(item.get("title", "Untitled"))[:80],
                "bullets": clean_bullets,
                "visual_required": bool(item.get("visual_required", False)),
                "visual_prompt": str(item.get("visual_prompt", ""))[:200],
            })

        if not cleaned:
            raise ValueError("Cleaned plan empty")

        # ✅ ENFORCE SLIDE COUNT
        if num_slides:
            if len(cleaned) > num_slides:
                cleaned = cleaned[:num_slides]
            while len(cleaned) < num_slides:
                last = dict(cleaned[-1])
                last["title"] += " (cont.)"
                cleaned.append(last)

        return cleaned

    except Exception as e:
        logger.warning(f"Invalid plan JSON, using fallback: {e}")
        return fallback(num_slides)


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    if not prompt:
        return None

    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=prompt,
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
# PPT BUILDER (STABLE FOR EXTENSIVE MODE)
# ------------------------------------------------------------
def build_ppt(slides):
    prs = Presentation()

    for sp in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = sp.get("title", "")

        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.auto_size = None

        for i, b in enumerate(sp.get("bullets", [])):
            if i == 0:
                tf.text = b
                tf.paragraphs[0].font.size = Pt(18)
            else:
                p = tf.add_paragraph()
                p.text = b
                p.font.size = Pt(18)

        body.top = slide.shapes.title.top + slide.shapes.title.height + Inches(0.3)

        img_path = sp.get("image_path")

        if img_path:
            try:
                img = Image.open(img_path)
                w, h = img.size
                aspect = w / h if h else 1.0

                max_w = Inches(3)
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
    template_style=None,
):

    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []
    if not refs:
        return None, {"error": True, "message": "No relevant content found."}

    reference_text = [(r.get("text") or "")[:500] for r in refs if r.get("text")]

    detected_slides = parse_user_intent(prompt)
    num_slides = requested_num_slides or detected_slides or 5

    plan = call_llm_plan(
        prompt=prompt,
        style=style,
        references_text=reference_text,
        num_slides=num_slides,
        theme=theme,
        text_density=text_density,
    )

    slides = []
    for sp in plan:
        img_path = generate_visual_image(sp.get("visual_prompt")) if sp.get("visual_required") else None
        slides.append({
            "title": sp.get("title"),
            "bullets": sp.get("bullets"),
            "image_path": img_path,
        })

    out_path = build_ppt(slides)

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

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
