# ============================================================
# generate_ppt.py – FINAL STABLE FIX (NO EMPTY PPT, NO CRASH)
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
# INITIAL SETUP
# ------------------------------------------------------------
ensure_dir("design_jsons")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)


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
# ✅ FIXED GPT PLAN GENERATOR (NO REQUIRED design_context)
# ------------------------------------------------------------
def call_llm_plan(prompt, style,
                  design_context=None, references_text=None,
                  num_slides=None, theme=None):

    design_context = design_context or []
    references_text = references_text or []

    sys_prompt = (
        "You are a professional presentation generator.\n"
        "Return STRICT JSON ONLY in this format:\n"
        "[{\"title\": str, \"bullets\": [str], \"visual_required\": bool, \"visual_prompt\": str}]\n"
        "If user mentions images → set visual_required=true.\n"
        "If no relevant data is found, still generate a general informative presentation.\n"
        "Never return empty JSON.\n"
    )

    user_prompt = f"Create a {style} presentation plan for: {prompt}"

    if num_slides:
        user_prompt += f". Make exactly {num_slides} slides."

    try:
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt}
            ],
            max_completion_tokens=1500,
            temperature=1
        )

        raw = resp.choices[0].message.content
        plan = safe_json_load(raw)

        if not isinstance(plan, list) or len(plan) == 0:
            raise ValueError("Invalid or empty plan JSON")

        return plan

    except Exception as e:
        logger.warning(f"[FALLBACK] Using default slide because LLM failed: {e}")

        return [{
            "title": "Introduction",
            "bullets": [
                f"Overview of: {prompt}",
                "Key objectives",
                "High-level summary"
            ],
            "visual_required": False,
            "visual_prompt": ""
        }]


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------
def generate_visual_image(prompt: str):
    img_prompt = (prompt or "") + " Minimal professional illustration. No text labels."

    try:
        resp = image_client.images.generate(
            model=IMAGE_MODEL,
            prompt=img_prompt,
            size="1024x1024"
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
# STYLE HELPER
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
    except:
        return None


# ------------------------------------------------------------
# PPT BUILDER (AUTO FULL WIDTH WHEN NO IMAGE)
# ------------------------------------------------------------
def build_ppt(slides):
    prs = Presentation()

    for s in slides:
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)

        # Title
        slide.shapes.title.text = s.get("title", "")

        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()

        for b in s.get("bullets", []):
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(18)

        body.top += Inches(0.3)

        # ✅ NO IMAGE → FULL WIDTH TEXT
        if not s.get("image_path"):
            body.left = Inches(0.5)
            body.width = prs.slide_width - Inches(1)
            continue

        # IMAGE MODE → SPLIT LAYOUT
        body.width = prs.slide_width - Inches(4)

        try:
            img_path = s["image_path"]
            img = Image.open(img_path)
            w, h = img.size
            aspect = w / h

            max_w = Inches(3.2)
            max_h = Inches(2.8)

            if aspect >= 1:
                final_w = max_w
                final_h = final_w / aspect
            else:
                final_h = max_h
                final_w = final_h * aspect

            left = prs.slide_width - final_w - Inches(0.5)
            top = slide.shapes.title.top + slide.shapes.title.height + Inches(0.3)

            slide.shapes.add_picture(img_path, left, top, width=final_w, height=final_h)

        except Exception:
            logger.exception("Image placement failed")

    out_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out_path)
    return out_path


# ------------------------------------------------------------
# ✅ MAIN GENERATION PIPELINE (ALWAYS RETURNS VALID PPT)
# ------------------------------------------------------------
def generate_presentation(prompt: str, style="Auto", requested_num_slides=None,
                          theme=None, tag_filters=None):

    detected_slides, detected_theme = parse_user_intent(prompt)
    requested_num_slides = requested_num_slides or detected_slides
    theme = theme or detected_theme

    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []

    # ✅ If nothing found → still allow generation
    if not refs:
        logger.warning("No Chroma matches found — generating generic presentation.")

    plan = call_llm_plan(
        prompt, style,
        design_context=None,
        references_text=None,
        num_slides=requested_num_slides,
        theme=theme
    )

    if not plan:
        raise RuntimeError("LLM returned empty plan even after fallback.")

    force_images = "image" in prompt.lower() or "images" in prompt.lower()

    slides = []

    for sp in plan:
        img = None
        if sp.get("visual_required") or force_images:
            img = generate_visual_image(sp.get("visual_prompt"))

        slides.append({
            "title": sp.get("title"),
            "bullets": sp.get("bullets", []),
            "image_path": img
        })

    out_path = build_ppt(slides)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "ppt_file": fname,
        "chroma_matches": len(refs)
    }

    upload_json_to_blob(json.dumps(log, indent=2).encode(), f"logs/{fname}.json")

    return out_path, log
