# ============================================================
# generate_ppt.py – Semantic + Fixed Layout + UI Friendly
# ============================================================

import os
import tempfile
import uuid
import json
import re
import base64
from typing import Optional

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
def parse_user_intent(prompt: str) -> Optional[int]:
    """Try to detect 'N slides' from user prompt."""
    match = re.search(r"(\d+)\s+slides?", prompt.lower())
    if match:
        return int(match.group(1))
    return None


# ------------------------------------------------------------
# LLM PLAN GENERATOR – JSON FORCED & ROBUST
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

    Returns a list of dicts: {title, bullets, visual_required, visual_prompt}.
    If anything goes wrong, we fall back to a simple deterministic plan.
    ALWAYS returns a list whose length == num_slides (if num_slides is given).
    """
    references_text = references_text or []

    def fallback_plan(n: Optional[int]):
        n = n or 3
        fb = []
        for i in range(n):
            fb.append(
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
        return fb

    density_instructions = {
        "Minimal": "Use at most 1–2 short bullet points per slide.",
        "Concise": "Use about 3 bullet points per slide.",
        "Detailed": "Use about 5 bullet points per slide.",
        "Extensive": "Use 6–8 detailed bullet points per slide.",
    }
    density_line = density_instructions.get(text_density, "")

    sys_prompt = (
        "You are a presentation planner.\n"
        "You MUST respond with JSON ONLY.\n\n"
        "Return either:\n"
        "1) An object of the form:\n"
        "{\n"
        '  \"slides\": [\n'
        '    {\n'
        '      \"title\": \"string\",\n'
        '      \"bullets\": [\"string\", \"string\", ...],\n'
        '      \"visual_required\": true or false,\n'
        '      \"visual_prompt\": \"string\"\n'
        "    }, ...\n"
        "  ]\n"
        "}\n"
        "or 2) a bare JSON array:\n"
        "[{...}, {...}, ...]\n\n"
        "Rules:\n"
        "- No markdown, no comments, no text outside the JSON.\n"
        "- Every bullet must be a plain string (no bullet symbols or numbering).\n"
        "- If the user asks for images, set visual_required=true.\n"
        "- Do NOT put any written text or labels inside the generated images.\n"
        f"- Text density guideline: {density_line}\n\n"
        "You may use these reference snippets as guidance:\n"
        f"{json.dumps(references_text)[:2000]}"
    )

    user_prompt = f"Create a professional presentation plan for: {prompt}"
    if num_slides:
        user_prompt += f" Use exactly {num_slides} slides."

    try:
        # 👉 Force JSON output from the model
        resp = text_client.chat.completions.create(
            model=CHAT_MODEL,
            messages=[
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": user_prompt},
            ],
            response_format={"type": "json_object"},
            max_completion_tokens=1200,
            temperature=1,
        )

        msg = resp.choices[0].message

        # Newer clients may provide a parsed object
        if hasattr(msg, "parsed") and msg.parsed is not None:
            parsed = msg.parsed
        else:
            raw = (msg.content or "").strip()
            parsed = safe_json_load(raw)

        # Try to locate slide list
        slides_raw = None
        if isinstance(parsed, dict) and isinstance(parsed.get("slides"), list):
            slides_raw = parsed["slides"]
        elif isinstance(parsed, list):
            slides_raw = parsed
        elif isinstance(parsed, dict):
            # last resort: walk values to find a list
            for v in parsed.values():
                if isinstance(v, list):
                    slides_raw = v
                    break

        if not isinstance(slides_raw, list) or not slides_raw:
            raise ValueError("Could not find slide list in JSON")

        # Normalize items into slide dicts
        cleaned = []
        for idx, item in enumerate(slides_raw):
            # dict case
            if isinstance(item, dict):
                title = item.get("title") or f"Slide {idx + 1}"
                bullets = item.get("bullets", [])
                if isinstance(bullets, str):
                    bullets = [bullets]
                elif not isinstance(bullets, list):
                    bullets = []
                else:
                    bullets = [str(b).strip() for b in bullets if str(b).strip()]

                visual_required = bool(item.get("visual_required", False))
                visual_prompt = item.get("visual_prompt", "")
            # string → title + single bullet
            elif isinstance(item, str):
                txt = item.strip()
                title = txt[:40] + "..." if len(txt) > 40 else (txt or f"Slide {idx + 1}")
                bullets = [txt] if txt else []
                visual_required = False
                visual_prompt = ""
            # list → bullets only
            elif isinstance(item, list):
                bullets = [str(x).strip() for x in item if str(x).strip()]
                title = f"Slide {idx + 1}"
                visual_required = False
                visual_prompt = ""
            else:
                title = f"Slide {idx + 1}"
                bullets = []
                visual_required = False
                visual_prompt = ""

            cleaned.append(
                {
                    "title": str(title),
                    "bullets": bullets,
                    "visual_required": visual_required,
                    "visual_prompt": str(visual_prompt),
                }
            )

        if not cleaned:
            raise ValueError("Cleaned slide list is empty")

        plan = cleaned

        # --- ENFORCE SLIDE COUNT ---
        if num_slides and num_slides > 0:
            if len(plan) > num_slides:
                plan = plan[:num_slides]
            elif len(plan) < num_slides:
                while len(plan) < num_slides:
                    last = dict(plan[-1])
                    last["title"] = last.get("title", "Slide") + " (cont.)"
                    plan.append(last)

        return plan

    except Exception as e:
        logger.warning(f"Invalid plan JSON, using fallback plan: {e}")
        return fallback_plan(num_slides)


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
# PPT BUILDER (FULL-WIDTH TEXT WHEN NO IMAGE)
# ------------------------------------------------------------
def build_ppt(slides):
    prs = Presentation()

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
    template_style=None,  # accepted for UI compatibility, not used yet
):
    """
    Main entry used by app.py / test scripts.

    Returns:
        (ppt_path or None, log_dict)
        - If log_dict["error"] is True, ppt_path will be None and no PPT should be offered.
    """

    # 1) Semantic search in Chroma
    refs = semantic_search(prompt, top_k=5, tags=tag_filters) or []

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

    # 6) Build PPT
    out_path = build_ppt(slides)

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
