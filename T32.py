# ============================================================
# UPDATED MODULES: app.py + generate_ppt.py
# FEATURE: Visual Templates Loaded From Azure Blob
# ============================================================

# =====================
# app.py (UPDATED)
# =====================

import streamlit as st

from generate_ppt import generate_presentation
from search_utils import semantic_search
from utils import logger, get_env
from azure_blob_utils import upload_source_ppt_to_blob, list_generated_presentations
from ingestion_chroma import process_blob as ingest_process_blob

st.set_page_config(page_title="AI PPT Generator", layout="wide", page_icon="📊")

st.title("📊 AI PowerPoint Generator")
st.write(
    "Generate PPT decks based on your existing slides. "
    "Upload sample PPTs, select a visual template, and generate new decks."
)

SIMILARITY_THRESHOLD = float(get_env("SIMILARITY_THRESHOLD", "1.1"))

st.markdown("---")

# ============================================================
# 1️⃣ UPLOAD SAMPLE PPTs (KNOWLEDGE BASE)
# ============================================================

st.subheader("1️⃣ Upload Sample PPT Files (Knowledge Base)")

uploaded_files = st.file_uploader(
    "Upload .pptx files to use as content references:",
    type=["pptx"],
    accept_multiple_files=True,
)

if st.button("📥 Add to Knowledge Base") and uploaded_files:
    with st.spinner("Uploading & indexing PPTs..."):
        for upl in uploaded_files:
            try:
                bytes_data = upl.read()
                blob_name = upl.name

                # Upload to blob
                upload_source_ppt_to_blob(bytes_data, blob_name)
                # Ingest into Chroma
                ingest_process_blob(blob_name)

                st.success(f"✅ Processed & indexed: {blob_name}")
            except Exception as e:
                logger.exception(f"Failed to process {upl.name}")
                st.error(f"❌ Error processing {upl.name}: {e}")

st.markdown("---")

# ============================================================
# 2️⃣ CONFIGURE PRESENTATION
# ============================================================

st.subheader("2️⃣ Create New Presentation")

prompt = st.text_area(
    "Enter your presentation prompt:",
    placeholder="Example: Create a 5-slide presentation about claims automation.",
    height=150,
)

col1, col2 = st.columns(2)

with col1:
    num_slides = st.number_input(
        "Number of Slides",
        min_value=1,
        max_value=25,
        value=5,
        step=1,
    )

# ✅ Visual Templates from Azure Blob
with col2:
    blob_templates = list_generated_presentations()
    template_options = ["Plain (Default)"] + blob_templates

    template_style = st.selectbox(
        "Visual Template (from Blob)",
        template_options,
        index=0,
    )

# ✅ Global Image Toggle
force_images = st.checkbox("Generate images for all slides", value=False)

st.markdown("---")

# ============================================================
# 🎯 GENERATE & DOWNLOAD PPT
# ============================================================

if st.button("🎯 Generate & Download PPT"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
    else:
        with st.spinner("Generating final PowerPoint..."):
            try:
                ppt_path, log = generate_presentation(
                    prompt=prompt,
                    requested_num_slides=num_slides,
                    template_style=template_style,
                    force_images=force_images,
                )

                if log.get("error"):
                    st.warning(f"⚠️ {log.get('message')}")
                else:
                    st.success("✅ PPT Generated Successfully!")
                    with open(ppt_path, "rb") as f:
                        st.download_button(
                            label="⬇️ Download PPT",
                            data=f,
                            file_name="generated_presentation.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )

                    st.subheader("📄 Generation Log")
                    st.json(log)

            except Exception as e:
                logger.exception("PPT generation failed")
                st.error(f"Failed to generate PPT: {e}")


# =====================
# generate_ppt.py (UPDATED)
# =====================

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
from azure_blob_utils import upload_ppt_to_blob

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
ensure_dir("generated")

CHAT_MODEL = get_env("CHAT_MODEL", required=True)
IMAGE_MODEL = get_env("IMAGE_MODEL", required=True)


# ------------------------------------------------------------
# LLM PLAN GENERATOR (ROBUST)
# ------------------------------------------------------------

def call_llm_plan(prompt, references_text, num_slides):
    sys_prompt = (
        "You are a presentation planner.\n"
        "Return STRICT JSON ONLY in this format:\n"
        "[{\"title\": str, \"bullets\": [str]}]\n"
    )

    user_prompt = f"Create a {num_slides}-slide professional presentation for: {prompt}"\

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
            raise ValueError("Invalid JSON")

        return plan[:num_slides]

    except Exception:
        logger.warning("LLM failed, using semantic fallback plan")

        fallback = []
        for i in range(num_slides):
            fallback.append(
                {
                    "title": f"Slide {i + 1}",
                    "bullets": references_text[i % len(references_text)].split(". ")[:4]
                    if references_text else ["Auto-generated content"],
                }
            )
        return fallback


# ------------------------------------------------------------
# IMAGE GENERATION
# ------------------------------------------------------------

def generate_visual_image(prompt: str):
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
# PPT BUILDER (TEMPLATE-AWARE)
# ------------------------------------------------------------

def build_ppt(slides, template_style=None):

    # ✅ If user selected a template from blob → use it as base
    if template_style and template_style != "Plain (Default)":
        try:
            prs = Presentation(template_style)
        except Exception:
            prs = Presentation()
    else:
        prs = Presentation()

    for sp in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        slide.shapes.title.text = sp.get("title", "")

        body = slide.placeholders[1]
        tf = body.text_frame
        tf.clear()

        for b in sp.get("bullets", []):
            p = tf.add_paragraph()
            p.text = b
            p.font.size = Pt(20)

        body.top = slide.shapes.title.top + slide.shapes.title.height + Inches(0.3)

        img_path = sp.get("image_path")

        if not img_path:
            body.left = Inches(0.5)
            body.width = prs.slide_width - Inches(1.0)
            continue

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

    out_path = os.path.join(tempfile.gettempdir(), f"generated_{uuid.uuid4().hex[:8]}.pptx")
    prs.save(out_path)
    return out_path


# ------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------

def generate_presentation(
    prompt: str,
    requested_num_slides=5,
    template_style=None,
    force_images=False,
):

    refs = semantic_search(prompt, top_k=5) or []

    if not refs:
        msg = "No matching content found in uploaded PPTs."
        return None, {"error": True, "message": msg}

    reference_text = [(r.get("text") or "")[:500] for r in refs]

    plan = call_llm_plan(prompt, reference_text, requested_num_slides)

    slides = []

    for sp in plan:
        img_path = None
        if force_images:
            img_path = generate_visual_image(sp.get("title"))

        slides.append(
            {
                "title": sp.get("title", "Untitled"),
                "bullets": sp.get("bullets", []),
                "image_path": img_path,
            }
        )

    out_path = build_ppt(slides, template_style)

    fname = f"generated_{uuid.uuid4().hex[:8]}.pptx"
    upload_ppt_to_blob(out_path, fname)

    log = {
        "timestamp": now_ts(),
        "prompt": prompt,
        "slides_generated": len(slides),
        "ppt_file": fname,
        "error": False,
        "template_style": template_style,
        "force_images": force_images,
    }

    return out_path, log
