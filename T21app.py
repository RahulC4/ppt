import streamlit as st

from generate_ppt import generate_presentation, call_llm_plan
from search_utils import semantic_search
from utils import logger, get_env
from azure_blob_utils import upload_source_ppt_to_blob
from ingestion_chroma import process_blob as ingest_process_blob

st.set_page_config(page_title="AI PPT Generator", layout="wide", page_icon="📊")

st.title("📊 AI PowerPoint Generator")
st.write(
    "Generate PPT decks based on your existing slides. "
    "Upload sample PPTs, and let the model enhance – not invent from scratch."
)

SIMILARITY_THRESHOLD = float(get_env("SIMILARITY_THRESHOLD", "1.1"))

st.markdown("---")

# ============================================================
# 1️⃣ UPLOAD SAMPLE PPTs
# ============================================================
st.subheader("1️⃣ Upload Sample PPT Files (Knowledge Base)")

uploaded_files = st.file_uploader(
    "Upload .pptx files to use as references:",
    type=["pptx"],
    accept_multiple_files=True,
)

if st.button("📥 Add to Knowledge Base") and uploaded_files:
    with st.spinner("Uploading & indexing PPTs..."):
        for upl in uploaded_files:
            try:
                bytes_data = upl.read()
                blob_name = upl.name

                upload_source_ppt_to_blob(bytes_data, blob_name)
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
    placeholder="Example: Create a 5-slide presentation about AI in healthcare using my uploaded decks.",
    height=150,
)

col1, col2, col3 = st.columns(3)

with col1:
    num_slides = st.number_input(
        "Number of Slides",
        min_value=1,
        max_value=25,
        value=5,
        step=1,
    )

with col2:
    template_style = st.selectbox(
        "Visual Template",
        ["Auto", "Plant", "Dark", "Minimal", "Corporate"],
        index=0,
    )

with col3:
    image_required = st.checkbox(
        "Generate images for all slides",
        value=True,
    )

st.markdown("---")

# ============================================================
# 🔍 PREVIEW (TEXT PLAN ONLY)
# ============================================================
if st.button("🔍 Preview Slide Plan"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
    else:
        with st.spinner("Searching knowledge base & generating preview..."):
            try:
                raw_refs = semantic_search(prompt, top_k=5)
                refs = [
                    r for r in (raw_refs or [])
                    if r.get("score") is None or r["score"] <= SIMILARITY_THRESHOLD
                ]

                if not refs:
                    st.warning(
                        "⚠️ I couldn’t find any relevant content in your uploaded PPTs "
                        "for this prompt. Please rephrase or upload a new PPT."
                    )
                else:
                    ref_text = [
                        (r.get("text") or "")[:400]
                        for r in refs
                        if r.get("text")
                    ]

                    plan = call_llm_plan(
                        prompt=prompt,
                        style="Auto",
                        references_text=ref_text,
                        num_slides=num_slides,
                    )

                    if not plan:
                        st.warning("Invalid preview plan.")
                    else:
                        st.subheader("📝 Slide Plan Preview")
                        for i, slide in enumerate(plan, start=1):
                            st.markdown(f"### Slide {i}: {slide.get('title','Untitled')}")
                            bullets = slide.get("bullets", [])
                            if bullets:
                                st.markdown("\n".join(f"- {b}" for b in bullets))

            except Exception as e:
                logger.exception("Preview failed")
                st.error(f"Error while generating preview: {e}")

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
                    style="Auto",
                    requested_num_slides=num_slides,
                    theme=None,
                    tag_filters=None,
                    template_style=template_style,
                    image_required=image_required,   # ✅ NEW
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
