import os
import streamlit as st

from generate_ppt import generate_presentation, call_llm_plan
from search_utils import semantic_search
from utils import logger, get_env
from azure_blob_utils import (
    upload_source_ppt_to_blob,
    list_source_ppt_blobs,
    delete_source_ppt_from_blob,
)
from ingestion_chroma import (
    process_blob as ingest_process_blob,
    delete_ppt_from_chroma,
)

st.set_page_config(page_title="AI PPT Generator", layout="wide", page_icon="📊")

# Session state for tracking generated PPTs this session
if "generated_ppts" not in st.session_state:
    st.session_state["generated_ppts"] = []

st.title("📊 AI PowerPoint Generator")
st.write(
    "Generate PPT decks based on your existing slides. "
    "Upload sample PPTs, and let the model enhance – not invent from scratch."
)

SIMILARITY_THRESHOLD = float(get_env("SIMILARITY_THRESHOLD", "1.1"))

st.markdown("---")

# ============================================================
# 📥 SIDEBAR: UPLOAD + MANAGE KNOWLEDGE BASE PPTs
# ============================================================
with st.sidebar:
    st.subheader("1️⃣ Upload Sample PPT Files (Knowledge Base)")

    uploaded_files = st.file_uploader(
        "Upload .pptx files to use as references:",
        type=["pptx"],
        accept_multiple_files=True,
        key="kb_uploader",
    )

    if st.button("📥 Add to Knowledge Base", key="add_kb_btn") and uploaded_files:
        with st.spinner("Uploading & indexing PPTs..."):
            for upl in uploaded_files:
                try:
                    bytes_data = upl.read()
                    blob_name = upl.name

                    # Upload to blob
                    upload_source_ppt_to_blob(bytes_data, blob_name)

                    # Ingest into Chroma (text + embeddings only)
                    ingest_process_blob(blob_name)

                    st.success(f"✅ Processed & indexed: {blob_name}")

                except Exception as e:
                    logger.exception(f"Failed to process {upl.name}")
                    st.error(f"❌ Error processing {upl.name}: {e}")

    st.markdown("---")
    st.subheader("📂 Available Knowledge Base PPTs")

    try:
        kb_files = list_source_ppt_blobs()
    except Exception as e:
        logger.exception("Failed to list knowledge base PPTs")
        st.error(f"❌ Error listing PPTs: {e}")
        kb_files = []

    if kb_files:
        for blob_name in kb_files:
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.caption(blob_name)
            with col_b:
                if st.button("🗑️ Remove", key=f"delete_{blob_name}"):
                    try:
                        # 1️⃣ Delete PPT from Blob
                        delete_source_ppt_from_blob(blob_name)
                        # 2️⃣ Delete all related slide indexes from Chroma
                        delete_ppt_from_chroma(blob_name)

                        st.success(f"🗑️ Removed: {blob_name}")
                        # Refresh UI so removed file disappears from the list
                        st.experimental_rerun()
                    except Exception as e:
                        logger.exception(f"Failed to delete PPT {blob_name}")
                        st.error(f"❌ Error deleting {blob_name}: {e}")
    else:
        st.caption("No PPTs in knowledge base yet.")

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
    # Renamed label to 'Themes' and restricted options
    template_style = st.selectbox(
        "Themes",
        ["Auto", "Corporate"],
        index=0,
    )

with col3:
    image_required = st.checkbox("🖼 Generate images for all slides")

st.markdown("---")

# ============================================================
# 🔍 PREVIEW (TEXT-BASED, NO JSON)
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
                        "for this prompt. Please rephrase using topics related to your "
                        "sample decks or upload a new PPT."
                    )
                else:
                    ref_text = [
                        (r.get("text") or "")[:400]
                        for r in refs
                        if r.get("text")
                    ]

                    plan = call_llm_plan(
                        prompt=prompt,
                        references_text=ref_text,
                        num_slides=num_slides,
                    )

                    if not plan:
                        st.warning(
                            "⚠️ Not enough relevant content to generate this many slides. "
                            "Try reducing slide count or rephrasing."
                        )
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
                    requested_num_slides=num_slides,
                    template_style=template_style,
                    image_required=image_required,
                )

                if log.get("error"):
                    st.warning(f"⚠️ {log.get('message')}")
                else:
                    st.success("✅ PPT Generated Successfully!")

                    # Derive a display-friendly name from path
                    display_name = os.path.basename(ppt_path) if ppt_path else "generated_presentation.pptx"

                    # Track in current session list
                    st.session_state["generated_ppts"].append(
                        {"path": ppt_path, "name": display_name}
                    )

                    with open(ppt_path, "rb") as f:
                        st.download_button(
                            label="⬇️ Download PPT",
                            data=f,
                            file_name=display_name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )

                    st.subheader("📄 Generation Log")
                    st.json(log)

            except Exception as e:
                logger.exception("PPT generation failed")
                st.error(f"Failed to generate PPT: {e}")

st.markdown("---")

# ============================================================
# 📂 THIS SESSION'S GENERATED PPTs
# ============================================================
st.subheader("📂 This Session's Generated PPTs")

if not st.session_state["generated_ppts"]:
    st.caption("No PPTs generated yet in this session.")
else:
    for idx, item in enumerate(st.session_state["generated_ppts"]):
        col1, col2 = st.columns([4, 2])
        with col1:
            st.write(f"{idx + 1}. {item['name']}")
        with col2:
            try:
                with open(item["path"], "rb") as f:
                    st.download_button(
                        label="⬇️ Download again",
                        data=f,
                        file_name=item["name"],
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        key=f"dl_session_{idx}",
                    )
            except Exception:
                st.caption("File no longer available on disk.")
