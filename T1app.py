import streamlit as st
import base64
import os
from generate_ppt import generate_presentation, call_llm_plan
from search_utils import semantic_search
from utils import logger, get_env, safe_json_load, text_client
from pptx import Presentation

st.set_page_config(page_title="AI PPT Generator", layout="wide")

# ---------------------------------------------
#  STREAMLIT UI
# ---------------------------------------------

st.title("📊 AI PowerPoint Generator")
st.write("Generate high-quality PPT decks using Azure OpenAI & your enterprise dataset.")

prompt = st.text_area(
    "Enter your presentation prompt:",
    placeholder="e.g., Create a 5-slide corporate deck about AI in healthcare..."
)

col1, col2, col3 = st.columns(3)
with col1:
    style = st.selectbox("Design Style", ["Auto", "Design"])
with col2:
    num_slides = st.number_input("Number of Slides", min_value=1, max_value=30, value=5)
with col3:
    theme = st.selectbox("Theme", ["Auto", "Corporate", "Modern", "Minimal", "Colorful", "Dark", "Light"])

tag_filters = st.multiselect(
    "Filter by Tags (optional):",
    ["Design", "Migration", "Claims", "Membership", "Test", "General"]
)

st.divider()

# ---------------------------------------------
#  PREVIEW BUTTON
# ---------------------------------------------

if st.button("🔍 Preview Presentation Plan"):
    if not prompt.strip():
        st.error("Please enter a prompt before previewing.")
        st.stop()

    with st.spinner("Retrieving relevant slides & generating preview..."):

        # Get Chroma references
        refs = semantic_search(prompt, top_k=5, tags=tag_filters)

        reference_text = [r.get("text", "")[:400] for r in refs if r.get("text")]

        # Dummy design context (we skip JSON reading for preview)
        design_context = []

        plan = call_llm_plan(
            prompt=prompt,
            style=style,
            design_context=design_context,
            references_text=reference_text,
            num_slides=num_slides,
            theme=None if theme == "Auto" else theme
        )

    st.subheader("📝 Slide Plan Preview")
    for i, slide in enumerate(plan, start=1):
        st.markdown(f"### **Slide {i}: {slide.get('title','Untitled')}**")

        bullets = slide.get("bullets", [])
        if bullets:
            st.markdown("\n".join([f"- {b}" for b in bullets]))

        if slide.get("visual_required"):
            st.info(f"📌 Visual Required → Prompt: `{slide.get('visual_prompt')}`")

    st.success("Preview generated. When ready, click **Generate PPT** below.")

st.divider()

# ---------------------------------------------
#  GENERATE PPT BUTTON
# ---------------------------------------------

if st.button("🎯 Generate PPT File"):
    if not prompt.strip():
        st.error("Please enter a prompt before generating.")
        st.stop()

    with st.spinner("Generating PPT... this may take ~10–20 seconds..."):

        ppt_path, log_data = generate_presentation(
            prompt=prompt,
            style=style,
            requested_num_slides=num_slides,
            theme=None if theme == "Auto" else theme,
            tag_filters=tag_filters
        )

    st.success("PPT Generated Successfully!")

    # Read PPT file
    with open(ppt_path, "rb") as f:
        ppt_bytes = f.read()

    st.download_button(
        label="⬇️ Download PPT",
        data=ppt_bytes,
        file_name=os.path.basename(ppt_path),
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    st.json(log_data)
