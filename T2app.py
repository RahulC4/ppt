import streamlit as st
from generate_ppt import generate_presentation
from utils import logger

# -----------------------------
# STREAMLIT PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="AI PowerPoint Generator",
    layout="wide",
    page_icon="📊"
)

# -----------------------------
# HEADER
# -----------------------------
st.title("📊 AI PowerPoint Generator")
st.write(
    "Generate high-quality PowerPoint presentations using Azure OpenAI + your enterprise slide dataset."
)

st.markdown("---")

# -----------------------------
# USER INPUTS (MINIMAL UI)
# -----------------------------
prompt = st.text_area(
    "Enter your presentation prompt:",
    placeholder="Example: Create a 5-slide presentation about AI in Healthcare. Include images on every slide.",
    height=150
)

num_slides = st.number_input(
    "Number of Slides",
    min_value=1,
    max_value=25,
    value=5,
    step=1
)

theme = st.selectbox(
    "Theme (optional):",
    ["None", "Dark", "Light", "Modern", "Corporate", "Minimal"]
)
theme = None if theme == "None" else theme  # convert dropdown value to Python None

st.markdown("---")

# -----------------------------
# PREVIEW BUTTON
# -----------------------------
if st.button("🔍 Preview Presentation Plan"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
    else:
        with st.spinner("Generating presentation preview..."):
            try:
                out_path, log = generate_presentation(
                    prompt=prompt,
                    requested_num_slides=num_slides,
                    theme=theme,
                    style="Auto",       # Design Style removed from UI
                    tag_filters=None    # Tags removed from UI
                )
                st.success("Preview generated successfully!")
                st.json(log)

                # Store for download
                st.session_state["ppt_path"] = out_path

            except Exception as e:
                logger.exception("Error generating preview")
                st.error(f"Failed to generate presentation: {e}")

# -----------------------------
# DOWNLOAD BUTTON
# -----------------------------
if "ppt_path" in st.session_state:
    st.markdown("---")
    st.success("Presentation file is ready!")

    with open(st.session_state["ppt_path"], "rb") as f:
        st.download_button(
            label="📥 Download Generated PPT",
            data=f,
            file_name="generated_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
