import streamlit as st
from generate_ppt import generate_presentation
from utils import logger

st.set_page_config(page_title="AI PowerPoint Generator", layout="wide")

st.title("📊 AI PowerPoint Generator")
st.write("Generate high-quality PPT decks using Azure OpenAI + your enterprise dataset.")

# --- User Input ---
prompt = st.text_area(
    "Enter your presentation prompt:",
    placeholder="e.g., Create a presentation deck about AI in healthcare. Include images.",
    height=150
)

num_slides = st.number_input(
    "Number of slides:",
    min_value=1,
    max_value=20,
    value=5
)

theme = st.selectbox("Theme (optional):", ["Auto", "Dark", "Light", "Corporate", "Modern"], index=0)


# ===============================
# PREVIEW BUTTON
# ===============================
if st.button("🔍 Preview Presentation Plan"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
    else:
        with st.spinner("Generating preview using Azure OpenAI..."):
            try:
                out_path, log = generate_presentation(
                    prompt=prompt,
                    requested_num_slides=num_slides,
                    theme=None if theme == "Auto" else theme.lower(),
                )

                st.success("Preview generated!")
                st.subheader("📄 Generated Slide Plan (from GPT)")
                st.json(log)

            except Exception as e:
                logger.exception("Preview generation failed")
                st.error(f"Error: {e}")


# ===============================
# FULL GENERATION BUTTON
# ===============================
if st.button("📥 Generate & Download PPT"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
    else:
        with st.spinner("Generating final PPT..."):
            try:
                out_path, log = generate_presentation(
                    prompt=prompt,
                    requested_num_slides=num_slides,
                    theme=None if theme == "Auto" else theme.lower(),
                )

                st.success("PPT Generated Successfully!")

                # Read the file and allow download
                with open(out_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Download PPT",
                        data=f,
                        file_name="generated_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

                st.subheader("Generation Log")
                st.json(log)

            except Exception as e:
                logger.exception("PPT generation failed")
                st.error(f"Error: {e}")
