import streamlit as st
from generate_ppt import generate_presentation, call_llm_plan
from search_utils import semantic_search
from utils import logger, safe_json_load

st.set_page_config(page_title="AI PowerPoint Generator", layout="wide")

st.title("📊 AI PowerPoint Generator")
st.write("Generate professional PPTs using Azure OpenAI + your slide dataset.")

st.markdown("---")

# -----------------------------
# USER INPUTS
# -----------------------------
prompt = st.text_area(
    "Enter your presentation prompt:",
    placeholder="Example: Create a 5-slide corporate presentation about AI in healthcare. Include images.",
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
theme = None if theme == "None" else theme

st.markdown("---")


# ============================================================
# FAST PREVIEW (LLM PLAN ONLY)
# ============================================================
if st.button("🔍 Preview Presentation Plan"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
        st.stop()

    with st.spinner("Generating preview (fast mode)..."):
        try:
            # Step 1: Retrieve references for context
            refs = semantic_search(prompt, top_k=5)

            ref_text = [(r.get("text") or "")[:400] for r in refs if r.get("text")]

            # Step 2: Preview ONLY the plan from GPT
            plan = call_llm_plan(
                prompt=prompt,
                style="Auto",
                design_context=[],        # design not needed for preview
                references_text=ref_text,
                num_slides=num_slides,
                theme=theme
            )

            st.success("Preview generated!")

            st.subheader("📝 Slide Plan (LLM Output)")

            for i, slide in enumerate(plan, start=1):
                st.markdown(f"### Slide {i}: {slide.get('title', 'Untitled')}")

                bullets = slide.get("bullets", [])
                if bullets:
                    st.write("\n".join([f"- {b}" for b in bullets]))

                if slide.get("visual_required"):
                    st.info(f"📌 Image Required → {slide.get('visual_prompt')}")

            # Save plan to session
            st.session_state["preview_plan"] = plan

        except Exception as e:
            logger.exception("Preview failed")
            st.error(f"Error generating preview: {e}")


st.markdown("---")


# ============================================================
# GENERATE FINAL PPT
# ============================================================
if st.button("📥 Generate & Download PPT"):
    if not prompt.strip():
        st.error("Please enter a prompt.")
        st.stop()

    with st.spinner("Generating full PowerPoint (this may take 10–25 seconds)..."):
        try:
            ppt_path, log = generate_presentation(
                prompt=prompt,
                requested_num_slides=num_slides,
                theme=theme,
                style="Auto",
                tag_filters=None
            )

            st.success("PPT Generated Successfully!")

            with open(ppt_path, "rb") as f:
                st.download_button(
                    label="⬇️ Download PPT File",
                    data=f,
                    file_name="generated_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            st.subheader("Generation Log")
            st.json(log)

        except Exception as e:
            logger.exception("PPT generation failed")
            st.error(f"Failed to generate PPT: {e}")
