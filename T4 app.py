import streamlit as st
import requests
import base64

st.set_page_config(page_title="AI PPT Generator", layout="wide")

BACKEND_URL = "http://localhost:8000"  # change if different


# -------------------------------
# SESSION STATE
# -------------------------------
if "selected_template" not in st.session_state:
    st.session_state.selected_template = "corporate"

if "preview_data" not in st.session_state:
    st.session_state.preview_data = None


# -------------------------------
# HEADER
# -------------------------------
st.markdown("## 🧠 AI PowerPoint Generator")
st.caption("Generate PPT using your enterprise knowledge base and fixed templates")


# -------------------------------
# KNOWLEDGE BASE UPLOAD
# -------------------------------
st.markdown("### 1️⃣ Upload Sample PPTs (Knowledge Base)")

uploaded_files = st.file_uploader(
    "Upload one or more PPTX files",
    type=["pptx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("📤 Add to Knowledge Base"):
        files = [("files", (f.name, f, "application/vnd.openxmlformats-officedocument.presentationml.presentation"))
                 for f in uploaded_files]

        res = requests.post(f"{BACKEND_URL}/upload-ppt", files=files)

        if res.status_code == 200:
            st.success("✅ PPTs uploaded and indexed successfully!")
        else:
            st.error("❌ Upload failed")


st.divider()


# -------------------------------
# CREATE PPT
# -------------------------------
st.markdown("### 2️⃣ Create New Presentation")

prompt = st.text_area("Enter your presentation prompt")

slides = st.number_input("Number of Slides", min_value=1, max_value=20, value=5)


# -------------------------------
# TEXT LENGTH SELECTION
# -------------------------------
st.markdown("#### 📝 Amount of text per slide")

text_density = st.radio(
    "",
    ["Minimal", "Concise", "Detailed", "Extensive"],
    horizontal=True,
    index=1
)


# -------------------------------
# TEMPLATE SELECTION (GAMMA STYLE)
# -------------------------------
st.markdown("#### 🎨 Select a Template")

template_cols = st.columns(4)

TEMPLATES = {
    "corporate": "https://cdn.gamma.app/template1.png",
    "modern": "https://cdn.gamma.app/template2.png",
    "minimal": "https://cdn.gamma.app/template3.png",
    "dark": "https://cdn.gamma.app/template4.png"
}

for idx, (template_name, img_url) in enumerate(TEMPLATES.items()):
    with template_cols[idx]:
        st.image(img_url, use_container_width=True)
        if st.button(template_name.capitalize()):
            st.session_state.selected_template = template_name

st.success(f"✅ Selected Template: **{st.session_state.selected_template.upper()}**")


st.divider()


# -------------------------------
# PREVIEW SLIDE PLAN
# -------------------------------
st.markdown("### 3️⃣ Preview Slide Plan")

if st.button("🔍 Preview Slide Plan"):
    payload = {
        "prompt": prompt,
        "num_slides": slides,
        "density": text_density,
        "template_style": st.session_state.selected_template
    }

    res = requests.post(f"{BACKEND_URL}/preview", json=payload)

    if res.status_code == 200:
        st.session_state.preview_data = res.json()
        st.success("✅ Preview Generated")
    else:
        st.error("❌ Error while previewing")


if st.session_state.preview_data:
    st.markdown("#### 📄 Slide Plan Preview")

    for idx, slide in enumerate(st.session_state.preview_data["slides"]):
        with st.expander(f"Slide {idx+1}: {slide['title']}"):
            for bullet in slide["bullets"]:
                st.write(f"- {bullet}")


st.divider()


# -------------------------------
# GENERATE & DOWNLOAD PPT
# -------------------------------
st.markdown("### 4️⃣ Generate & Download PPT")

if st.button("🚀 Generate & Download PPT"):
    payload = {
        "prompt": prompt,
        "num_slides": slides,
        "density": text_density,
        "template_style": st.session_state.selected_template
    }

    res = requests.post(f"{BACKEND_URL}/generate", json=payload)

    if res.status_code == 200:
        data = res.json()
        ppt_base64 = data["ppt_base64"]

        ppt_bytes = base64.b64decode(ppt_base64)

        st.download_button(
            label="⬇️ Download PPT",
            data=ppt_bytes,
            file_name="generated_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        st.success("✅ PPT Generated Successfully")

    else:
        st.error("❌ Failed to generate PPT")
