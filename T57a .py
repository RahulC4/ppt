# ============================================================
# 🎯 GENERATE PPT
# ============================================================
if generated_clicked:
    if not prompt.strip():
        st.error("Please enter a prompt.")
    else:
        with st.spinner("Generating final PowerPoint , usually takes less than 60s..."):
            try:
                # ✅ Generate using correct engine
                if template_style.lower() == "corporate":
                    ppt_path, log = generate_presentation(
                        prompt=prompt,
                        requested_num_slides=num_slides,
                        template_style=template_style,
                        image_required=image_required,
                    )
                else:
                    ppt_path, log = generate_presentation_auto(
                        prompt=prompt,
                        requested_num_slides=num_slides,
                        template_style=template_style,
                        image_required=image_required,
                    )

                # ✅ Unified success + error handling (for BOTH themes)
                if log.get("error"):
                    st.warning(f"⚠️ {log.get('message')}")
                else:
                    st.success("✅ PPT Generated Successfully!")

                    display_name = os.path.basename(ppt_path) if ppt_path else "generated_presentation.pptx"

                    # ✅ ALWAYS append (this was broken before)
                    st.session_state["generated_ppts"].append(
                        {"path": ppt_path, "name": display_name}
                    )

                    # ✅ Store last log for later display (optional)
                    st.session_state["last_generation_log"] = log

            except Exception as e:
                logger.exception("PPT generation failed")
                st.error(f"Failed to generate PPT: {e}")
